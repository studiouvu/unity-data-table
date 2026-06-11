using System.Globalization;
using System.IO.Compression;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace DTable
{
    /// <summary>
    /// export 컬럼의 외부 참조 수식이 오래된 캐시 값으로 변환되는 것을 막기 위한 예외.
    /// </summary>
    internal sealed class StaleExternalCacheException : Exception
    {
        public StaleExternalCacheException(string message) : base(message) { }
    }

    /// <summary>
    /// 수식의 계산 값이 파일에 저장되어 있지 않아 (Excel 외 도구로 생성된 파일)
    /// 변환 결과가 비게 되는 것을 막기 위한 예외.
    /// </summary>
    internal sealed class MissingFormulaCacheException : Exception
    {
        public MissingFormulaCacheException(string message) : base(message) { }
    }

    /// <summary>
    /// 수식 에러 값 (#N/A 등). Interop Value2와 동일한 에러 코드 숫자로 export된다.
    /// </summary>
    internal readonly struct CellError
    {
        public readonly string Code;
        public CellError(string code) => Code = code;
    }

    internal sealed class XlsxSheet
    {
        public string Name;
        public int RowCount;
        public int ColumnCount;
        public readonly Dictionary<(int Row, int Column), object> Cells = new Dictionary<(int Row, int Column), object>();
        // 외부 워크북([1]DataText! 형태)을 참조하는 수식 셀
        public readonly HashSet<(int Row, int Column)> ExternalFormulaCells = new HashSet<(int Row, int Column)>();
        // 수식은 있는데 계산 값이 저장되어 있지 않은 셀
        public readonly HashSet<(int Row, int Column)> UncachedFormulaCells = new HashSet<(int Row, int Column)>();

        public object GetValue(int row, int column)
        {
            return Cells.TryGetValue((row, column), out var value) ? value : null;
        }
    }

    /// <summary>
    /// xlsx(Open XML)를 Excel 없이 직접 읽는다.
    /// Interop의 UsedRange.Value2와 동일한 값이 나오도록 변환한다:
    /// 날짜/시간은 OLE 시리얼 숫자 그대로, 문자열은 rich text run을 전부 이어붙인 원문.
    /// </summary>
    internal sealed class XlsxFile
    {
        private static readonly XNamespace Main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        private static readonly XNamespace DocRel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        private static readonly XNamespace PkgRel = "http://schemas.openxmlformats.org/package/2006/relationships";

        // 수식 텍스트의 [1], [2] 같은 외부 워크북 인덱스
        private static readonly Regex ExternalRefPattern = new Regex(@"\[\d+\]", RegexOptions.Compiled);
        // OOXML 문자 이스케이프 (_x000D_ 등)
        private static readonly Regex EscapedCharPattern = new Regex("_x([0-9A-Fa-f]{4})_", RegexOptions.Compiled);

        public readonly List<XlsxSheet> Sheets = new List<XlsxSheet>();

        private readonly List<ExternalLink> externalLinks = new List<ExternalLink>();

        private sealed class ExternalLink
        {
            public string TargetFileName;
            public readonly Dictionary<string, Dictionary<(int Row, int Column), string>> CachedSheets
                = new Dictionary<string, Dictionary<(int Row, int Column), string>>();
        }

        public static XlsxFile Load(string fileExcel)
        {
            var file = new XlsxFile();

            // Excel에서 열려 있는 파일도 읽을 수 있도록 FileShare.ReadWrite로 연다 (Interop ReadOnly와 동일한 동작)
            using var stream = new FileStream(fileExcel, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var archive = new ZipArchive(stream, ZipArchiveMode.Read);

            var sharedStrings = LoadSharedStrings(archive);

            foreach (var (sheetName, entryPath) in GetSheetEntryPaths(archive))
            {
                var sheet = new XlsxSheet { Name = sheetName };
                file.Sheets.Add(sheet);

                // #시트는 변환되지 않으므로 셀을 읽지 않는다
                if (sheetName.Contains('#'))
                    continue;

                var entry = archive.GetEntry(entryPath);
                if (entry != null)
                    LoadSheetCells(entry, sharedStrings, sheet);
            }

            foreach (var linkEntry in archive.Entries)
            {
                if (!linkEntry.FullName.StartsWith("xl/externalLinks/externalLink") || !linkEntry.FullName.EndsWith(".xml"))
                    continue;

                var link = LoadExternalLink(archive, linkEntry);
                if (link != null)
                    file.externalLinks.Add(link);
            }

            return file;
        }

        /// <summary>
        /// 모든 외부 참조 캐시를 원본 파일(같은 폴더)의 현재 값과 비교한다.
        /// 불일치가 있으면 StaleExternalCacheException을 던져 변환을 실패시킨다.
        /// </summary>
        public void ValidateCacheFreshness(string fileExcel, List<string> exportedExternalCells)
        {
            var problems = new List<string>();
            var folder = Path.GetDirectoryName(Path.GetFullPath(fileExcel)) ?? "";

            foreach (var link in externalLinks)
            {
                var targetPath = Path.Combine(folder, link.TargetFileName);
                if (!File.Exists(targetPath))
                {
                    problems.Add($"원본 파일을 찾을 수 없어 캐시를 검증할 수 없습니다: {link.TargetFileName}");
                    continue;
                }

                var targetFile = Load(targetPath);

                foreach (var pair in link.CachedSheets)
                {
                    var sheetName = pair.Key;
                    var cachedCells = pair.Value;

                    var targetSheet = targetFile.Sheets.FirstOrDefault(s => s.Name == sheetName);
                    if (targetSheet == null)
                    {
                        problems.Add($"{link.TargetFileName}에서 '{sheetName}' 시트를 찾을 수 없습니다.");
                        continue;
                    }

                    var liveCells = new Dictionary<(int Row, int Column), string>();
                    foreach (var cell in targetSheet.Cells)
                    {
                        var token = MakeLiveToken(cell.Value);
                        if (token != null)
                            liveCells[cell.Key] = token;
                    }

                    var cachedColumns = new HashSet<int>(cachedCells.Keys.Select(c => c.Column));
                    var mismatches = new List<string>();

                    // 캐시에 있는 값이 원본에서 바뀌었거나 사라진 경우
                    foreach (var cell in cachedCells)
                    {
                        liveCells.TryGetValue(cell.Key, out var liveToken);
                        if (liveToken != cell.Value)
                            mismatches.Add(CellReference(cell.Key.Row, cell.Key.Column));
                    }

                    // 캐시된 컬럼 범위에 원본이 새 값을 추가한 경우 (예: 새 행 추가)
                    foreach (var cellRef in liveCells.Keys)
                    {
                        if (cachedColumns.Contains(cellRef.Column) && !cachedCells.ContainsKey(cellRef))
                            mismatches.Add(CellReference(cellRef.Row, cellRef.Column));
                    }

                    if (mismatches.Count > 0)
                    {
                        mismatches.Sort();
                        problems.Add(
                            $"{link.TargetFileName} '{sheetName}' 시트가 캐시와 {mismatches.Count}건 불일치 (예: {string.Join(", ", mismatches.Take(5))})");
                    }
                }
            }

            if (problems.Count == 0)
                return;

            throw new StaleExternalCacheException(
                "외부 참조 캐시가 최신이 아닙니다. 이 파일을 Excel에서 열어 링크를 갱신하고 저장한 뒤 다시 변환하세요.\n" +
                $"  - 외부 참조 수식이 export되는 셀: {Summarize(exportedExternalCells)}\n" +
                string.Join("\n", problems.Select(p => $"  - {p}")));
        }

        public static string Summarize(List<string> cells)
        {
            var summary = string.Join(", ", cells.Take(5));
            if (cells.Count > 5)
                summary += $" 외 {cells.Count - 5}건";
            return summary;
        }

        private static void LoadSheetCells(ZipArchiveEntry sheetEntry, List<string> sharedStrings, XlsxSheet sheet)
        {
            var doc = LoadXml(sheetEntry);
            var sheetData = doc.Root?.Element(Main + "sheetData");
            if (sheetData == null)
                return;

            // 공유 수식(si)은 첫 셀에만 수식 텍스트가 있으므로 si별 외부 참조 여부를 기억해 전파
            var sharedFormulaExternal = new Dictionary<string, bool>();

            foreach (var cell in sheetData.Elements(Main + "row").SelectMany(r => r.Elements(Main + "c")))
            {
                if (!TryParseCellReference((string)cell.Attribute("r"), out var rowCol))
                    continue;

                if (rowCol.Row > sheet.RowCount)
                    sheet.RowCount = rowCol.Row;
                if (rowCol.Column > sheet.ColumnCount)
                    sheet.ColumnCount = rowCol.Column;

                var formula = cell.Element(Main + "f");
                if (formula != null)
                    ClassifyFormula(formula, cell, rowCol, sharedFormulaExternal, sheet);

                var value = ReadCellValue(cell, sharedStrings, formula != null);
                if (value != null)
                    sheet.Cells[rowCol] = value;
            }
        }

        private static void ClassifyFormula(
            XElement formula,
            XElement cell,
            (int Row, int Column) rowCol,
            Dictionary<string, bool> sharedFormulaExternal,
            XlsxSheet sheet)
        {
            var text = formula.Value;
            var sharedIndex = (string)formula.Attribute("si");

            bool isExternal;
            if (!string.IsNullOrEmpty(text))
            {
                isExternal = ExternalRefPattern.IsMatch(text);
                if (sharedIndex != null)
                    sharedFormulaExternal[sharedIndex] = isExternal;
            }
            else if (sharedIndex != null)
            {
                sharedFormulaExternal.TryGetValue(sharedIndex, out isExternal);
            }
            else
            {
                isExternal = false;
            }

            if (isExternal)
                sheet.ExternalFormulaCells.Add(rowCol);

            // 계산 값이 저장되어 있는지 검사:
            // - <v>가 아예 없으면 캐시 없음
            // - <v/>가 비어 있는데 타입이 숫자(t 없음/"n")면 캐시 없음 (숫자는 빈 값이 될 수 없다)
            //   t="str"의 빈 <v/>는 정상적인 빈 문자열 결과("")이므로 제외
            var value = cell.Element(Main + "v");
            var type = (string)cell.Attribute("t");
            if (value == null || (value.Value.Length == 0 && (type == null || type == "n")))
                sheet.UncachedFormulaCells.Add(rowCol);
        }

        private static object ReadCellValue(XElement cell, List<string> sharedStrings, bool hasFormula)
        {
            var type = (string)cell.Attribute("t") ?? "n";

            if (type == "inlineStr")
            {
                var inline = cell.Element(Main + "is");
                return inline != null ? GetRichText(inline) : null;
            }

            var value = cell.Element(Main + "v")?.Value;
            if (value == null)
                return null;

            switch (type)
            {
                case "s":
                    return int.TryParse(value, out var index) && index >= 0 && index < sharedStrings.Count
                        ? sharedStrings[index]
                        : null;
                case "str":
                    return UnescapeText(value);
                case "b":
                    return value == "1";
                case "e":
                    return new CellError(value);
                default: // "n" = 숫자 (날짜/시간도 시리얼 숫자 그대로)
                    return double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var number)
                        ? number
                        : (object)value;
            }
        }

        private static IEnumerable<(string SheetName, string EntryPath)> GetSheetEntryPaths(ZipArchive archive)
        {
            var workbook = LoadXml(archive, "xl/workbook.xml");
            var rels = LoadXml(archive, "xl/_rels/workbook.xml.rels");
            if (workbook?.Root == null || rels?.Root == null)
                yield break;

            var targets = new Dictionary<string, string>();
            foreach (var rel in rels.Root.Elements(PkgRel + "Relationship"))
            {
                var id = (string)rel.Attribute("Id");
                var target = (string)rel.Attribute("Target");
                if (id != null && target != null)
                    targets[id] = target;
            }

            var sheets = workbook.Root.Element(Main + "sheets");
            if (sheets == null)
                yield break;

            foreach (var sheet in sheets.Elements(Main + "sheet"))
            {
                var name = (string)sheet.Attribute("name");
                var rid = (string)sheet.Attribute(DocRel + "id");
                if (name == null || rid == null || !targets.TryGetValue(rid, out var target))
                    continue;

                var entryPath = target.StartsWith("/") ? target.TrimStart('/') : "xl/" + target;
                yield return (name, entryPath);
            }
        }

        private static ExternalLink LoadExternalLink(ZipArchive archive, ZipArchiveEntry linkEntry)
        {
            var doc = LoadXml(linkEntry);

            // DDE/OLE 링크 등 워크북 참조가 아닌 외부 링크는 검사 대상이 아니다
            var book = doc.Root?.Element(Main + "externalBook");
            if (book == null)
                return null;

            var rid = (string)book.Attribute(DocRel + "id");
            var rels = LoadXml(archive, $"xl/externalLinks/_rels/{Path.GetFileName(linkEntry.FullName)}.rels");
            var target = rels?.Root?.Elements(PkgRel + "Relationship")
                .Where(r => (string)r.Attribute("Id") == rid)
                .Select(r => (string)r.Attribute("Target"))
                .FirstOrDefault();
            if (string.IsNullOrEmpty(target))
                return null;

            var link = new ExternalLink { TargetFileName = ExtractFileName(target) };

            var sheetNames = book.Element(Main + "sheetNames")?.Elements(Main + "sheetName")
                .Select(e => (string)e.Attribute("val"))
                .ToList() ?? new List<string>();

            var sheetDataSet = book.Element(Main + "sheetDataSet");
            if (sheetDataSet != null)
            {
                foreach (var sheetData in sheetDataSet.Elements(Main + "sheetData"))
                {
                    var sheetId = (int?)sheetData.Attribute("sheetId") ?? 0;
                    if (sheetId < 0 || sheetId >= sheetNames.Count || sheetNames[sheetId] == null)
                        continue;

                    var cells = new Dictionary<(int Row, int Column), string>();
                    foreach (var cell in sheetData.Elements(Main + "row").SelectMany(r => r.Elements(Main + "cell")))
                    {
                        if (!TryParseCellReference((string)cell.Attribute("r"), out var rowCol))
                            continue;

                        var token = MakeCachedToken((string)cell.Attribute("t"), cell.Element(Main + "v")?.Value);
                        if (token != null)
                            cells[rowCol] = token;
                    }

                    link.CachedSheets[sheetNames[sheetId]] = cells;
                }
            }

            return link;
        }

        /// <summary>
        /// 외부 참조 캐시 셀을 비교용 토큰으로 정규화한다.
        /// 숫자는 표기 차이("1" vs "1.0")가 흡수되도록 파싱 후 라운드트립 포맷으로 통일.
        /// 빈 값은 null을 반환해 "셀 없음"과 동일하게 취급한다.
        /// </summary>
        private static string MakeCachedToken(string type, string value)
        {
            switch (type)
            {
                case "b":
                    return "B:" + (value == "1" ? "1" : "0");
                case "e":
                    return "E:" + value;
                case "str":
                    return string.IsNullOrEmpty(value) ? null : "S:" + UnescapeText(value);
                default: // null 또는 "n" = 숫자
                    if (string.IsNullOrEmpty(value))
                        return null;
                    return double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var number)
                        ? "N:" + number.ToString("R", CultureInfo.InvariantCulture)
                        : "S:" + value;
            }
        }

        private static string MakeLiveToken(object value)
        {
            switch (value)
            {
                case null:
                    return null;
                case string text:
                    return text.Length == 0 ? null : "S:" + text;
                case double number:
                    return "N:" + number.ToString("R", CultureInfo.InvariantCulture);
                case bool flag:
                    return "B:" + (flag ? "1" : "0");
                case CellError error:
                    return "E:" + error.Code;
                default:
                    return "S:" + value;
            }
        }

        private static List<string> LoadSharedStrings(ZipArchive archive)
        {
            var result = new List<string>();
            var entry = archive.GetEntry("xl/sharedStrings.xml");
            if (entry == null)
                return result;

            var doc = LoadXml(entry);
            if (doc.Root == null)
                return result;

            foreach (var si in doc.Root.Elements(Main + "si"))
                result.Add(GetRichText(si));

            return result;
        }

        /// <summary>
        /// si/is 요소의 텍스트를 추출한다. rich text run(r)은 전부 이어붙이고,
        /// 발음 표기(rPh)는 제외한다. xml:space="preserve"가 없는 공백 run도 보존된다.
        /// </summary>
        private static string GetRichText(XElement container)
        {
            var direct = container.Element(Main + "t");
            if (direct != null)
                return UnescapeText(direct.Value);

            return string.Concat(container.Elements(Main + "r")
                .Select(r => UnescapeText(r.Element(Main + "t")?.Value ?? "")));
        }

        // OOXML은 제어 문자 등을 _xHHHH_로 이스케이프한다 (_x005F_는 리터럴 '_x'의 이스케이프)
        private static string UnescapeText(string text)
        {
            if (text == null || !text.Contains("_x"))
                return text;

            return EscapedCharPattern.Replace(text, m =>
                ((char)Convert.ToInt32(m.Groups[1].Value, 16)).ToString());
        }

        private static string ExtractFileName(string target)
        {
            var path = Uri.UnescapeDataString(target);
            var slash = path.LastIndexOfAny(new[] { '/', '\\' });
            return slash >= 0 ? path.Substring(slash + 1) : path;
        }

        private static XDocument LoadXml(ZipArchive archive, string entryPath)
        {
            var entry = archive.GetEntry(entryPath);
            return entry == null ? null : LoadXml(entry);
        }

        private static XDocument LoadXml(ZipArchiveEntry entry)
        {
            using var stream = entry.Open();
            // 공백만 있는 rich text run(<t> </t>)이 사라지지 않도록 공백을 보존한다
            return XDocument.Load(stream, LoadOptions.PreserveWhitespace);
        }

        private static bool TryParseCellReference(string reference, out (int Row, int Column) rowCol)
        {
            rowCol = default;
            if (string.IsNullOrEmpty(reference))
                return false;

            var column = 0;
            var index = 0;
            while (index < reference.Length && char.IsLetter(reference[index]))
                column = column * 26 + (char.ToUpperInvariant(reference[index++]) - 'A' + 1);

            if (column == 0 || index >= reference.Length || !int.TryParse(reference.Substring(index), out var row))
                return false;

            rowCol = (row, column);
            return true;
        }

        public static string CellReference(int row, int column)
        {
            var letters = "";
            while (column > 0)
            {
                column--;
                letters = (char)('A' + column % 26) + letters;
                column /= 26;
            }
            return letters + row;
        }
    }
}

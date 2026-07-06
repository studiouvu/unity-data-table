using System.Collections.Concurrent;
using System.Globalization;
using System.Text;
using Newtonsoft.Json;

namespace DTable
{
    internal class ExcelToJson
    {
        private static int errorCount = 0;
        private static readonly object consoleLock = new object();
        private static readonly Dictionary<string, int> fileLineMap = new Dictionary<string, int>();
        private static readonly List<string> errorMessages = new List<string>();

        // Interop Value2가 수식 에러에 대해 반환하는 코드 값 (기존 출력과의 호환용)
        private static readonly Dictionary<string, string> errorCodes = new Dictionary<string, string>
        {
            { "#NULL!", "-2146826288" },
            { "#DIV/0!", "-2146826281" },
            { "#VALUE!", "-2146826273" },
            { "#REF!", "-2146826265" },
            { "#NAME?", "-2146826259" },
            { "#NUM!", "-2146826252" },
            { "#N/A", "-2146826246" },
        };

        private static string JsonOutputFolder => Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Jsons");

        public static int Main(string[] args)
        {
            try
            {
                Console.OutputEncoding = Encoding.UTF8;
                Console.WriteLine("START");

                var cpuCores = Environment.ProcessorCount;
                var workerCount = Math.Min(cpuCores, args.Length);
                Console.WriteLine($"CPU cores: {cpuCores} | Workers: {workerCount} | Files: {args.Length}");

                // 이전 실행의 결과물이 이번 실행 결과에 섞이지 않도록 출력 폴더를 비운다
                Directory.CreateDirectory(JsonOutputFolder);
                foreach (var oldJson in Directory.GetFiles(JsonOutputFolder, "*.json"))
                    File.Delete(oldJson);

                // 고정 뷰: 파일 목록을 미리 출력하고 각 라인 번호를 기록 (리다이렉트 시에는 일반 로그로 대체)
                foreach (var path in args)
                {
                    if (!Console.IsOutputRedirected)
                        fileLineMap[path] = Console.CursorTop;
                    Console.WriteLine($"⏳ - {Path.GetFileName(path)}");
                }

                var queue = new ConcurrentQueue<string>(args);
                var threads = new List<Thread>();

                for (var i = 0; i < workerCount; i++)
                {
                    var thread = new Thread(() => Worker(queue));
                    thread.Start();
                    threads.Add(thread);
                }

                foreach (var thread in threads)
                {
                    thread.Join();
                }

                Console.WriteLine("🏁 END");
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                RecordError(e.ToString());
            }

            if (errorCount > 0)
            {
                Console.WriteLine();
                Console.WriteLine("=== ❌ errors ===");
                lock (errorMessages)
                {
                    foreach (var msg in errorMessages)
                        Console.WriteLine(msg);
                }

                if (!Console.IsInputRedirected)
                {
                    Console.WriteLine($"오류 {errorCount}건 발생. 종료하려면 아무 키나 누르세요...");
                    Console.ReadKey();
                }
                return 1;
            }

            return 100;
        }

        private static void RecordError(string message)
        {
            Interlocked.Increment(ref errorCount);
            lock (errorMessages) errorMessages.Add(message);
        }

        private static void UpdateStatus(string filePath, string status)
        {
            lock (consoleLock)
            {
                if (!fileLineMap.TryGetValue(filePath, out var line))
                {
                    Console.WriteLine($"{status} - {Path.GetFileName(filePath)}");
                    return;
                }

                int savedTop = Console.CursorTop;
                int savedLeft = Console.CursorLeft;
                try
                {
                    Console.SetCursorPosition(0, line);
                    var text = $"{status} - {Path.GetFileName(filePath)}";
                    int width = Console.WindowWidth - 1;
                    if (width > 0)
                    {
                        text = text.Length >= width ? text.Substring(0, width) : text.PadRight(width);
                    }
                    Console.Write(text);
                }
                catch
                {
                    // 콘솔이 리다이렉트 됐거나 좌표가 유효하지 않으면 무시
                }
                finally
                {
                    try { Console.SetCursorPosition(savedLeft, savedTop); } catch { }
                }
            }
        }

        private static void Worker(ConcurrentQueue<string> queue)
        {
            while (queue.TryDequeue(out var path))
            {
                Process(path);
            }
        }

        public static void Process(string fileExcel)
        {
            if (File.Exists(fileExcel) == false)
            {
                UpdateStatus(fileExcel, "🚫 file not found");
                RecordError($"[{Path.GetFileName(fileExcel)}] File not found: {fileExcel}");
                return;
            }

            try
            {
                var file = XlsxFile.Load(fileExcel);
                var outputs = new List<(string SheetName, string Json)>();
                var exportedExternalCells = new List<string>();
                var exportedUncachedCells = new List<string>();

                foreach (var sheet in file.Sheets)
                {
                    if (sheet.Name.Contains('#'))
                        continue;

                    var json = ConvertSheet(sheet, exportedExternalCells, exportedUncachedCells);
                    outputs.Add((sheet.Name, json));
                }

                // export되는 셀의 수식 계산 값이 파일에 없으면 빈 값이 나가므로 실패시킨다
                if (exportedUncachedCells.Count > 0)
                {
                    throw new MissingFormulaCacheException(
                        "수식의 계산 값이 파일에 저장되어 있지 않습니다 (Google 시트 다운로드 등 Excel 외 도구로 생성된 파일). " +
                        "Excel에서 열어 저장한 뒤 다시 변환하세요.\n" +
                        $"  - 계산 값이 없는 export 셀: {XlsxFile.Summarize(exportedUncachedCells)}");
                }

                // export되는 셀에 외부 파일 참조 수식이 있으면, 그 캐시가 원본 파일과 일치할 때만 통과
                if (exportedExternalCells.Count > 0)
                    file.ValidateCacheFreshness(fileExcel, exportedExternalCells);

                foreach (var (sheetName, json) in outputs)
                {
                    var fileName = Path.Combine(JsonOutputFolder, $"{sheetName.Replace(".xlsx", "")}.json");
                    File.WriteAllText(fileName, json);
                }

                UpdateStatus(fileExcel, "✅ ");
            }
            catch (StaleExternalCacheException e)
            {
                UpdateStatus(fileExcel, "❌ ");
                RecordError($"[{Path.GetFileName(fileExcel)}] {e.Message}");
            }
            catch (MissingFormulaCacheException e)
            {
                UpdateStatus(fileExcel, "❌ ");
                RecordError($"[{Path.GetFileName(fileExcel)}] {e.Message}");
            }
            catch (Exception e)
            {
                UpdateStatus(fileExcel, "❌ ");
                RecordError($"[{Path.GetFileName(fileExcel)}] {e}");
            }
        }

        private static string ConvertSheet(
            XlsxSheet sheet,
            List<string> exportedExternalCells,
            List<string> exportedUncachedCells)
        {
            var rowCount = sheet.RowCount;
            var columnCount = sheet.ColumnCount;

            // 헤더 행에서 #이 있는 컬럼 미리 표시 + 유효한 컬럼 끝 계산
            var skipColumn = new bool[columnCount + 2];
            var maxColumn = 0;
            for (var column = 1; column <= columnCount; column++)
            {
                var header = sheet.GetValue(1, column);
                if (column > 1 && header == null)
                    break;

                maxColumn = column;

                if (column > 1 && header != null && header.ToString().Contains('#'))
                    skipColumn[column] = true;
            }

            var rows = new List<List<string>>();

            for (var row = 1; row <= rowCount; row++)
            {
                if (row > 2 && sheet.GetValue(row, 2) == null)
                    break;

                // #행 여부 검사 (1번 행은 예외)
                var isHashRow = false;
                if (row > 1)
                {
                    var firstCell = sheet.GetValue(row, 1);
                    if (firstCell != null && firstCell.ToString().Contains('#'))
                        isHashRow = true;
                }

                var list = new List<string>(maxColumn);
                rows.Add(list);

                for (var column = 1; column <= maxColumn; column++)
                {
                    // 1번 행/열은 항상 실제 값 (#마커 보존)
                    // 그 외 셀에서 #행 또는 #열이면 값 읽지 않고 빈 문자열
                    if (row > 1 && column > 1 && (isHashRow || skipColumn[column]))
                    {
                        list.Add("");
                        continue;
                    }

                    if (sheet.ExternalFormulaCells.Contains((row, column)))
                        exportedExternalCells.Add($"{sheet.Name}!{XlsxFile.CellReference(row, column)}");
                    if (sheet.UncachedFormulaCells.Contains((row, column)))
                        exportedUncachedCells.Add($"{sheet.Name}!{XlsxFile.CellReference(row, column)}");

                    list.Add(CellToString(sheet.GetValue(row, column)));
                }
            }

            return JsonConvert.SerializeObject(rows, Formatting.Indented);
        }

        private static string CellToString(object cellValue)
        {
            switch (cellValue)
            {
                case null:
                    return "";
                case string text:
                    return text;
                case double number:
                    // OS locale에 따라 소수점 표기가 달라져 JSON diff가 생기지 않도록 고정
                    return number.ToString(CultureInfo.InvariantCulture);
                case bool flag:
                    return flag.ToString();
                case CellError error:
                    return errorCodes.TryGetValue(error.Code, out var code) ? code : error.Code;
                default:
                    return cellValue.ToString();
            }
        }
    }
}

using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using Newtonsoft.Json;
using Excel = Microsoft.Office.Interop.Excel;

namespace DTable
{
    internal class ExcelToJson
    {
        private static int errorCount = 0;
        private static readonly object consoleLock = new object();
        private static readonly Dictionary<string, int> fileLineMap = new Dictionary<string, int>();
        private static readonly List<string> errorMessages = new List<string>();

        public static int Main(string[] args)
        {
            try
            {
                Console.OutputEncoding = Encoding.UTF8;
                Console.WriteLine("START");

                // 고정 뷰: 파일 목록을 미리 출력하고 각 라인 번호를 기록
                foreach (var path in args)
                {
                    fileLineMap[path] = Console.CursorTop;
                    Console.WriteLine($"⏳ - {Path.GetFileName(path)}");
                }

                var threads = new List<Thread>();

                foreach (var path in args)
                {
                    var thread = new Thread(() => Process(path));
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
                Interlocked.Increment(ref errorCount);
                lock (errorMessages) errorMessages.Add(e.ToString());
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
                Console.WriteLine($"오류 {errorCount}건 발생. 종료하려면 아무 키나 누르세요...");
                Console.ReadKey();
                return 1;
            }

            return 100;
        }

        private static void UpdateStatus(string filePath, string status)
        {
            lock (consoleLock)
            {
                if (!fileLineMap.TryGetValue(filePath, out var line))
                    return;

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

        public static void Process(string fileExcel)
        {
            if (File.Exists(fileExcel) == false)
            {
                UpdateStatus(fileExcel, "🚫 file not found");
                Interlocked.Increment(ref errorCount);
                lock (errorMessages)
                    errorMessages.Add($"[{Path.GetFileName(fileExcel)}] File not found: {fileExcel}");
                return;
            }

            Excel.Application excelApp = null;
            Excel.Workbook workBook = null;

            try
            {
                excelApp = new Excel.Application();
                excelApp.Visible = false;
                excelApp.ScreenUpdating = false;
                excelApp.DisplayAlerts = false;
                excelApp.EnableEvents = false;

                workBook = excelApp.Workbooks.Open(fileExcel, UpdateLinks: 0, ReadOnly: true);

                // Calculation은 워크북이 열려있어야 설정 가능
                excelApp.Calculation = Excel.XlCalculation.xlCalculationManual;

                foreach (Excel.Worksheet workSheet in workBook.Worksheets)
                {
                    if (workSheet.Name.Contains('#'))
                        continue;

                    Excel.Range range = workSheet.UsedRange;
                    object[,] cellValues = (object[,])range.Value2;

                    var rowCount = cellValues.GetLength(0);
                    var columnCount = cellValues.GetLength(1);

                    // 헤더 행에서 #이 있는 컬럼 미리 표시 + 유효한 컬럼 끝 계산
                    var skipColumn = new bool[columnCount + 2];
                    var maxColumn = 0;
                    for (var column = 1; column <= columnCount; column++)
                    {
                        var header = cellValues[1, column];
                        if (column > 1 && header == null)
                            break;

                        maxColumn = column;

                        if (column > 1 && header != null && header.ToString().Contains('#'))
                            skipColumn[column] = true;
                    }

                    var dictionary = new Dictionary<int, List<string>>();

                    for (var row = 1; row <= rowCount; row++)
                    {
                        if (row > 2 && (cellValues[row, 2] == null))
                            break;

                        // #행 여부 검사 (1번 행은 예외)
                        var isHashRow = false;
                        if (row > 1)
                        {
                            var firstCell = cellValues[row, 1];
                            if (firstCell != null && firstCell.ToString().Contains('#'))
                                isHashRow = true;
                        }

                        var list = new List<string>(maxColumn);
                        dictionary.Add(row, list);

                        for (var column = 1; column <= maxColumn; column++)
                        {
                            // 1번 행/열은 항상 실제 값 (#마커 보존)
                            // 그 외 셀에서 #행 또는 #열이면 값 읽지 않고 빈 문자열
                            if (row > 1 && column > 1 && (isHashRow || skipColumn[column]))
                            {
                                list.Add("");
                                continue;
                            }

                            var cellValue = cellValues[row, column];
                            list.Add(cellValue == null ? "" : cellValue.ToString());
                        }
                    }

                    var json = JsonConvert.SerializeObject(dictionary, Formatting.Indented);

                    var saveFolder = new DirectoryInfo($"{AppDomain.CurrentDomain.BaseDirectory}/Jsons");
                    if (saveFolder.Exists == false)
                        saveFolder.Create();

                    var fileName = $"{AppDomain.CurrentDomain.BaseDirectory}/Jsons/{workSheet.Name.Replace(".xlsx", "")}.json";
                    File.WriteAllText(fileName, json);
                }

                object missing = Type.Missing;
                object noSave = false;
                workBook.Close(noSave, missing, missing);
                excelApp.Quit();

                UpdateStatus(fileExcel, "✅ ");
            }
            catch (Exception e)
            {
                UpdateStatus(fileExcel, "❌ ");
                Interlocked.Increment(ref errorCount);
                lock (errorMessages)
                    errorMessages.Add($"[{Path.GetFileName(fileExcel)}] {e}");

                try { workBook?.Close(false); } catch { }
                try { excelApp?.Quit(); } catch { }
            }
            finally
            {
                ReleaseObject(workBook);
                ReleaseObject(excelApp);
            }
        }

        public static void ReleaseObject(object obj)
        {
            try
            {
                if (obj != null)
                {
                    Marshal.ReleaseComObject(obj);
                    obj = null;
                }
            }
            catch (Exception ex)
            {
                obj = null;
                throw ex;
            }
            finally
            {
                GC.Collect();
            }
        }
    }
}

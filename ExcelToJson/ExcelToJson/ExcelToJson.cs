using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using Newtonsoft.Json;
using Excel = Microsoft.Office.Interop.Excel;

namespace DTable
{
    internal class ExcelToJson
    {
        public static int Main(string[] args)
        {
            try
            {
                Console.WriteLine($"START");

                foreach (var path in args)
                    Process(path);

                Console.WriteLine($"END");
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                Console.ReadKey();
                throw;
            }

            return 100;
        }

        public static void Process(string fileExcel)
        {
            Console.WriteLine($"Read Start {fileExcel}");

            if (File.Exists(fileExcel) == false)
            {
                Console.WriteLine($"File not found");
                return;
            }

            Excel.Application excelApp = null;
            Excel.Workbook workBook = null;

            try
            {
                excelApp = new Excel.Application(); // 엑셀 어플리케이션 생성
                workBook = excelApp.Workbooks.Open(fileExcel,
                    0,
                    true,
                    5,
                    "",
                    "",
                    true,
                    Excel.XlPlatform.xlWindows,
                    "\t",
                    false,
                    false,
                    0,
                    true,
                    1,
                    0);

                // Sheet항목들을 돌아가면서 내용을 확인
                foreach (Excel.Worksheet workSheet in workBook.Worksheets)
                {
                    Console.WriteLine($"Sheet name : {workSheet.Name}");

                    Excel.Range range = workSheet.UsedRange; // 사용중인 셀 범위를 가져오기

                    var dictionary = new Dictionary<int, List<string>>();

                    for (int row = 1; row <= range.Rows.Count; row++)
                    {
                        dictionary.Add(row, new List<string>());

                        if (row > 2 && (range.Cells[row, 2] == null || string.IsNullOrEmpty((range.Cells[row, 2] as Excel.Range)?.Value2?.ToString())))
                            continue;

                        for (int column = 1; column <= range.Columns.Count; column++)
                        {
                            if (column > 1 && (range.Cells[1, column] == null || string.IsNullOrEmpty((range.Cells[1, column] as Excel.Range)?.Value2?.ToString())))
                                break;

                            var obj = ((Excel.Range)range.Cells[row, column])?.Value2;

                            var str = obj == null ? "" : obj.ToString(); // 셀 데이터 가져옴

                            dictionary[row].Add(str);
                        }
                    }

                    var json = JsonConvert.SerializeObject(dictionary, Formatting.Indented);

                    var saveFolder = new DirectoryInfo($"{AppDomain.CurrentDomain.BaseDirectory}/Jsons");
                    
                    if (saveFolder.Exists == false)
                        saveFolder.Create();

                    var fileName = $"{AppDomain.CurrentDomain.BaseDirectory}/Jsons/{workBook.Name.Replace(".xlsx", "")}.json";

                    File.WriteAllText(fileName, json);

                    Console.WriteLine($"Save file : {fileName}");
                }

                object missing = Type.Missing;
                object noSave = false;
                workBook.Close(noSave, missing, missing); // 엑셀 웨크북 종료
                excelApp.Quit(); // 엑셀 어플리케이션 종료
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
                    // 객체 메모리 해제
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
                GC.Collect(); // 가비지 수집
            }
        }
    }

}

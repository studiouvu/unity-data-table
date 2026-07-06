using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using UnityEditor;
using UnityEngine;

namespace TableManager
{
    public static class TableManagerExport
    {
        [MenuItem("TableManager/Export %F8")]
        public static async void Export()
        {
            Debug.Log($"[{nameof(TableManager)}] Export Start");

            var excelToJsonResult = await ExportExcelToJson();

            if (!excelToJsonResult)
                return;

            Debug.Log($"[{nameof(TableManager)}] Create Class Start");
            await CreateClassFiles();

            Debug.Log($"[{nameof(TableManager)}] Export End");

            AssetDatabase.Refresh();
        }

        private static async Task<bool> ExportExcelToJson()
        {
            // Application.dataPath 의존 경로는 메인 스레드에서 미리 캐싱한다 (Task.Run 내부는 백그라운드 스레드).
            var launchFolderPath = TableManagerConfig.LaunchFolderPath;
            var jsonPath = TableManagerConfig.JsonDataPath;

            var directoryInfo = new DirectoryInfo(TableManagerConfig.ExcelDataPath);
            var files = directoryInfo.GetFiles("*", SearchOption.AllDirectories);

            var stringBuilder = new StringBuilder();

            foreach (var fileInfo in files.Where(info => !info.Name.Contains($".meta") && !info.Name.Contains($"~")))
            {
                Debug.Log($"[{nameof(TableManager)} Export asset find] {fileInfo.FullName}");

                // 경로에 공백이 있어도 인자가 쪼개지지 않도록 따옴표로 감싼다
                stringBuilder.Append($"\"{fileInfo.FullName}\" ");
            }

            var exeName = launchFolderPath + "/ExcelToJson.exe";
            var arguments = stringBuilder.ToString();

            // 프로세스 대기 + 디렉토리 이동/삭제를 백그라운드 스레드로 옮겨 메인 스레드(에디터) 블로킹을 막는다.
            return await Task.Run(() =>
            {
                var process = System.Diagnostics.Process.Start(exeName, arguments);
                process.WaitForExit();

                // 변환이 실패하면 일부 JSON이 누락된 상태이므로 기존 Jsons를 건드리지 않는다
                if (process.ExitCode != 100)
                    return false;

                if (Directory.Exists(jsonPath))
                    Directory.Delete(jsonPath, true);

                Directory.Move(launchFolderPath + "/Jsons", jsonPath);

                return true;
            });
        }

        private static async Task CreateClassFiles()
        {
            // Application.dataPath 의존 경로는 메인 스레드에서 미리 캐싱한다 (Task.Run 내부는 백그라운드 스레드).
            var jsonDataPath = TableManagerConfig.JsonDataPath;
            var folderPath = TableManagerConfig.FolderPath;
            var classPath = folderPath + "/Class";

            await Task.Run(() =>
            {
                var directoryInfo = new DirectoryInfo(jsonDataPath);
                var files = directoryInfo.GetFiles();

                Directory.Delete(classPath, true);
                Directory.CreateDirectory(classPath);

                var template = File.ReadAllText(folderPath + "/Editor/ClassTemplate.txt");

                foreach (var fileInfo in files.Where(info => !info.Name.Contains($".meta")))
                {
                    Debug.Log($"[{nameof(TableManager)} Export create class] {fileInfo.FullName}");

                    var fileName = fileInfo.Name.Replace(".json", "");

                    var text = File.ReadAllText(fileInfo.FullName);
                    var rows = JsonConvert.DeserializeObject<List<List<string>>>(text);

                    var stringBuilder = new StringBuilder();

                    for (var i = 0; i < rows[0].Count; i++)
                    {
                        var valueType = rows[0][i];

                        if (string.IsNullOrEmpty(valueType) || valueType == "#" || valueType == "[]")
                            continue;

                        var valueDesc = rows[1][i];
                        var valueName = rows[2][i];

                        stringBuilder.Append($"        /* {valueDesc} */\n        public {valueType} {valueName} {{ get; set; }} \n\n");
                    }

                    File.WriteAllText($"{classPath}/{fileName}Row.cs",
                        template
                            .Replace("#", $"{fileName}Row")
                            .Replace("@", stringBuilder.ToString()));
                }
            });
        }
    }
}

using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using Newtonsoft.Json;
using UnityEditor;
using UnityEngine;
using Debug = UnityEngine.Debug;

namespace TableManager
{
    public static class TableManagerExport
    {
        [MenuItem("TableManager/Export %F8")]
        public static void Export()
        {
            Debug.Log($"[{nameof(TableManager)}] Export Start");
            var result = ExportExcelToJson();

            if (result != 100)
                return;

            Debug.Log($"[{nameof(TableManager)}] Create Class Start");
            CreateClassFiles();

            Debug.Log($"[{nameof(TableManager)}] Export End");
            AssetDatabase.Refresh();

            LocalDb.Init();
        }

        private static int ExportExcelToJson()
        {
            var directoryInfo = new DirectoryInfo(TableManagerConfig.ExcelDataPath);
            var files = directoryInfo.GetFiles("*", SearchOption.AllDirectories);

            var stringBuilder = new StringBuilder();

            foreach (var fileInfo in files.Where(info => !info.Name.Contains($".meta") && !info.Name.Contains($"~")))
            {
                Debug.Log($"[{nameof(TableManager)} Export asset find] {fileInfo.FullName}");

                stringBuilder.Append($"{fileInfo.FullName} ");
            }

            var jsonPath = TableManagerConfig.JsonDataPath;
            if (Directory.Exists(jsonPath))
                Directory.Delete(jsonPath, true);
            Directory.CreateDirectory(jsonPath);

            var exe_name = TableManagerConfig.FolderPath + "/ExcelToJson.exe";
            var process = Process.Start(exe_name, stringBuilder.ToString());

            process.WaitForExit();
            return process.ExitCode;
        }

        private static void CreateClassFiles()
        {
            var directoryInfo = new DirectoryInfo(TableManagerConfig.JsonDataPath);
            var files = directoryInfo.GetFiles();

            var classPath = TableManagerConfig.FolderPath + "/Class";
            Directory.Delete(classPath, true);
            Directory.CreateDirectory(classPath);

            foreach (var fileInfo in files.Where(info => !info.Name.Contains($".meta")))
            {
                Debug.Log($"[{nameof(TableManager)} Export create class] {fileInfo.FullName}");

                var fileName = fileInfo.Name.Replace(".json", "");

                var text = File.ReadAllText(fileInfo.FullName);
                var jsonDictionary = JsonConvert.DeserializeObject<Dictionary<int, List<string>>>(text);

                var stringBuilder = new StringBuilder();

                for (var i = 0; i < jsonDictionary[1].Count; i++)
                {
                    if (string.IsNullOrEmpty(jsonDictionary[1][i]) || jsonDictionary[1][i] == "#")
                        continue;

                    stringBuilder.Append($"        /* {jsonDictionary[2][i]} */\n        public {jsonDictionary[1][i]} {jsonDictionary[3][i]} {{ get; set; }} \n\n");
                }

                var template = File.ReadAllText(TableManagerConfig.FolderPath + "/Editor/ClassTemplate.txt");

                File.WriteAllText($"{Application.dataPath}/{nameof(TableManager)}/Class/{fileName}Row.cs",
                    template
                        .Replace("#", $"{fileName}Row")
                        .Replace("@", stringBuilder.ToString()));
            }
        }
    }
}

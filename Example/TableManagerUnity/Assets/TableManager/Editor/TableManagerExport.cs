using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using TableManager.Editor;
using Unity.Plastic.Newtonsoft.Json.Linq;
using UnityEditor;
using UnityEngine;
using UnityEngine.Networking;

namespace TableManager
{
    public static class TableManagerExport
    {
        private static readonly string GoogleConfigJsonPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
            ".table-manager",
            $"{nameof(Editor.GoogleConfig)}.json");

        private static readonly string GoogleSheetIdsJsonPath = Path.Combine(
            Application.dataPath,
            "TableManager",
            "Editor",
            $"{nameof(Editor.GoogleSheetIds)}.json");

        private static GoogleConfig GoogleConfig
        {
            get
            {
                if (config != null)
                    return config;

                config = new GoogleConfig();

                if (File.Exists(GoogleConfigJsonPath))
                {
                    JsonUtility.FromJsonOverwrite(File.ReadAllText(GoogleConfigJsonPath), config);
                    return config;
                }

                SaveConfigJson();
                EditorUtility.RevealInFinder(GoogleConfigJsonPath);
                return config;
            }
        }

        private static GoogleConfig config;

        private static GoogleSheetIds GoogleSheetIds
        {
            get
            {
                if (sheetIds != null)
                    return sheetIds;

                sheetIds = new GoogleSheetIds();

                if (File.Exists(GoogleSheetIdsJsonPath))
                {
                    JsonUtility.FromJsonOverwrite(File.ReadAllText(GoogleSheetIdsJsonPath), sheetIds);
                    return sheetIds;
                }

                SaveSheetIdsJson();
                EditorUtility.RevealInFinder(GoogleSheetIdsJsonPath);
                return sheetIds;
            }
        }

        private static GoogleSheetIds sheetIds;

        private static void SaveConfigJson()
        {
            var dir = Path.GetDirectoryName(GoogleConfigJsonPath);
            if (!string.IsNullOrEmpty(dir))
                Directory.CreateDirectory(dir);
            File.WriteAllText(GoogleConfigJsonPath, JsonUtility.ToJson(config, true));
        }

        private static void SaveSheetIdsJson()
        {
            var dir = Path.GetDirectoryName(GoogleSheetIdsJsonPath);
            if (!string.IsNullOrEmpty(dir))
                Directory.CreateDirectory(dir);
            File.WriteAllText(GoogleSheetIdsJsonPath, JsonUtility.ToJson(sheetIds, true));
        }


        [MenuItem("TableManager/Export %F8")]
        public static async void Export()
        {
            Debug.Log($"[{nameof(TableManager)}] Export Start");

            var targetSheetIds = GoogleSheetIds.sheetIds;

            if (targetSheetIds != null && targetSheetIds.Count > 0)
            {
                if (!ValidateConfig())
                    return;

                foreach (var sheetId in targetSheetIds)
                {
                    if (string.IsNullOrEmpty(sheetId))
                        continue;

                    var sheetDownloadResult = await GoogleSheetDownload(sheetId);

                    if (!sheetDownloadResult)
                        return;
                }
            }

            var excelToJsonResult = ExportExcelToJson();

            if (!excelToJsonResult)
                return;

            Debug.Log($"[{nameof(TableManager)}] Create Class Start");
            CreateClassFiles();

            Debug.Log($"[{nameof(TableManager)}] Export End");

            AssetDatabase.Refresh();

        }

        private static bool ValidateConfig()
        {
            var c = GoogleConfig;
            if (string.IsNullOrEmpty(c.clientId)
                || string.IsNullOrEmpty(c.clientSecret)
                || string.IsNullOrEmpty(c.apiKey))
            {
                EditorUtility.RevealInFinder(GoogleConfigJsonPath);
                Debug.LogError(
                    $"[{nameof(TableManager)}] Please fill in clientId/clientSecret/apiKey in {GoogleConfigJsonPath}.");
                return false;
            }
            return true;
        }

        private static async Task<bool> GoogleSheetDownload(string sheetId)
        {
            Debug.Log($"TableManagerExport.GoogleSheetDownload - Start sheetId={sheetId}");

            var ab = new GoogleAuthApi();

            string token = null;

            if (!File.Exists(GoogleAuthApi.GoogleLoginDataPath))
            {
                await GetToken(ab);
            }
            else
            {
                var text = await File.ReadAllTextAsync(GoogleAuthApi.GoogleLoginDataPath);
                var loginData = JsonConvert.DeserializeObject<GoogleLogin>(text);

                if (DateTime.Now > loginData.expireTime)
                {
                    await GoogleAuthApi.RequestRefreshToken(config.clientId, config.clientSecret, loginData);
                }

                token = loginData.accessToken;
            }

            var jsonText = await ab.RequestSheetInfo(config.apiKey, token, sheetId);

            Debug.Log($"TableManagerExport.GoogleSheetDownload - json : {jsonText}");

            var jObject = JObject.Parse(jsonText);
            var sheetTitle = jObject["properties"]?["title"]?.ToString();

            var url = $"https://www.googleapis.com/drive/v3/files/{sheetId}/export?mimeType=application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

            var www = UnityWebRequest.Get(url);
            www.SetRequestHeader("Authorization", "Bearer " + token);

            var ao = www.SendWebRequest(); // 응답이 올때까지 대기한다.

            await ao;

            await File.WriteAllBytesAsync(Application.dataPath + $"/Excels/{sheetTitle}.xlsx", ao.webRequest.downloadHandler.data);

            Debug.Log($"TableManagerExport.GoogleSheetDownload - {ao.webRequest.result.ToString()}");

            return ao.webRequest.result == UnityWebRequest.Result.Success;
        }

        private static async Task GetToken(GoogleAuthApi authApi)
        {
            try
            {
                await authApi.DoOAuthAsync(config.clientId, config.clientSecret);
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                throw;
            }
        }

        private static bool ExportExcelToJson()
        {
            var directoryInfo = new DirectoryInfo(TableManagerConfig.ExcelDataPath);
            var files = directoryInfo.GetFiles("*", SearchOption.AllDirectories);

            var stringBuilder = new StringBuilder();

            foreach (var fileInfo in files.Where(info => !info.Name.Contains($".meta") && !info.Name.Contains($"~")))
            {
                Debug.Log($"[{nameof(TableManager)} Export asset find] {fileInfo.FullName}");

                stringBuilder.Append($"{fileInfo.FullName} ");
            }

            var exeName = TableManagerConfig.LaunchFolderPath + "/ExcelToJson.exe";
            var process = System.Diagnostics.Process.Start(exeName, stringBuilder.ToString());
            process.WaitForExit();

            var jsonPath = TableManagerConfig.JsonDataPath;

            if (Directory.Exists(jsonPath))
                Directory.Delete(jsonPath, true);

            Directory.Move(TableManagerConfig.LaunchFolderPath + "/Jsons", jsonPath);

            return process.ExitCode == 100;
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
                    var valueType = jsonDictionary[1][i];

                    if (string.IsNullOrEmpty(valueType) || valueType == "#" || valueType == "[]")
                        continue;

                    var valueDesc = jsonDictionary[2][i];
                    var valueName = jsonDictionary[3][i];

                    stringBuilder.Append($"        /* {valueDesc} */\n        public {valueType} {valueName} {{ get; set; }} \n\n");
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

public struct UnityWebRequestAwaiter : INotifyCompletion
{
    private UnityWebRequestAsyncOperation asyncOp;
    private Action continuation;

    public UnityWebRequestAwaiter(UnityWebRequestAsyncOperation asyncOp)
    {
        this.asyncOp = asyncOp;
        continuation = null;
    }

    public bool IsCompleted { get { return asyncOp.isDone; } }

    public void GetResult() {}

    public void OnCompleted(Action continuation)
    {
        this.continuation = continuation;
        asyncOp.completed += OnRequestCompleted;
    }

    private void OnRequestCompleted(AsyncOperation obj)
    {
        continuation?.Invoke();
    }
}

public static class ExtensionMethods
{
    public static UnityWebRequestAwaiter GetAwaiter(this UnityWebRequestAsyncOperation asyncOp)
    {
        return new UnityWebRequestAwaiter(asyncOp);
    }
}

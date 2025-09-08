using System;
using System.Collections.Generic;
using System.Diagnostics;
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
using Debug = UnityEngine.Debug;

namespace TableManager
{
    public static class TableManagerExport
    {
        private static string ClientId => Config.clientId;
        private static string ClientSecret => Config.clientSecret;
        private static string ApiKey => Config.apiKey;
        private static string SheetId => Config.sheetId;

        private static GoogleSheetConfig Config
        {
            get {
                if (config == null)
                    config = AssetDatabase.LoadAssetAtPath<GoogleSheetConfig>($"Assets/{nameof(TableManager)}/Editor/GoogleSheetConfig.asset");
                return config;
            }
        }

        private static GoogleSheetConfig config;


        [MenuItem("TableManager/Export %F8")]
        public static async void Export()
        {
            try
            {
                Debug.Log($"[{nameof(TableManager)}] Export Start");

                var sheetDownloadResult = await GoogleSheetDownload();

                if (!sheetDownloadResult)
                    return;

                var excelToJsonResult = ExportExcelToJson();
                
                if (!excelToJsonResult)
                    return;

                Debug.Log($"[{nameof(TableManager)}] Create Class Start");
                CreateClassFiles();

                Debug.Log($"[{nameof(TableManager)}] Export End");

                AssetDatabase.Refresh();
            }
            catch (Exception e)
            {
                Debug.LogError(e);
                throw;
            }
        }

        private static async Task<bool> GoogleSheetDownload()
        {
            Debug.Log($"TableManagerExport.GoogleSheetDownload - Start");

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
                    await GoogleAuthApi.RequestRefreshToken(ClientId, ClientSecret, loginData);
                }

                token = loginData.accessToken;
            }

            var jsonText = await ab.RequestSheetInfo(ApiKey, token, SheetId);

            Debug.Log($"TableManagerExport.GoogleSheetDownload - json : {jsonText}");

            var jObject = JObject.Parse(jsonText);
            var sheetTitle = jObject["properties"]?["title"]?.ToString();

            var url = $"https://www.googleapis.com/drive/v3/files/{SheetId}/export?mimeType=application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

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
                await authApi.DoOAuthAsync(ClientId, ClientSecret);
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

            var jsonPath = TableManagerConfig.JsonDataPath;
            if (Directory.Exists(jsonPath))
                Directory.Delete(jsonPath, true);
            Directory.CreateDirectory(jsonPath);

            var exe_name = TableManagerConfig.FolderPath + "/ExcelToJson.exe";
            var process = Process.Start(exe_name, stringBuilder.ToString());
            process.WaitForExit();

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

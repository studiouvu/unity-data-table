using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using TableManager.Editor;
using UnityEditor;
using UnityEngine;
using UnityEngine.Networking;
using Debug = UnityEngine.Debug;

namespace TableManager
{
    public static class TableManagerExport
    {
        public const string ClientId = "";
        public const string ClientSecret = "";
        public const string ApiKey = "";
        
        [MenuItem("TableManager/Export %F8")]
        public static async void Export()
        {
            try
            {
                Debug.Log($"[{nameof(TableManager)}] Export Start");

                await GoogleSheetDownload();

                var result = ExportExcelToJson();

                if (result != 100)
                    return;

                Debug.Log($"[{nameof(TableManager)}] Create Class Start");
                CreateClassFiles();

                Debug.Log($"[{nameof(TableManager)}] Export End");
                
                AssetDatabase.Refresh();
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                throw;
            }
        }

        private static async Task GoogleSheetDownload()
        {
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

            var json = await ab.RequestSheetInfo(ApiKey, token, "1I_-ztdMD6ffRZ2BFE55J4F9_yNSuyy_nZUKB-UHbMss");

            var url = "https://www.googleapis.com/drive/v3/files/1I_-ztdMD6ffRZ2BFE55J4F9_yNSuyy_nZUKB-UHbMss/export?mimeType=application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

            var www = UnityWebRequest.Get(url);
            www.SetRequestHeader("Authorization", "Bearer " + token);

            var ao = www.SendWebRequest(); // 응답이 올때까지 대기한다.

            await ao;

            await File.WriteAllBytesAsync(Application.dataPath + "/Excels/download.xlsx", ao.webRequest.downloadHandler.data);

            Debug.Log(ao.webRequest.result.ToString());
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

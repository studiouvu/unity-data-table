using UnityEngine;
namespace TableManager
{
    public static class TableManagerConfig
    {
        public static string FolderPath => Application.dataPath + "/TableManager";
        public static string ExcelDataPath => $"{Application.dataPath}/Excels";
        public static string JsonDataPath => $"{FolderPath}/Jsons";
    }
}

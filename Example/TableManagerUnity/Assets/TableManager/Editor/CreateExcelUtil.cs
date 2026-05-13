using System.IO;
using TableManager;
namespace UnityEditor
{
    public static class CreateExcelUtil
    {
        [MenuItem("Assets/Create Excel File", priority = -999)]
        public static void CreateExcelFile()
        {
            var path = "Assets";
            var obj = Selection.activeObject;
            if (obj && AssetDatabase.Contains(obj))
            {
                path = AssetDatabase.GetAssetPath(obj);
                if (!Directory.Exists(path))
                {
                    path = Path.GetDirectoryName(path);
                }
            }

            var dest = $"{path}/NewExcelFile.xlsx";
            var index = 2;
            while (File.Exists(dest))
            {
                dest = $"{path}/NewExcelFile{index}.xlsx";
                index++;
            }

            File.Copy(TableManagerConfig.FolderPath + "/Editor/ExcelFileTemplate.xlsx", dest);
            AssetDatabase.Refresh();
            Selection.activeObject = AssetDatabase.LoadAssetAtPath<DefaultAsset>(dest);
        }
    }
}

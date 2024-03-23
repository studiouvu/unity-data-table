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

            // var b = File.ReadAllText(ATableConfig.FolderPath + "/Editor/ExcelFileTemplate.xlsx");
            // ProjectWindowUtil.CreateAssetWithContent($"a.xlsx", b);
            // EditorGUIUtility.PingObject(a);

            var dest = $"{path}/NewExcelFile.xlsx";

            File.Copy(TableManagerConfig.FolderPath + "/Editor/ExcelFileTemplate.xlsx", dest);
            AssetDatabase.Refresh();
            Selection.activeObject = AssetDatabase.LoadAssetAtPath<DefaultAsset>(dest);
        }
    }
}

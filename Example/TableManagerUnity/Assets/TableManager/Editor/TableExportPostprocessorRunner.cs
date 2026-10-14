using System;
using System.IO;
using System.Linq;
using UnityEditor;
using UnityEngine;

namespace TableManager
{
    public static class TableExportPostprocessorRunner
    {
        // 후처리기를 새로 추가했거나 결과물을 다시 만들어야 할 때 사용한다.
        [MenuItem("TableManager/Run Export Postprocessors")]
        public static void RunForAllTables()
        {
            var jsonDataPath = TableManagerConfig.JsonDataPath;
            var allTables = Directory.GetFiles(jsonDataPath, "*.json").Select(Path.GetFileNameWithoutExtension);

            Run(new TableExportContext(jsonDataPath, allTables));
        }

        public static void Run(TableExportContext context)
        {
            var changedTables = context.ChangedTables.Count == 0 ? "(none)" : string.Join(", ", context.ChangedTables);
            Debug.Log($"[{nameof(TableManager)}] Changed tables: {changedTables}");

            var postprocessors = TypeCache.GetTypesDerivedFrom<ITableExportPostprocessor>()
                .Where(type => !type.IsAbstract && !type.IsInterface)
                .Select(type => (ITableExportPostprocessor)Activator.CreateInstance(type))
                .OrderBy(postprocessor => postprocessor.CallbackOrder)
                .ThenBy(postprocessor => postprocessor.GetType().FullName, StringComparer.Ordinal)
                .ToList();

            foreach (var postprocessor in postprocessors)
            {
                Debug.Log($"[{nameof(TableManager)} Postprocess] {postprocessor.GetType().FullName}");

                try
                {
                    postprocessor.OnPostExport(context);
                }
                catch (Exception e)
                {
                    Debug.LogException(e);
                }
            }
        }
    }
}

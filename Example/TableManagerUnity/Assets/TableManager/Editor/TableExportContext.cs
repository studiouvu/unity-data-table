using System.Collections.Generic;

namespace TableManager
{
    public sealed class TableExportContext
    {
        private readonly HashSet<string> _changedTables;

        public string JsonDataPath { get; }

        public IReadOnlyCollection<string> ChangedTables => _changedTables;

        public TableExportContext(string jsonDataPath, IEnumerable<string> changedTables)
        {
            JsonDataPath = jsonDataPath;
            _changedTables = new HashSet<string>(changedTables);
        }

        public bool IsChanged(string tableName) => _changedTables.Contains(tableName);
    }
}

namespace TableManager
{
    /// <summary>
    /// 프로젝트 전용 후처리는 TableManager 폴더 밖에 구현해야 TableManager를 업데이트할 때 덮어써지지 않는다.
    /// </summary>
    public interface ITableExportPostprocessor
    {
        int CallbackOrder { get; }

        // Export 직후 Row 클래스 재컴파일(도메인 리로드)이 이어지므로 비동기 작업을 남기지 말고 동기로 끝낸다.
        void OnPostExport(TableExportContext context);
    }
}

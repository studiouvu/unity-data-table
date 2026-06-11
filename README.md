# UnityExcelTable

## 개요
---
Excel 파일의 데이터를 JSON으로 변환해 Unity에서 사용할 수 있게 해주는 툴입니다.

## 데이터 만들기
---
- Excel 파일을 `Assets/Excels` 폴더에 둡니다.
- 시트는 아래 구조를 따릅니다.
  - 1행: 컬럼 타입 (`string`, `int`, `bool` 등)
  - 2행: 컬럼 설명 (생성되는 클래스에 주석으로 들어갑니다)
  - 3행: 필드 이름
  - 4행부터: 데이터 (B열이 비어 있는 행부터는 읽지 않습니다)
- `#` 표시로 export 대상에서 제외할 수 있습니다.
  - 시트 이름에 `#` → 시트 전체 제외
  - 1행(타입 행) 셀에 `#` → 해당 컬럼 제외 (작업용 메모/미리보기 컬럼)
  - 행의 첫 칸(A열)에 `#` → 해당 행 제외
- 메뉴에서 `TableManager > Export`(Ctrl+F8)를 실행하면 `Assets/Jsons`에 JSON이, `Assets/TableManager/Class`에 Row 클래스가 생성됩니다.
- 다른 Excel 파일을 참조하는 수식(XLOOKUP 등)이 export 컬럼에 있으면 저장된 값이 원본 파일과 일치하는지 검증하며, 오래된 값이면 변환이 실패합니다. 해당 파일을 Excel에서 열어 저장한 뒤 다시 Export하면 됩니다.

## 사용법
---
- 앱 초기 진입점에서 `LocalDb.Init();`를 호출합니다.
- 아래와 같이 데이터를 불러와 사용합니다.

```cs
var itemDataRow = LocalDb.Get<DataItemRow>(id);

var name = itemDataRow.name;
var age = itemDataRow.age;

var testDataList = LocalDb.GetEnumerable<DataTestRow>();

foreach (var dataRow in testDataList)
{
    var testDataId = dataRow.id;
    var testValue = dataRow.value;
}
```

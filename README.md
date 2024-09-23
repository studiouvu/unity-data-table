# UnityExcelTable

## 개요
---
Unity 엔진에서 Excel파일, GoogleSheet의 데이터를 내려받아 사용 가능하게 해주는 툴입니다.

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
    var testDataId = itemDataRow.id;
    var testValue = itemDataRow.value;
}
```

- 작성중

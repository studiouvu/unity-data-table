using System.Collections;
using TableManager;
using UnityEngine;

public class LocalDbTest : MonoBehaviour
{
    private IEnumerator Start()
    {
        LocalDb.Init();

        yield return new WaitUntil(() => LocalDb.HasInit);

        var itemDataRow = LocalDb.Get<DataItemRow>("test");
        Debug.Log($"LocalDbTest.Start - DataItemRow {itemDataRow.id} {itemDataRow.price} {itemDataRow.currencyType}");

        var stringBuilder = new System.Text.StringBuilder();
        
        foreach (var testDataRow in LocalDb.GetEnumerable<DataCurrencyRow>())
        {
            stringBuilder.Clear();
            
            foreach (var variable in testDataRow.testArray)
                stringBuilder.Append($"{variable.ToString()} ");

            Debug.Log($"LocalDbTest.Start - DataCurrencyRow id : {testDataRow.id}\nnameId : {testDataRow.nameId}" +
                      $"\ntestArray : {stringBuilder}");
        }
    }
}

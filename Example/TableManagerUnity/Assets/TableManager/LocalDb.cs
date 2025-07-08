using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using Cysharp.Threading.Tasks;
using Newtonsoft.Json;
using UnityEngine;
namespace TableManager
{
    public interface IRow
    {
        string id { get; }
    }

    public interface IIndex
    {
        int index { get; }
    }

    public static class LocalDb
    {
        private static Dictionary<Type, Dictionary<string, object>> Dictionary = new();
        private static Dictionary<Type, Dictionary<int, object>> IndexDictionary = new();

        public static UniTask.Awaiter Init()
        {
            Dictionary.Clear();
            IndexDictionary.Clear();
            return InitAsync().GetAwaiter();
        }

        private static async UniTask InitAsync()
        {
            var types = Assembly
                .GetExecutingAssembly()
                .GetTypes()
                .Where(_ => typeof(IRow).IsAssignableFrom(_))
                .Where(_ => _.IsAbstract == false);

            foreach (var type in types)
            {
                try
                {
                    var typeName = type
                        .ToString()
                        .Replace($"{nameof(TableManager)}.", "")
                        .Replace("Row", "");

                    var json = $"Jsons/{typeName}.json".Load<TextAsset>().text;

                    var jsonDictionary = JsonConvert.DeserializeObject<Dictionary<int, List<string>>>(json);

                    Dictionary.Add(type, new Dictionary<string, object>());
                    IndexDictionary.Add(type, new Dictionary<int, object>());

                    for (var i = 4; i <= jsonDictionary.Count; i++)
                    {
                        if (jsonDictionary[i].Count == 0)
                            continue;
                        if (jsonDictionary[i][0] == "#" || string.IsNullOrEmpty(jsonDictionary[i][1]))
                            continue;

                        var instance = Activator.CreateInstance(type);

                        var tableFields = type
                            .GetProperties()
                            .ToDictionary(info => info.Name, info => info);

                        var isArray = false;
                        var arrayFieldName = string.Empty;
                        var arrayValueList = new LinkedList<string>();

                        for (var c = 0; c < jsonDictionary[1].Count; c++)
                        {
                            var fieldType = jsonDictionary[1][c];

                            if (fieldType == "#")
                                continue;

                            var fieldName = jsonDictionary[3][c];
                            var value = jsonDictionary[i][c];

                            if (fieldType.Contains("[]"))
                            {
                                if (!string.IsNullOrEmpty(fieldName))
                                {
                                    isArray = true;
                                    arrayFieldName = fieldName.Replace("[]", "");
                                    arrayValueList.Clear();
                                    if (!string.IsNullOrEmpty(value))
                                        arrayValueList.AddLast(value);
                                }
                                else
                                {
                                    if (!string.IsNullOrEmpty(value))
                                    {
                                        arrayValueList.AddLast(value);
                                    }
                                    else
                                    {
                                        ApplyValueArray(tableFields[arrayFieldName], instance, arrayValueList.ToArray());
                                        isArray = false;
                                    }
                                }
                            }
                            else
                            {
                                if (isArray)
                                {
                                    ApplyValueArray(tableFields[arrayFieldName], instance, arrayValueList.ToArray());
                                    isArray = false;
                                }

                                ApplyValue(tableFields[fieldName], instance, value);
                            }
                        }

                        var id = (instance as IRow)?.id;

                        Dictionary[type].TryAdd(id, instance);

                        if (instance is IIndex indexRow)
                            IndexDictionary[type].TryAdd(indexRow.index, instance);
                    }
                }
                catch (Exception e)
                {
                    Debug.LogError($"{type} {e.ToString()}");
                    throw;
                }
                await UniTask.Yield();
            }
        }

        public static T Get<T>(string id) where T : class, IRow, new()
        {
            return Dictionary[typeof(T)][id] as T;
        }

        public static T GetToIndex<T>(int index) where T : class, IRow, new()
        {
            return IndexDictionary[typeof(T)][index] as T;
        }

        public static T TryGet<T>(string id) where T : class, IRow, new()
        {
            if (!Dictionary[typeof(T)].ContainsKey(id))
                return null;

            return Dictionary[typeof(T)][id] as T;
        }

        public static IEnumerable<T> GetEnumerable<T>() where T : class, IRow, new()
        {
            return Dictionary[typeof(T)].Select(pair => pair.Value as T);
        }

        public static Dictionary<string, T> GetDictionary<T>() where T : class, IRow, new()
        {
            return Dictionary[typeof(T)].ToDictionary(pair => pair.Key, pair => pair.Value as T);
        }

        private static void ApplyValue(PropertyInfo field, object instance, string rvalue)
        {
            if (field.PropertyType == typeof(int))
            {
                field.SetValue(instance, int.Parse(rvalue));
            }
            else if (field.PropertyType == typeof(float))
            {
                field.SetValue(instance, float.Parse(rvalue));
            }
            else if (field.PropertyType == typeof(string))
            {
                field.SetValue(instance, rvalue);
            }
            else if (field.PropertyType == typeof(double))
            {
                field.SetValue(instance, double.Parse(rvalue));
            }
            else if (field.PropertyType == typeof(long))
            {
                field.SetValue(instance, long.Parse(rvalue));
            }
            else if (field.PropertyType == typeof(bool))
            {
                field.SetValue(instance, bool.Parse(rvalue));
            }
            else
            {
                try
                {
                    field.SetValue(instance, Enum.Parse(field.PropertyType, rvalue));
                }
                catch (Exception e)
                {
                    Debug.LogError($"e");
                    throw new ArgumentOutOfRangeException(
                        $"string => {field.PropertyType} {rvalue} 에 대한 바인딩 정의가 없습니다.");
                }
            }
        }

        private static void ApplyValueArray(PropertyInfo field, object instance, string[] values)
        {
            if (field.PropertyType == typeof(int[]))
            {
                field.SetValue(instance, values.Select(int.Parse).ToArray());
            }
            else if (field.PropertyType == typeof(float[]))
            {
                field.SetValue(instance, values.Select(float.Parse).ToArray());
            }
            else if (field.PropertyType == typeof(string[]))
            {
                field.SetValue(instance, values);
            }
            else if (field.PropertyType == typeof(double[]))
            {
                field.SetValue(instance, values.Select(double.Parse).ToArray());
            }
            else if (field.PropertyType == typeof(long[]))
            {
                field.SetValue(instance, values.Select(long.Parse).ToArray());
            }
            else if (field.PropertyType == typeof(bool[]))
            {
                field.SetValue(instance, values.Select(bool.Parse).ToArray());
            }
            else
            {
                throw new ArgumentOutOfRangeException(
                    $"string => {field.PropertyType} {values} 에 대한 바인딩 정의가 없습니다.");
            }
        }
    }
}

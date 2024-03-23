using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using Newtonsoft.Json;
using UnityEngine;
namespace TableManager
{
    public interface IRow
    {
        string id { get; }
    }

    public static class LocalDb
    {
        private static Dictionary<Type, Dictionary<string, object>> Dictionary = new();

        static LocalDb()
        {
            Init();
        }

        public static void Init()
        {
            Dictionary.Clear();

            var types = Assembly
                .GetExecutingAssembly()
                .GetTypes()
                .Where(_ => typeof(IRow).IsAssignableFrom(_))
                .Where(_ => _.IsAbstract == false);

            foreach (var type in types)
            {
                var typeName = type
                    .ToString()
                    .Replace($"{nameof(TableManager)}.", "")
                    .Replace("Row", "");

                var json = $"Jsons/{typeName}.json".Load<TextAsset>().text;

                var jsonDictionary = JsonConvert.DeserializeObject<Dictionary<int, List<string>>>(json);

                Dictionary.Add(type, new Dictionary<string, object>());

                for (var i = 4; i <= jsonDictionary.Count; i++)
                {
                    if (jsonDictionary[i][0] == "#" || string.IsNullOrEmpty(jsonDictionary[i][1]))
                        continue;

                    var instance = Activator.CreateInstance(type);

                    var tableFields = type
                        .GetProperties()
                        .ToDictionary(info => info.Name, info => info);

                    for (var c = 0; c < jsonDictionary[1].Count; c++)
                    {
                        if (string.IsNullOrEmpty(jsonDictionary[1][c]) || jsonDictionary[1][c] == "#")
                            continue;

                        var fieldName = jsonDictionary[3][c];
                        ApplyValue(tableFields[fieldName], instance, jsonDictionary[i][c]);
                    }

                    Dictionary[type].Add((instance as IRow).id, instance);
                }
            }
        }

        public static T Get<T>(string id) where T : class, IRow, new()
        {
            return Dictionary[typeof(T)][id] as T;
        }

        public static IEnumerable<T> GetEnumerable<T>() where T : class, IRow, new()
        {
            return Dictionary[typeof(T)].Select(pair => pair.Value as T);
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
            else
            {
                try
                {
                    field.SetValue(instance, Enum.Parse(field.PropertyType, rvalue));
                }
                catch (Exception e)
                {
                    throw new ArgumentOutOfRangeException(
                        $"string => {field.PropertyType} {rvalue} 에 대한 바인딩 정의가 없습니다.");
                }
            }
        }

    }
}

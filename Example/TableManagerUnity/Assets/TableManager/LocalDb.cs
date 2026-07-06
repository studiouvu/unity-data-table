using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Linq.Expressions;
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
        public static bool HasInit { get; private set; }

        private static readonly Dictionary<Type, Dictionary<string, object>> Dictionary = new();
        private static readonly Dictionary<Type, Dictionary<int, object>> IndexDictionary = new();
        private static bool isRunning;

        public static void Init()
        {
            if (isRunning)
                return;

            isRunning = true;

            Dictionary.Clear();
            IndexDictionary.Clear();
            InitAsync().Forget();
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

                var typeName = type
                    .ToString()
                    .Replace($"{nameof(TableManager)}.", "")
                    .Replace("Row", "");

                var json = $"Jsons/{typeName}.json".Load<TextAsset>().text;

                var rows = JsonConvert.DeserializeObject<List<List<string>>>(json);

                Dictionary.Add(type, new Dictionary<string, object>());
                IndexDictionary.Add(type, new Dictionary<int, object>());

                var tableFields = type
                    .GetProperties()
                    .ToDictionary(info => info.Name, info => info);

                // 0~2행은 헤더(타입/설명/이름), 3행부터 데이터
                for (var y = 3; y < rows.Count; y++)
                {
                    if (rows[y].Count == 0)
                        continue;
                    if (rows[y][0] == "#" || string.IsNullOrEmpty(rows[y][1]))
                        continue;

                    var instance = Activator.CreateInstance(type);

                    var arrayFieldName = string.Empty;
                    var arrayValueList = new LinkedList<string>();

                    for (var x = 1; x < rows[0].Count; x++)
                    {
                        var fieldType = rows[0][x];
                        var nextFieldType = (x + 1 < rows[0].Count) ? rows[0][x + 1] : null;

                        if (fieldType == "#")
                            continue;

                        var fieldName = rows[2][x];
                        var value = rows[y][x];

                        if (fieldType.Contains("[]"))
                        {
                            if (!string.IsNullOrEmpty(fieldName))
                            {
                                arrayFieldName = fieldName.Replace("[]", "");
                                arrayValueList.Clear();
                                if (!string.IsNullOrEmpty(value))
                                    arrayValueList.AddLast(value);
                            }
                            else
                            {
                                if (!string.IsNullOrEmpty(value))
                                    arrayValueList.AddLast(value);

                                if (nextFieldType != "[]")
                                    ApplyValueArray(tableFields[arrayFieldName], instance, arrayValueList.ToArray());
                            }
                        }
                        else
                        {
                            if (string.IsNullOrEmpty(fieldName))
                            {
                                Debug.LogError($"fieldName is null. type: {type}, row: {y + 1}, column: {x}");
                                continue;
                            }

                            ApplyValue(tableFields[fieldName], instance, value);
                        }
                    }

                    var id = (instance as IRow)?.id;

                    Dictionary[type].TryAdd(id, instance);

                    if (instance is IIndex indexRow)
                        IndexDictionary[type].TryAdd(indexRow.index, instance);
                }

                await UniTask.Yield();
            }
            HasInit = true;
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

        // 셀마다 reflection(SetValue)과 타입 분기를 반복하지 않도록,
        // 프로퍼티별 파싱+대입 델리게이트를 최초 사용 시 만들어 캐싱한다.
        private static readonly Dictionary<PropertyInfo, Action<object, string>> ValueSetters = new();
        private static readonly Dictionary<PropertyInfo, Action<object, string[]>> ArraySetters = new();

        private static void ApplyValue(PropertyInfo field, object instance, string rvalue)
        {
            if (!ValueSetters.TryGetValue(field, out var apply))
                ValueSetters[field] = apply = CreateValueSetter(field);

            apply(instance, rvalue);
        }

        private static void ApplyValueArray(PropertyInfo field, object instance, string[] values)
        {
            if (!ArraySetters.TryGetValue(field, out var apply))
                ArraySetters[field] = apply = CreateArraySetter(field);

            apply(instance, values);
        }

        private static Action<object, object> CompileSetter(PropertyInfo field)
        {
            var instance = Expression.Parameter(typeof(object), "instance");
            var value = Expression.Parameter(typeof(object), "value");
            var assign = Expression.Assign(
                Expression.Property(Expression.Convert(instance, field.DeclaringType), field),
                Expression.Convert(value, field.PropertyType));
            return Expression.Lambda<Action<object, object>>(assign, instance, value).Compile();
        }

        private static Action<object, string> CreateValueSetter(PropertyInfo field)
        {
            var set = CompileSetter(field);
            var type = field.PropertyType;

            if (type == typeof(int))
                return (instance, value) => set(instance, int.Parse(value, CultureInfo.InvariantCulture));
            if (type == typeof(float))
                return (instance, value) => set(instance, float.Parse(value, CultureInfo.InvariantCulture));
            if (type == typeof(string))
                return (instance, value) => set(instance, value);
            if (type == typeof(double))
                return (instance, value) => set(instance, double.Parse(value, CultureInfo.InvariantCulture));
            if (type == typeof(long))
                return (instance, value) => set(instance, long.Parse(value, CultureInfo.InvariantCulture));
            if (type == typeof(bool))
                return (instance, value) => set(instance, bool.Parse(value));

            return (instance, value) =>
            {
                try
                {
                    set(instance, Enum.Parse(type, value));
                }
                catch (Exception)
                {
                    throw new ArgumentOutOfRangeException(
                        $"string => {type} {value} 에 대한 바인딩 정의가 없습니다.");
                }
            };
        }

        private static Action<object, string[]> CreateArraySetter(PropertyInfo field)
        {
            var set = CompileSetter(field);
            var type = field.PropertyType;

            if (type == typeof(int[]))
                return (instance, values) => set(instance, values.Select(v => int.Parse(v, CultureInfo.InvariantCulture)).ToArray());
            if (type == typeof(float[]))
                return (instance, values) => set(instance, values.Select(v => float.Parse(v, CultureInfo.InvariantCulture)).ToArray());
            if (type == typeof(string[]))
                return (instance, values) => set(instance, values);
            if (type == typeof(double[]))
                return (instance, values) => set(instance, values.Select(v => double.Parse(v, CultureInfo.InvariantCulture)).ToArray());
            if (type == typeof(long[]))
                return (instance, values) => set(instance, values.Select(v => long.Parse(v, CultureInfo.InvariantCulture)).ToArray());
            if (type == typeof(bool[]))
                return (instance, values) => set(instance, values.Select(bool.Parse).ToArray());

            throw new ArgumentOutOfRangeException(
                $"string => {type} 에 대한 바인딩 정의가 없습니다.");
        }
    }
}

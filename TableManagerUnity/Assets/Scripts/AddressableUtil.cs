using System;
using System.Collections.Generic;
using System.IO;
using UnityEngine;
using UnityEngine.AddressableAssets;
using Debug = UnityEngine.Debug;
using Object = UnityEngine.Object;
public static class AddressableUtil
{
    private static readonly Dictionary<string, Object> AssetCache = new();
    
    private static string ConvertKeyToLower(string key)
    {
        var lower = $"{Path.GetDirectoryName(key)}/{Path.GetFileNameWithoutExtension(key)}".Replace('\\', '/').ToLower();
        if (lower.StartsWith("/"))
            lower = lower.Remove(0, 1);
        return lower;
    }

    public static T LoadAsset<T>(string path) where T : Object
    {
        if (AssetCache.TryGetValue(path, out var asset)) 
            return asset as T;
        
        try
        {
            var a = Addressables
                .LoadAssetAsync<T>(path)
                .WaitForCompletion();
            AssetCache[path] = a;
            return a;
        }
        catch (Exception e)
        {
            Debug.LogError(e);
        }
        
        return default(T);
    }

    public static T Instantiate<T>(string path, Transform parent = default) where T : Object
    {
        try
        {
            var a = Addressables
                .InstantiateAsync(path, parent)
                .WaitForCompletion()
                .GetComponent<T>();
            return a;
        }
        catch (Exception e)
        {
            Debug.LogError(e);
        }
        return default(T);
    }

    public static T Load<T>(this string path) where T : Object
        => LoadAsset<T>(path);
    
    public static T LoadPrefab<T>(this string path) where T : Object
        => LoadAsset<GameObject>(path).GetComponent<T>();
    
    public static void UnloadAsset(this Object asset)
    {
        Addressables.Release(asset);
    }
}
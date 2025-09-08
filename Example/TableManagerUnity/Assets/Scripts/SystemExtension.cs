using System;
public static class SystemExtension
{
    public static bool IsNullOrEmpty(this string value) => string.IsNullOrEmpty(value);
    public static bool IsNotEmpty(this string value) => !string.IsNullOrEmpty(value);

    public static int AddFlag(this int value, int flag) => value | flag;
    public static int RemoveFlag(this int value, int flag) => value & ~flag;
    public static bool HasFlag(this int value, int flag) => (value & flag) == flag;
        
    public static int FloorToInt(this float value) => (int)Math.Floor(value);

    public static int RoundToInt(this float value) => (int)Math.Round(value);

    public static int CeilToInt(this float value) => (int)Math.Ceiling(value);
}
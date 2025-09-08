using System;
using UnityEngine;
namespace TableManager.Editor
{
    [Serializable]
    public class GoogleSheetConfig : ScriptableObject
    {
        public string clientId;
        public string clientSecret;
        public string apiKey;
        public string sheetId;
    }
}

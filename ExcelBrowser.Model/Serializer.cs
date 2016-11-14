using Newtonsoft.Json;

namespace ExcelBrowser.Model {

    internal static class Serializer {
        
        public static string Serialize<T>(T obj) {
            return JsonConvert.SerializeObject(obj, Formatting.Indented)
                .Replace("\"", ""); //Pretty JSON
        }
    }
}

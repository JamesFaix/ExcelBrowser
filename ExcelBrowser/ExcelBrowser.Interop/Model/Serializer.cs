using Newtonsoft.Json;

namespace ExcelBrowser.Model {

    internal static class Serializer {

        public static string Serialize<T>(T obj) =>
            JsonConvert.SerializeObject(obj, Formatting.Indented)
                .Replace("\"", ""); //Pretty JSON
               
    }
}

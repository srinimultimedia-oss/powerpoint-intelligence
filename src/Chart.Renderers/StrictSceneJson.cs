using Newtonsoft.Json.Linq;
using PowerPointIntelligence.Contracts;
namespace PowerPointIntelligence.Rendering
{
    internal static class StrictSceneJson
    {
        internal static JToken Parse(string json) { return BoundedJson.Parse(json, SceneCodec.MaxCharacters, "scene"); }
    }
    internal static class SceneText
    {
        internal static bool Valid(string value, int maximum) { return ContractText.Valid(value, maximum); }
        internal static bool NonWhitespace(string value) { return ContractText.NonWhitespace(value); }
    }
}
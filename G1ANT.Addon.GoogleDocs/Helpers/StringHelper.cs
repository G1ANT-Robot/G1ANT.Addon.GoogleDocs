using G1ANT.Language;

namespace G1ANT.Addon.GoogleDocs.Helpers
{
    public static class StringHelper
    {
        public static bool IsNullOrEmpty(this string text) => string.IsNullOrEmpty(text);

        public static bool IsNullOrEmpty(this TextStructure text) => string.IsNullOrEmpty(text?.Value);
    }
}

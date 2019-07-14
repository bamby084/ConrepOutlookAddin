
namespace ConrepOutlookAddin.ExtensionMethods
{
    public static class UrlExtensions
    {
        public static string EnsureStartsWithHttps(this string url)
        {
            if (url.StartsWith("http://") || url.StartsWith("https://"))
                return url;

            return "https://" + url;
        }
    }
}

using System.DirectoryServices;

namespace Microsoft.Internal.MSContactImporter
{
    internal static class ExtensionMethods
    {
        internal static string GetProperty(this DirectoryEntry directoryEntry, string propertyName)
        {
            string result;
            try
            {
                object value = directoryEntry.Properties[propertyName].Value;
                if (value == null)
                {
                    result = string.Empty;
                }
                else
                {
                    result = value.ToString();
                }
            }
            catch
            {
                result = string.Empty;
            }
            return result;
        }
    }
}
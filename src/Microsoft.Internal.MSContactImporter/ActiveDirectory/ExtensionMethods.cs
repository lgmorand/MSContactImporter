using System.Collections.Generic;
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

        //Get members of a distribution list
        internal static List<string> GetMembers(this DirectoryEntry directoryEntry)
        {
            List<string> members;
            try
            {
                PropertyValueCollection values = directoryEntry.Properties["member"];
                if (values.Count == 0)
                {
                    members = null;
                }
                else
                {
                    members = new List<string>();
                    foreach (var member in values)
                    {
                        members.Add(member.ToString());
                    }
                    
                }
            }
            catch
            {
                members = null;
            }
            return members;
        }
    }
}
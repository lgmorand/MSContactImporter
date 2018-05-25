using System;
using System.Collections.Generic;
using System.DirectoryServices;
using System.Linq;
using System.Security.Principal;
using System.Xml.Linq;
using Microsoft.Internal.MSContactImporter.Properties;

namespace Microsoft.Internal.MSContactImporter
{
    internal class ADUtils
    {
        internal Dictionary<string, MSFTee> Managers = new Dictionary<string, MSFTee>();
        internal Dictionary<string, MSFTee> MSFTees = new Dictionary<string, MSFTee>();
        private readonly string _login;
        private readonly string _password;

        public ADUtils(string login, string password)
        {
            _login = login;
            _password = password;
        }

        internal string GetDistinguishedName(string logon)
        {
            string result2;
            using (DirectorySearcher directorySearcher = new DirectorySearcher())
            {
                directorySearcher.SearchRoot = GetDirectoryEntry(Settings.Default.RootDirectoryEntry);
                directorySearcher.Filter = string.Format("(sAMAccountName={0})", logon);
                directorySearcher.PropertiesToLoad.Add("distinguishedName");
                SearchResult result = directorySearcher.FindOne();
                if (result == null)
                {
                    Console.WriteLine(string.Format("GetDistinguishedName(logon={0})", logon));
                    result2 = null;
                }
                else
                {
                    result2 = result.Properties["distinguishedName"][0].ToString();
                }
            }
            return result2;
        }

        internal void LoadActiveDirectory(string alias, bool recursive, int level, Func<bool> cancel, Action<string> label)
        {
            string distinguishedName = this.GetDistinguishedName(alias);
            this.LoadTree(distinguishedName, null, recursive, level, cancel, label);
        }

        internal void LoadFromXml(string filePath)
        {
            this.MSFTees = new Dictionary<string, MSFTee>();
            XElement.Load(filePath).Elements("Microsoftee").ToList<XElement>().ForEach(delegate (XElement xe)
            {
                MSFTee msftee = MSFTee.FromXElement(xe);
                this.MSFTees.Add(msftee.DistinguishedName, msftee);
            });
            this.MSFTees.ToList<KeyValuePair<string, MSFTee>>().ForEach(delegate (KeyValuePair<string, MSFTee> msftee)
            {
                msftee.Value.Manager = (from msftee2 in this.MSFTees.Values
                                        where msftee2.Alias == msftee.Value.ManagerAlias
                                        select msftee2).FirstOrDefault<MSFTee>();
            });
        }

        internal MSFTee LoadIndependantMSFTee(string distinguishedName, bool recurse)
        {
            if (distinguishedName == null)
            {
                return null;
            }

            string name = string.Format(Settings.Default.DistinguishedNameFormat, distinguishedName);
            if (DoesUserExists(name))
            {
                DirectoryEntry directory = GetDirectoryEntry(name);
                string domaineName = "UNKNOWN";
                byte[] objectsid = (byte[])directory.Properties["objectsid"].Value;
                if (objectsid != null)
                {
                    try
                    {
                        domaineName = ((NTAccount)new SecurityIdentifier(objectsid, 0).Translate(typeof(NTAccount))).Value.ToLower().Split(new char[]
                        {
                        '\\'
                        })[0];
                    }
                    catch { } //Catch if caller not in a domain
                }

                MSFTee msftee = new MSFTee
                {
                    ID = Guid.NewGuid(),
                    Alias = directory.GetProperty("sAMAccountName").ToLower(),
                    Domain = domaineName,
                    FullLogon = string.Format("{0}\\{1}", domaineName, directory.GetProperty("sAMAccountName").ToLower()),
                    Email = directory.GetProperty("mail").ToLower(),
                    FirstName = directory.GetProperty("givenName"),
                    LastName = directory.GetProperty("sn"),
                    FullName = directory.GetProperty("displayName"),
                    DistinguishedName = directory.GetProperty("distinguishedName"),
                    StreetAddress = directory.GetProperty("streetAddress"),
                    PostalCode = directory.GetProperty("postalCode"),
                    City = directory.GetProperty("l"),
                    State = directory.GetProperty("st"),
                    Country = directory.GetProperty("co"),
                    TelephoneNumber = this.CleanNumber(directory.GetProperty("telephoneNumber")),
                    HomePhone = this.CleanNumber(directory.GetProperty("homePhone")),
                    MobilePhone = this.CleanNumber(directory.GetProperty("mobile")),
                    Fax = this.CleanNumber(directory.GetProperty("facsimileTelephoneNumber")),
                    Title = directory.GetProperty("title"),
                    Department = directory.GetProperty("department"),
                    PhysicalDeliveryOfficeName = directory.GetProperty("physicalDeliveryOfficeName"),
                    Company = directory.GetProperty("company")
                };

                if (recurse)
                {
                    msftee.Manager = this.LoadIndependantMSFTeeManager(distinguishedName, false);
                    msftee.ManagerAlias = msftee.Manager.Alias;
                }
                return msftee;
            }
            else
            {
                Logger.LogMessageToConsole(string.Format("No user {0} found ", distinguishedName));
            }
            return null;
        }

        internal MSFTee LoadIndependantMSFTeeManager(string distinguishedName, bool recurse)
        {
            MSFTee result2 = null;
            using (DirectorySearcher directorySearcher = new DirectorySearcher())
            {
                directorySearcher.SearchRoot = GetDirectoryEntry(Settings.Default.RootDirectoryEntry);
                directorySearcher.Filter = string.Format("(distinguishedName={0})", distinguishedName);
                directorySearcher.PropertiesToLoad.Add("manager");
                SearchResult result = directorySearcher.FindOne();
                if (result != null)
                {
                    if (result.Properties["manager"].Count > 0)
                    {
                        string managerDistinguishedName = result.Properties["manager"][0].ToString();
                        if (this.Managers.ContainsKey(managerDistinguishedName))
                        {
                            result2 = this.Managers[managerDistinguishedName];
                        }
                        else
                        {
                            MSFTee managerMSFTee = this.LoadIndependantMSFTee(managerDistinguishedName, recurse);
                            this.Managers.Add(managerDistinguishedName, managerMSFTee);
                            result2 = managerMSFTee;
                        }
                    }
                }
                else
                {
                    result2 = null;
                }
            }
            return result2;
        }

        internal void SaveToXml(string filePath)
        {
            if (this.MSFTees != null)
            {
                new XElement("Microsoftees", MSFTees.Select(m => m.Value.ToXml())).Save(filePath);
            }
        }

        internal bool TestConnection(string aliasToTest)
        {
            try
            {
                Logger.LogMessageToConsole("== Testing connection to Active Directory ==");
                using (DirectorySearcher directorySearcher = new DirectorySearcher())
                {
                    directorySearcher.SearchRoot = GetDirectoryEntry(Settings.Default.RootDirectoryEntry);
                    directorySearcher.Filter = string.Format("(sAMAccountName={0})", aliasToTest);
                    directorySearcher.PropertiesToLoad.Add("distinguishedName");
                    SearchResult result = directorySearcher.FindOne();
                    if (result != null)
                    {
                        Logger.LogMessageToConsole("Connection to AD SUCCESSFUL");
                        return true;
                    }
                    else
                    {
                        Logger.LogMessageToConsole("Connection to AD FAILED");
                        Logger.LogMessageToConsole("Please ensure that credentials are correct and VPN is connected");
                        return false;
                    }
                }
            }
            catch (Exception)
            {
                Logger.LogMessageToConsole(string.Format("ERROR: cannot connect to AD ({0})", Settings.Default.RootDirectoryEntry));
                Logger.LogMessageToConsole("Please ensure that credentials are correct and VPN is connected");
                return false;
            }
        }

        private string CleanNumber(string phone)
        {
            if (!string.IsNullOrEmpty(phone))
            {
                string newPhone = phone.Replace(" ", string.Empty).Replace(".", string.Empty).Replace("-", string.Empty).Replace("(", string.Empty).Replace(")", string.Empty);
                if (newPhone.StartsWith("01"))
                {
                    newPhone = string.Format("+33{0}", newPhone.Substring(1));
                }
                if (newPhone.StartsWith("03"))
                {
                    newPhone = string.Format("+33{0}", newPhone.Substring(1));
                }
                if (newPhone.StartsWith("06"))
                {
                    newPhone = string.Format("+33{0}", newPhone.Substring(1));
                }
                if (newPhone.StartsWith("07"))
                {
                    newPhone = string.Format("+33{0}", newPhone.Substring(1));
                }
                if (newPhone.StartsWith("33"))
                {
                    newPhone = string.Format("+{0}", newPhone);
                }
                if (newPhone.StartsWith("0033"))
                {
                    newPhone = string.Format("+{0}", newPhone.Substring(2));
                }
                if (newPhone.StartsWith("66440"))
                {
                    newPhone = string.Format("+33{0}", newPhone);
                }
                if (newPhone.StartsWith("+33166440"))
                {
                    newPhone = newPhone.Replace("+33166440", "+3366440");
                }
                if (newPhone.StartsWith("+3316440"))
                {
                    newPhone = newPhone.Replace("+3316440", "+3366440");
                }
                if (newPhone.StartsWith("+330"))
                {
                    newPhone = string.Format("+33{0}", newPhone.Substring(4));
                }
                if ((newPhone.StartsWith("+336") || newPhone.StartsWith("+337")) && newPhone.Contains("X"))
                {
                    newPhone = newPhone.Replace("X", string.Empty).Substring(0, 12);
                }
                return newPhone;
            }
            return phone;
        }

        private bool DoesUserExists(string path)
        {
            DirectoryEntry directoryEntry = GetDirectoryEntry(path);

            bool exists = false;
            try
            {
                var tmp = directoryEntry.Guid;
                exists = true;
            }
            catch
            {
                Logger.LogMessageToConsole(string.Format("user {0} does not exist ", path));
                exists = false;
            }

            return exists;
        }

        private string Escape2DistinguishedName(string distinguishedName)
        {
            return distinguishedName.Replace("*", "\\2a").Replace("(", "\\28").Replace(")", "\\29").Replace("\\", "\\5c").Replace("/", "\\2f");
        }

        private string EscapeDistinguishedName(string distinguishedName)
        {
            int index = distinguishedName.IndexOf(",OU=");
            if (index == -1)
            {
                return distinguishedName;
            }
            return string.Join<char>("", distinguishedName.Take(index).ToArray<char>()).Replace("/", "\\/") + distinguishedName.Substring(index);
        }

        private DirectoryEntry GetDirectoryEntry(string path)
        {
            return new DirectoryEntry(path, _login, _password);
        }

        private void GetDirectReports(string distinguishedName, bool recursive, int level, Func<bool> cancel, Action<string> label)
        {
            if (cancel())
            {
                return;
            }

            using (DirectorySearcher directorySearcher = new DirectorySearcher())
            {
                directorySearcher.SearchRoot = GetDirectoryEntry(Settings.Default.RootDirectoryEntry);
                directorySearcher.Filter = string.Format("(distinguishedName={0})", this.Escape2DistinguishedName(distinguishedName));
                directorySearcher.ExtendedDN = ExtendedDN.None;
                directorySearcher.PropertiesToLoad.Add("directReports");
                SearchResult result = directorySearcher.FindOne();
                if (result != null)
                {
                    int directReportsNumber = result.Properties["directReports"].Count;

                    if (directReportsNumber == 0)
                    {
                        Logger.LogMessageToConsole(string.Format("No direct reports found for {0}", distinguishedName));
                    }
                    else
                    {
                        Logger.LogMessageToConsole(string.Format("{1} direct reports found for {0}", distinguishedName, directReportsNumber));
                    }

                    for (int index = 0; index < directReportsNumber; index++)
                    {
                        if (cancel())
                        {
                            break;
                        }
                        this.LoadTree(result.Properties["directReports"][index].ToString(), distinguishedName, recursive, level, cancel, label);
                    }
                }
                else
                {
                    Logger.LogMessageToConsole(string.Format("No info found in AD for {0}", distinguishedName));
                }
            }
        }

        private void LoadMSFTee(string distinguishedName, string manager, Func<bool> cancel, Action<string> label)
        {
            if (cancel())
            {
                return;
            }

            if (!this.MSFTees.ContainsKey(distinguishedName))
            {
                string name = string.Format(Settings.Default.DistinguishedNameFormat, this.EscapeDistinguishedName(distinguishedName));
                try
                {
                    if (DoesUserExists(name))
                    {
                        DirectoryEntry directory = GetDirectoryEntry(name);
                        string domaineName = "UNKNOWN";
                        byte[] objectsid = (byte[])directory.Properties["objectsid"].Value;
                        if (objectsid != null)
                        {
                            try
                            {
                                domaineName = ((NTAccount)new SecurityIdentifier(objectsid, 0).Translate(typeof(NTAccount))).Value.ToLower().Split(new char[]
                                {
                                '\\'
                                })[0];
                            }
                            catch { } //Catch if caller not in a domain
                        }
                        MSFTee managerObject = null;
                        if (!string.IsNullOrEmpty(manager))
                        {
                            if (this.MSFTees.ContainsKey(manager))
                            {
                                managerObject = this.MSFTees[manager];
                            }
                            else
                            {
                                managerObject = this.LoadIndependantMSFTee(manager, false);
                            }
                        }
                        label(directory.GetProperty("displayName"));

                        this.MSFTees.Add(directory.GetProperty("distinguishedName"), new MSFTee
                        {
                            ID = Guid.NewGuid(),
                            Alias = directory.GetProperty("sAMAccountName").ToLower(),
                            Domain = domaineName,
                            FullLogon = string.Format("{0}\\{1}", domaineName, directory.GetProperty("sAMAccountName").ToLower()),
                            Email = directory.GetProperty("mail").ToLower(),
                            FirstName = directory.GetProperty("givenName"),
                            LastName = directory.GetProperty("sn"),
                            FullName = directory.GetProperty("displayName"),
                            DistinguishedName = directory.GetProperty("distinguishedName"),
                            StreetAddress = directory.GetProperty("streetAddress"),
                            PostalCode = directory.GetProperty("postalCode"),
                            City = directory.GetProperty("l"),
                            State = directory.GetProperty("st"),
                            Country = directory.GetProperty("co"),
                            TelephoneNumber = this.CleanNumber(directory.GetProperty("telephoneNumber")),
                            HomePhone = this.CleanNumber(directory.GetProperty("homePhone")),
                            MobilePhone = this.CleanNumber(directory.GetProperty("mobile")),
                            Fax = this.CleanNumber(directory.GetProperty("facsimileTelephoneNumber")),
                            Title = directory.GetProperty("title"),
                            Department = directory.GetProperty("department"),
                            PhysicalDeliveryOfficeName = directory.GetProperty("physicalDeliveryOfficeName"),
                            Company = directory.GetProperty("company"),
                            Manager = managerObject,
                            ManagerAlias = ((managerObject == null) ? null : managerObject.Alias)
                        });
                    }
                }
                catch (Exception ex)
                {
                    Logger.LogMessageToConsole("Error while loading Msftee" + distinguishedName, ex);
                }
            }
        }

        private void LoadTree(string distinguishedName, string manager, bool recursive, int level, Func<bool> cancel, Action<string> label)
        {
            if (cancel())
            {
                return;
            }

            if (manager == null)
            {
                using (DirectorySearcher directorySearcher = new DirectorySearcher())
                {
                    directorySearcher.SearchRoot = GetDirectoryEntry(Settings.Default.RootDirectoryEntry);
                    directorySearcher.Filter = string.Format("(distinguishedName={0})", this.EscapeDistinguishedName(distinguishedName));
                    directorySearcher.PropertiesToLoad.Add("manager");
                    SearchResult result = directorySearcher.FindOne();
                    if (result != null && result.Properties["manager"].Count > 0)
                    {
                        manager = result.Properties["manager"][0].ToString();
                    }
                }
            }

            this.LoadMSFTee(distinguishedName, manager, cancel, label);
            if (recursive && (level > 0 || level == -1))
            {
                if (level != -1)
                {
                    level--;
                }
                this.GetDirectReports(distinguishedName, recursive, level, cancel, label);
            }
        }
    }
}
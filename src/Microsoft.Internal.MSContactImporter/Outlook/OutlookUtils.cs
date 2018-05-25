using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.Internal.MSContactImporter.Properties;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace Microsoft.Internal.MSContactImporter
{
    internal class OutlookUtils : IDisposable
    {
        private readonly Outlook.Application outlook;
        private readonly Outlook.MAPIFolder contactsFolder;
        private readonly Outlook.Items contactsItems;

        internal List<Outlook.ContactItem> Contacts;

        public OutlookUtils()
        {
            outlook = new Outlook.Application();

            //Searching for Contacts folder if Microsoft corp mailbox is not the default one
            Outlook.Folders folders = outlook.Session.Folders;
            foreach (Outlook.Folder f in folders)
            {
                if (f.Name.ToLower().EndsWith("@microsoft.com"))
                {
                    foreach (Outlook.Folder subf in f.Folders)
                    {
                        if (subf.Name.ToLower() == "contacts")
                        {
                            contactsFolder = outlook.Session.GetFolderFromID(subf.EntryID);
                            break;
                        }
                    }

                    if (contactsFolder != null) //if contacts folder was found
                        break;
                }
            }

            contactsItems = contactsFolder.Items;

            CreateCategory();
        }

        private void CreateCategory()
        {
            //checking if category exists, if not create it
            Outlook.NameSpace ns = outlook.GetNamespace("MAPI");
            bool categoryExists = false;

            foreach (Outlook.Category categorie in ns.Categories)
            {
                if (categorie.Name == Settings.Default.Categories)
                {
                    categoryExists = true;
                    break;
                }
            }

            if (!categoryExists)
                ns.Categories.Add(Settings.Default.Categories);
        }

        internal bool InsertOrUpdateContact(MSFTee msftee, GraphUtils graphUtils)
        {
            if (string.IsNullOrEmpty(msftee.Alias))
            {
                return false;
            }

            Outlook.ContactItem contact = this.Contacts.FirstOrDefault(c =>
            {
                Outlook.PropertyAccessor pa = c.PropertyAccessor;
                bool contactFound = pa.GetProperty(Settings.Default.ExtendedPropertySchema + Settings.Default.MsStaffId) == msftee.Alias;
                Marshal.ReleaseComObject(pa);
                return contactFound;
            });

            bool found = contact != null;
            if (!found)
            {
                contact = contactsItems.Add();
                Outlook.PropertyAccessor pa = contact.PropertyAccessor;
                pa.SetProperty(Settings.Default.ExtendedPropertySchema + Settings.Default.MsStaffId, msftee.Alias);
                Marshal.ReleaseComObject(pa);
            }
            contact.Categories = Settings.Default.Categories;

            contact.CompanyName = msftee.Company;
            contact.BusinessAddressCity = msftee.City;
            contact.BusinessAddressCountry = msftee.Country;
            contact.BusinessAddressPostalCode = msftee.PostalCode;
            contact.BusinessAddressState = msftee.State;
            contact.BusinessAddressStreet = msftee.StreetAddress;
            contact.Department = msftee.Department;
            contact.Email1Address = msftee.Email;

            if (!string.IsNullOrEmpty(msftee.Fax))
            {
                contact.BusinessFaxNumber = msftee.Fax;
            }
            contact.FirstName = msftee.FirstName;
            contact.FileAs = string.Format("MS: {0}", msftee.FullName);
            contact.FullName = msftee.FullName;
            if (!string.IsNullOrEmpty(msftee.HomePhone))
            {
                contact.HomeTelephoneNumber = msftee.HomePhone;
            }
            contact.LastName = msftee.LastName;
            if (msftee.Manager != null)
            {
                contact.ManagerName = string.Format("{0} [{1}]", msftee.Manager.FullName, msftee.Manager.Alias);
            }
            if (!string.IsNullOrEmpty(msftee.MobilePhone))
            {
                contact.MobileTelephoneNumber = msftee.MobilePhone;
            }
            contact.OfficeLocation = msftee.PhysicalDeliveryOfficeName;
            if (!string.IsNullOrEmpty(msftee.TelephoneNumber))
            {
                contact.BusinessTelephoneNumber = msftee.TelephoneNumber;
            }
            contact.JobTitle = msftee.Title;

            if (graphUtils != null) //we want to import user photo
            {
                try
                {
                    string picturePath = graphUtils.GetPhotoForUserAsync(msftee).Result;
                    if (!string.IsNullOrEmpty(picturePath))
                        contact.AddPicture(picturePath); //Not using async/await pattern to avoid changing all the rest of the calling code
                }
                catch { } //Simply ignoring the error, contact will have no picture
            }

            Logger.LogMessageToConsole("Saving changes back to Outlook...");
            if (found)
            {
                contact.Save();
                return false;
            }
            contact.Save();

            Marshal.ReleaseComObject(contact);
            return true;
        }

        internal bool IsAnImportedContact(Outlook.ContactItem contact)
        {
            Outlook.PropertyAccessor pa = contact.PropertyAccessor;
            bool isAnImportedContact = pa.GetProperty(Settings.Default.ExtendedPropertySchema + Settings.Default.MsStaffId) != null;
            Marshal.ReleaseComObject(pa);
            return isAnImportedContact;
        }

        internal void LoadContacts()
        {
            this.Contacts = new List<Outlook.ContactItem>();

            Outlook.ContactItem contact = contactsItems.Find("[Categories] = '" + Settings.Default.Categories + "'");

            if (contact != null)
                this.Contacts.Add(contact);

            do
            {
                contact = contactsItems.FindNext();
                if (contact != null)
                    this.Contacts.Add(contact);

            }
            while (contact != null);

            Logger.LogMessageToConsole(string.Format("{0} of your MS contacts found", this.Contacts.Count));
        }

        internal void UpdateContact(Outlook.ContactItem contact, ADUtils adUtils, GraphUtils graphUtils)
        {
            Outlook.PropertyAccessor pa = contact.PropertyAccessor;
            string logon = pa.GetProperty(Settings.Default.ExtendedPropertySchema + Settings.Default.MsStaffId);
            Marshal.ReleaseComObject(pa);

            MSFTee msftee = adUtils.LoadIndependantMSFTee(adUtils.GetDistinguishedName(logon), true);
            if (msftee != null)
            {
                contact.Categories = Settings.Default.Categories;
                contact.CompanyName = msftee.Company;
                contact.BusinessAddressCity = msftee.City;
                contact.BusinessAddressCountry = msftee.Country;
                contact.BusinessAddressPostalCode = msftee.PostalCode;
                contact.BusinessAddressState = msftee.State;
                contact.BusinessAddressStreet = msftee.StreetAddress;
                contact.Department = msftee.Department;
                contact.Email1Address = msftee.Email;

                if (!string.IsNullOrEmpty(msftee.Fax))
                {
                    contact.BusinessFaxNumber = msftee.Fax;
                }
                contact.FirstName = msftee.FirstName;
                contact.FileAs = string.Format("MS: {0}", msftee.FullName);
                contact.FullName = msftee.FullName;
                if (!string.IsNullOrEmpty(msftee.HomePhone))
                {
                    contact.HomeTelephoneNumber = msftee.HomePhone;
                }
                contact.LastName = msftee.LastName;
                if (msftee.Manager != null)
                {
                    contact.ManagerName = string.Format("{0} [{1}]", msftee.Manager.FullName, msftee.Manager.Alias);
                }
                if (!string.IsNullOrEmpty(msftee.MobilePhone))
                {
                    contact.MobileTelephoneNumber = msftee.MobilePhone;
                }
                contact.OfficeLocation = msftee.PhysicalDeliveryOfficeName;
                if (!string.IsNullOrEmpty(msftee.TelephoneNumber))
                {
                    contact.BusinessTelephoneNumber = msftee.TelephoneNumber;
                }
                contact.JobTitle = msftee.Title;

                if (graphUtils != null) //we want to import user photo
                {
                    try
                    {
                        string picturePath = graphUtils.GetPhotoForUserAsync(msftee).Result;
                        if (!string.IsNullOrEmpty(picturePath))
                        {
                            contact.RemovePicture();
                            contact.AddPicture(picturePath); //Not using async/await pattern to avoid changing all the rest of the calling code
                        }
                    }
                    catch { } //Simply ignoring the error, contact will have no picture
                }
                Logger.LogMessageToConsole("Saving changes back to Outlook...");
                contact.Save();
            }
        }

        /// <summary>
        /// Test the connection to Outlook
        /// </summary>
        /// <param name="useAutodiscoverMode"></param>
        /// <param name="onCorpNetwork"></param>
        /// <param name="login"></param>
        /// <param name="password"></param>
        /// <param name="url"></param>
        /// <returns></returns>
        internal bool TestConnection()
        {
            try
            {
                Logger.LogMessageToConsole("== Testing connection to Outlook ==");

                if (contactsFolder == null)
                {
                    Logger.LogMessageToConsole("ERROR: Unable to find your Contacts folder in Outlook");
                    return false;
                }
            }
            catch (Exception ex)
            {
                Logger.LogMessageToConsole("ERROR: " + ex.Message);
                return false;
            }

            Logger.LogMessageToConsole("Connection to Outlook SUCCESSFUL");
            return true;
        }

        public void Dispose()
        {
            if (this.Contacts != null)
            {
                foreach (Outlook.ContactItem contact in this.Contacts)
                {
                    Marshal.ReleaseComObject(contact);
                }
            }
            Marshal.ReleaseComObject(contactsItems);
            Marshal.ReleaseComObject(contactsFolder);
            Marshal.ReleaseComObject(outlook);
        }
    }
}
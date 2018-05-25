using System;
using System.Xml.Linq;

namespace Microsoft.Internal.MSContactImporter
{
    internal class MSFTee
    {
        internal Guid ID
        {
            get;
            set;
        }

        internal string Domain
        {
            get;
            set;
        }

        internal string Alias
        {
            get;
            set;
        }

        internal string FullLogon
        {
            get;
            set;
        }

        internal string FullName
        {
            get;
            set;
        }

        internal string Email
        {
            get;
            set;
        }

        internal string FirstName
        {
            get;
            set;
        }

        internal string LastName
        {
            get;
            set;
        }

        internal string Title
        {
            get;
            set;
        }

        internal string Department
        {
            get;
            set;
        }

        internal string DistinguishedName
        {
            get;
            set;
        }

        internal string StreetAddress
        {
            get;
            set;
        }

        internal string PostalCode
        {
            get;
            set;
        }

        internal string City
        {
            get;
            set;
        }

        internal string State
        {
            get;
            set;
        }

        internal string Country
        {
            get;
            set;
        }

        internal string TelephoneNumber
        {
            get;
            set;
        }

        internal string HomePhone
        {
            get;
            set;
        }

        internal string MobilePhone
        {
            get;
            set;
        }

        internal string Fax
        {
            get;
            set;
        }

        internal string PhysicalDeliveryOfficeName
        {
            get;
            set;
        }

        internal string Company
        {
            get;
            set;
        }

        internal MSFTee Manager
        {
            get;
            set;
        }

        internal string ManagerAlias
        {
            get;
            set;
        }

        internal static MSFTee FromXElement(XElement xElement)
        {
            return new MSFTee
            {
                Domain = xElement.Attribute("domain").Value,
                Alias = xElement.Attribute("alias").Value,
                FullLogon = xElement.Attribute("fullLogon").Value,
                FullName = xElement.Attribute("fullName").Value,
                Email = xElement.Attribute("email").Value,
                FirstName = xElement.Attribute("firstName").Value,
                LastName = xElement.Attribute("lastName").Value,
                Title = xElement.Attribute("title").Value,
                Department = xElement.Attribute("department").Value,
                DistinguishedName = xElement.Attribute("distinguishedName").Value,
                StreetAddress = xElement.Attribute("streetAddress").Value,
                PostalCode = xElement.Attribute("postalCode").Value,
                City = xElement.Attribute("city").Value,
                State = xElement.Attribute("state").Value,
                Country = xElement.Attribute("country").Value,
                TelephoneNumber = xElement.Attribute("telephoneNumber").Value,
                HomePhone = xElement.Attribute("homePhone").Value,
                MobilePhone = xElement.Attribute("mobilePhone").Value,
                Fax = xElement.Attribute("fax").Value,
                PhysicalDeliveryOfficeName = xElement.Attribute("physicalDeliveryOfficeName").Value,
                Company = xElement.Attribute("company").Value,
                ManagerAlias = xElement.Attribute("manager").Value
            };
        }

        internal XElement ToXml()
        {
            return new XElement("Microsoftee", new object[]
            {
                new XAttribute("domain", this.Domain),
                new XAttribute("alias", this.Alias),
                new XAttribute("fullLogon", this.FullLogon),
                new XAttribute("fullName", this.FullName),
                new XAttribute("email", this.Email),
                new XAttribute("firstName", this.FirstName),
                new XAttribute("lastName", this.LastName),
                new XAttribute("title", this.Title),
                new XAttribute("department", this.Department),
                new XAttribute("distinguishedName", this.DistinguishedName),
                new XAttribute("streetAddress", this.StreetAddress),
                new XAttribute("postalCode", this.PostalCode),
                new XAttribute("city", this.City),
                new XAttribute("state", this.State),
                new XAttribute("country", this.Country),
                new XAttribute("telephoneNumber", this.TelephoneNumber),
                new XAttribute("homePhone", this.HomePhone),
                new XAttribute("mobilePhone", this.MobilePhone),
                new XAttribute("fax", this.Fax),
                new XAttribute("physicalDeliveryOfficeName", this.PhysicalDeliveryOfficeName),
                new XAttribute("company", this.Company),
                new XAttribute("manager", (this.Manager == null) ? string.Empty : this.Manager.Alias)
            });
        }

        public override string ToString()
        {
            return this.FullName;
        }
    }
}
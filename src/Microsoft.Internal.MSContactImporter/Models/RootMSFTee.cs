using System.Globalization;
using System.Xml.Linq;
using System.Xml.Serialization;

namespace Microsoft.Internal.MSContactImporter
{
    [XmlType("RootMSFTee")]
    public class RootMSFTee
    {
        [XmlAttribute("logon")]
        public string Logon
        {
            get;
            set;
        }

        [XmlAttribute("recurseLevel")]
        public int RecurseLevel
        {
            get;
            set;
        }

        public XElement ToXml()
        {
            return new XElement("RootMSFTee", new object[]
            {
                new XAttribute("logon", this.Logon),
                new XAttribute("recurseLevel", this.RecurseLevel.ToString(new CultureInfo("en-us")))
            });
        }

        public static RootMSFTee FromXml(XElement xElement)
        {
            return new RootMSFTee
            {
                Logon = xElement.Attribute("logon").Value,
                RecurseLevel = int.Parse(xElement.Attribute("recurseLevel").Value, new CultureInfo("en-us"))
            };
        }
    }
}
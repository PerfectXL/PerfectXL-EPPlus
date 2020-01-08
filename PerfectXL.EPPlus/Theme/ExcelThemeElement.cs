using System.Xml;

namespace OfficeOpenXml.Theme
{
    public class ExcelThemeElement : XmlHelper
    {
        protected ExcelThemeElement(XmlNamespaceManager nameSpaceManager, XmlElement themeElement) : base(nameSpaceManager, themeElement)
        {
            Name = themeElement.GetAttribute("name");
        }

        internal string ElementName => TopNode.LocalName;

        public string Name { get; }

        public static ExcelThemeElement Create(XmlNamespaceManager nameSpaceManager, XmlElement themeElement)
        {
            switch (themeElement.LocalName)
            {
                case "clrScheme": return new ExcelThemeColorScheme(nameSpaceManager, themeElement);
                case "fontScheme": // To be implemented
                case "fmtScheme": // To be implemented
                default: return new ExcelThemeElement(nameSpaceManager, themeElement);
            }
        }
    }
}
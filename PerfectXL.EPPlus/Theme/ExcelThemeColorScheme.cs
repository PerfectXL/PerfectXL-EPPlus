using System.Collections.Generic;
using System.Linq;
using System.Xml;

namespace OfficeOpenXml.Theme
{
    public class ExcelThemeColorScheme : ExcelThemeElement
    {
        private IList<SchemeColor> _schemeColors;
        public ExcelThemeColorScheme(XmlNamespaceManager nameSpaceManager, XmlElement themeElement) : base(nameSpaceManager, themeElement) { }

        public IList<SchemeColor> SchemeColors => _schemeColors ?? (_schemeColors = GetSchemeColors());

        private IList<SchemeColor> GetSchemeColors()
        {
            XmlNodeList xmlNodeList = TopNode.SelectNodes("*", NameSpaceManager);
            return xmlNodeList == null ? new List<SchemeColor>() : xmlNodeList.OfType<XmlElement>().Select(x => new SchemeColor(x)).ToList();
        }
    }
}
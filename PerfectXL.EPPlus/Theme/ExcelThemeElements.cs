using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Xml;
using OfficeOpenXml.Packaging;

namespace OfficeOpenXml.Theme
{
    public class ExcelThemeElements : XmlHelper, IEnumerable<ExcelThemeElement>
    {
        private readonly List<ExcelThemeElement> _themeElements = new List<ExcelThemeElement>();

        private ExcelThemeElements(ExcelPackage package, XmlNamespaceManager nameSpaceManager) : base(nameSpaceManager, null)
        {
            TopNode = GetThemesXmlTopNode(package);

            XmlNodeList themeNodes = TopNode?.SelectNodes("//a:themeElements/*", NameSpaceManager);

            if (themeNodes == null)
            {
                return;
            }

            foreach (XmlElement themeElement in themeNodes)
            {
                _themeElements.Add(ExcelThemeElement.Create(NameSpaceManager, themeElement));
            }
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return _themeElements.GetEnumerator();
        }

        public IEnumerator<ExcelThemeElement> GetEnumerator()
        {
            return _themeElements.GetEnumerator();
        }

        public static ExcelThemeElements Create(ExcelPackage package)
        {
            var namespaceManager = new XmlNamespaceManager(new NameTable());
            namespaceManager.AddNamespace("a", ExcelPackage.schemaDrawings);
            return new ExcelThemeElements(package, namespaceManager);
        }

        private static XmlNode GetThemesXmlTopNode(ExcelPackage package)
        {
            ZipPackageRelationship zipPackageRelationship =
                package.Workbook.Part.GetRelationshipsByType(ExcelPackage.schemaRelationships + "/theme").FirstOrDefault();
            if (zipPackageRelationship == null)
            {
                return null;
            }

            var partUri = new Uri($"/xl/{zipPackageRelationship.TargetUri}", UriKind.Relative);

            return !package.Package.PartExists(partUri) ? null : package.GetXmlFromUri(partUri).DocumentElement;
        }
    }
}
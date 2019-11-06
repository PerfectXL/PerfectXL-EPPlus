using System;
using System.Collections.Generic;
using System.Xml;
using OfficeOpenXml.Packaging;

namespace OfficeOpenXml
{
    public class ExcelExternalFilePaths : XmlHelper
    {
        public Dictionary<int, string> ExternalFilePaths { get; }
        public ExcelExternalFilePaths(ExcelPackage package, XmlNamespaceManager namespaceManager) : base(namespaceManager)
        {
            ExternalFilePaths = new Dictionary<int, string>();
            Uri externalLinkUri = new Uri($"xl/externalLinks/externalLink1.xml", UriKind.Relative);
            var i = 1;

            while (package.Package.PartExists(externalLinkUri))
            {
                XmlDocument externalLinkXml = new XmlDocument();
                ZipPackagePart zipPackagePart = package.Package.GetPart(externalLinkUri);
                LoadXmlSafe(externalLinkXml, zipPackagePart.GetStream());

                TopNode = externalLinkXml.DocumentElement;
                string rId = GetXmlNodeString("/d:externalLink/d:externalBook/@r:id");
                if (!string.IsNullOrEmpty(rId))
                {
                    var relation = zipPackagePart.GetRelationship(rId);
                    Uri targetUri = relation.TargetUri;
                    string absolutePath = targetUri.IsAbsoluteUri 
                        ? targetUri.LocalPath 
                        : $"{package.File.DirectoryName}\\{Uri.UnescapeDataString(relation.TargetUri.OriginalString)}";

                    ExternalFilePaths.Add(i, absolutePath);
                }

                i++;
                externalLinkUri = new Uri($"xl/externalLinks/externalLink{i}.xml", UriKind.Relative);
            }
        }
    }
}
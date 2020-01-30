using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml;
using OfficeOpenXml.Packaging;

namespace OfficeOpenXml
{
    public class ExcelExternalFiles : XmlHelper
    {
        public Dictionary<int, Uri> ExternalFileUri { get; }
        private string BaseDirectory { get; }
        public ExcelExternalFiles(ExcelPackage package, XmlNamespaceManager namespaceManager) : base(namespaceManager)
        {
            BaseDirectory = package.File?.DirectoryName;
            ExternalFileUri = new Dictionary<int, Uri>();
            Uri externalLinkUri = new Uri($"xl/externalLinks/externalLink1.xml", UriKind.Relative);
            var i = 1;

            while (package.Package.PartExists(externalLinkUri))
            {
                XmlDocument externalLinkXml = new XmlDocument();
                ZipPackagePart zipPackagePart = package.Package.GetPart(externalLinkUri);
                LoadXmlSafe(externalLinkXml, zipPackagePart.GetStream());

                TopNode = externalLinkXml.DocumentElement;
                string rId = GetXmlNodeString("/d:externalLink/d:externalBook/@r:id");
                if (zipPackagePart.TryGetRelationshipById(rId, out var relation))
                {
                    ExternalFileUri.Add(i, relation.TargetUri);
                }

                i++;
                externalLinkUri = new Uri($"xl/externalLinks/externalLink{i}.xml", UriKind.Relative);
            }
        }

        public IEnumerable<(int, string)> GetAbsolutePaths()
        {
            foreach (KeyValuePair<int, Uri> pair in ExternalFileUri)
            {
                string absolutePath;
                if (pair.Value.IsAbsoluteUri)
                {
                    absolutePath = pair.Value.LocalPath;
                }
                else
                {
                    absolutePath = string.IsNullOrEmpty(BaseDirectory) ? Uri.UnescapeDataString(pair.Value.OriginalString) : $"{BaseDirectory}\\{Uri.UnescapeDataString(pair.Value.OriginalString)}";
                }

                yield return (pair.Key, absolutePath);
            }
        }

        public IEnumerable<(int, string)> GetFileNames()
        {
            foreach ((var fileNumber, var absolutePath) in GetAbsolutePaths())
            {
                yield return (fileNumber, absolutePath.Split('\\').Last());
            }
        }
    }
}
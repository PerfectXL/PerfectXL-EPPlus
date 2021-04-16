using System;
using System.IO;
using System.Linq;
using System.Xml;
using OfficeOpenXml.Packaging;

namespace OfficeOpenXml
{
    public class ExcelDataMashup : XmlHelper
    {
        internal ExcelDataMashup(ExcelPackage package, XmlNamespaceManager nameSpaceManager) : base(nameSpaceManager)
        {
            try
            {
                var item1Uri = new Uri("/customXml/item1.xml", UriKind.Relative);
                if (!package.Package.PartExists(item1Uri))
                {
                    return;
                }

                var item1Xml = new XmlDocument();
                LoadXmlSafe(item1Xml, package.Package.GetPart(item1Uri).GetStream());
                TopNode = item1Xml.DocumentElement;
                byte[] dataMashup = Convert.FromBase64String(TopNode.InnerText);
                GetSection1M(dataMashup);
            }
            catch (Exception)
            {
                //customXml/item1.xml most likely not always intended as data mashup
                PowerQueryFormulas = null;
            }
        }

        private void GetSection1M(byte[] dataMashupBytes)
        {
            //Only reading section1.M from packaging parts
            int packagingPartsLength = BitConverter.ToUInt16(dataMashupBytes.Skip(4).Take(4).ToArray(), 0);
            byte[] packagingPartsBytes = dataMashupBytes.Skip(8).Take(packagingPartsLength).ToArray();

            using (MemoryStream packagingPartsStream = new MemoryStream(packagingPartsBytes))
            {
                var packagingParts = new ZipPackage(packagingPartsStream);
                ZipPackagePart section1M = packagingParts.GetPart(new Uri("/Formulas/Section1.m", UriKind.Relative));
                if (section1M == null)
                {
                    return;
                }

                using (var reader = new StreamReader(section1M.GetStream()))
                {
                    PowerQueryFormulas = reader.ReadToEnd();
                }
            }
        }
        public string PowerQueryFormulas { get; private set; }
    }
}

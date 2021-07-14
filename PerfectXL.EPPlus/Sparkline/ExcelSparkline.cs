﻿using System.Xml;

namespace OfficeOpenXml.Sparkline
{
    /// <summary>
    /// Represents a single sparkline within the sparkline group
    /// </summary>
    public class ExcelSparkline : XmlHelper
    {
        internal ExcelSparkline(XmlNamespaceManager nsm, XmlNode topNode) : base(nsm, topNode)
        {
            SchemaNodeOrder = new string[] { "f", "sqref" };
        }

        private const string _fPath = "xm:f";
        /// <summary>
        /// The datarange
        /// </summary>
        public ExcelAddressBase RangeAddress
        {
            get
            {
                return new ExcelAddressBase(GetXmlNodeString(_fPath));
            }
            internal set
            {
                SetXmlNodeString(_fPath, value.FullAddress);
            }
        }

        private const string _sqrefPath = "xm:sqref";
        /// <summary>
        /// Location of the sparkline
        /// </summary>
        public ExcelCellAddress Cell
        {
            get
            {
                return new ExcelCellAddress(GetXmlNodeString(_sqrefPath));
            }
            internal set
            {
                SetXmlNodeString("xm:sqref", value.Address);
            }
        }
        public override string ToString()
        {
            return Cell.Address + ", " + RangeAddress.Address;
        }
    }
}

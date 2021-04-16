using System;
using System.Xml;


/// <summary>
/// Line end style.
/// </summary>
public enum eEndStyle   //ST_LineEndType
{
    /// <summary>
    /// No end
    /// </summary>
    None,
    /// <summary>
    /// Triangle arrow head
    /// </summary>
    Triangle,
    /// <summary>
    /// Stealth arrow head
    /// </summary>
    Stealth,
    /// <summary>
    /// Diamond
    /// </summary>
    Diamond,
    /// <summary>
    /// Oval
    /// </summary>
    Oval,
    /// <summary>
    /// Line arrow head
    /// </summary>
    Arrow
}

/// <summary>
/// Lend end size.
/// </summary>
public enum eEndSize
{
    /// <summary>
    /// Smal
    /// </summary>
    Small,
    /// <summary>
    /// Medium
    /// </summary>
    Medium,
    /// <summary>
    /// Large
    /// </summary>
    Large
}

namespace OfficeOpenXml.Drawing
{
    /// <summary>
    /// Properties for drawing line ends
    /// </summary>
    public sealed class ExcelDrawingLineEnd : XmlHelper
    {

/* Unmerged change from project 'PerfectXL.EPPlus (net462)'
Before:
         string _linePath;
After:
        private string _linePath;
*/
        private readonly string _linePath;
        internal ExcelDrawingLineEnd(XmlNamespaceManager nameSpaceManager, XmlNode topNode, string linePath) :
            base(nameSpaceManager, topNode)
        {
            SchemaNodeOrder = new string[] { "headEnd", "tailEnd" };
            _linePath = linePath;
        }


/* Unmerged change from project 'PerfectXL.EPPlus (net462)'
Before:
        string _headEndStylePath = "xdr:sp/xdr:spPr/a:ln/a:headEnd/@type";
After:
        private string _headEndStylePath = "xdr:sp/xdr:spPr/a:ln/a:headEnd/@type";
*/
        private readonly string _headEndStylePath = "xdr:sp/xdr:spPr/a:ln/a:headEnd/@type";
        /// <summary>
        /// HeaderEnd
        /// </summary>
        public eEndStyle HeadEnd
        {
            get
            {
                return TranslateEndStyle(GetXmlNodeString(_headEndStylePath));
            }
            set
            {
                CreateNode(_linePath, false);
                SetXmlNodeString(_headEndStylePath, TranslateEndStyleText(value));
            }
        }


/* Unmerged change from project 'PerfectXL.EPPlus (net462)'
Before:
        string _tailEndStylePath = "xdr:sp/xdr:spPr/a:ln/a:tailEnd/@type";
After:
        private string _tailEndStylePath = "xdr:sp/xdr:spPr/a:ln/a:tailEnd/@type";
*/
        private readonly string _tailEndStylePath = "xdr:sp/xdr:spPr/a:ln/a:tailEnd/@type";
        /// <summary>
        /// HeaderEnd
        /// </summary>
        public eEndStyle TailEnd
        {
            get
            {
                return TranslateEndStyle(GetXmlNodeString(_tailEndStylePath));
            }
            set
            {
                CreateNode(_linePath, false);
                SetXmlNodeString(_tailEndStylePath, TranslateEndStyleText(value));
            }
        }


/* Unmerged change from project 'PerfectXL.EPPlus (net462)'
Before:
        string _tailEndSizeWidthPath = "xdr:sp/xdr:spPr/a:ln/a:tailEnd/@w";
After:
        private string _tailEndSizeWidthPath = "xdr:sp/xdr:spPr/a:ln/a:tailEnd/@w";
*/
        private readonly string _tailEndSizeWidthPath = "xdr:sp/xdr:spPr/a:ln/a:tailEnd/@w";
        /// <summary>
        /// TailEndSizeWidth
        /// </summary>
        public eEndSize TailEndSizeWidth
        {
            get
            {
                return TranslateEndSize(GetXmlNodeString(_tailEndSizeWidthPath));
            }
            set
            {
                CreateNode(_linePath, false);
                SetXmlNodeString(_tailEndSizeWidthPath, TranslateEndSizeText(value));
            }
        }


/* Unmerged change from project 'PerfectXL.EPPlus (net462)'
Before:
        string _tailEndSizeHeightPath = "xdr:sp/xdr:spPr/a:ln/a:tailEnd/@len";
After:
        private string _tailEndSizeHeightPath = "xdr:sp/xdr:spPr/a:ln/a:tailEnd/@len";
*/
        private readonly string _tailEndSizeHeightPath = "xdr:sp/xdr:spPr/a:ln/a:tailEnd/@len";
        /// <summary>
        /// TailEndSizeHeight
        /// </summary>
        public eEndSize TailEndSizeHeight
        {
            get
            {
                return TranslateEndSize(GetXmlNodeString(_tailEndSizeHeightPath));
            }
            set
            {
                CreateNode(_linePath, false);
                SetXmlNodeString(_tailEndSizeHeightPath, TranslateEndSizeText(value));
            }
        }


/* Unmerged change from project 'PerfectXL.EPPlus (net462)'
Before:
        string _headEndSizeWidthPath = "xdr:sp/xdr:spPr/a:ln/a:headEnd/@w";
After:
        private string _headEndSizeWidthPath = "xdr:sp/xdr:spPr/a:ln/a:headEnd/@w";
*/
        private readonly string _headEndSizeWidthPath = "xdr:sp/xdr:spPr/a:ln/a:headEnd/@w";
        /// <summary>
        /// TailEndSizeWidth
        /// </summary>
        public eEndSize HeadEndSizeWidth
        {
            get
            {
                return TranslateEndSize(GetXmlNodeString(_headEndSizeWidthPath));
            }
            set
            {
                CreateNode(_linePath, false);
                SetXmlNodeString(_headEndSizeWidthPath, TranslateEndSizeText(value));
            }
        }


/* Unmerged change from project 'PerfectXL.EPPlus (net462)'
Before:
        string _headEndSizeHeightPath = "xdr:sp/xdr:spPr/a:ln/a:headEnd/@len";
After:
        private string _headEndSizeHeightPath = "xdr:sp/xdr:spPr/a:ln/a:headEnd/@len";
*/
        private readonly string _headEndSizeHeightPath = "xdr:sp/xdr:spPr/a:ln/a:headEnd/@len";
        /// <summary>
        /// TailEndSizeHeight
        /// </summary>
        public eEndSize HeadEndSizeHeight
        {
            get
            {
                return TranslateEndSize(GetXmlNodeString(_headEndSizeHeightPath));
            }
            set
            {
                CreateNode(_linePath, false);
                SetXmlNodeString(_headEndSizeHeightPath, TranslateEndSizeText(value));
            }
        }

        #region "Translate Enum functions"
        private string TranslateEndStyleText(eEndStyle value)
        {
            return value.ToString().ToLower();
        }
        private eEndStyle TranslateEndStyle(string text)
        {
            switch (text)
            {
                case "none":
                case "arrow":
                case "diamond":
                case "oval":
                case "stealth":
                case "triangle":
                    return (eEndStyle)Enum.Parse(typeof(eEndStyle), text, true);
                default:
                    throw (new Exception("Invalid Endstyle"));
            }
        }

        private string TranslateEndSizeText(eEndSize value)
        {
            string text = value.ToString();
            switch (value)
            {
                case eEndSize.Small:
                    return "sm";
                case eEndSize.Medium:
                    return "med";
                case eEndSize.Large:
                    return "lg";
                default:
                    throw (new Exception("Invalid Endsize"));
            }
        }
        private eEndSize TranslateEndSize(string text)
        {
            switch (text)
            {
                case "sm":
                case "med":
                case "lg":
                    return (eEndSize)Enum.Parse(typeof(eEndSize), text, true);
                default:
                    throw (new Exception("Invalid Endsize"));
            }
        }
        #endregion
    }
}
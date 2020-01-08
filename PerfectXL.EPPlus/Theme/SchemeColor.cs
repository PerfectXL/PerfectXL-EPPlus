using System;
using System.Xml;

namespace OfficeOpenXml.Theme
{
    public class SchemeColor
    {
        public SchemeColor(XmlNode node)
        {
            ThemeColorName = InitializeThemeColorName(node);

            if (!(node.SelectSingleNode("*") is XmlElement colorNode))
            {
                return;
            }

            ThemeColorType = InitializeThemeColorType(colorNode);
            Value = ThemeColorType == ThemeColorType.SystemColor ? colorNode.GetAttribute("lastClr") : colorNode.GetAttribute("val");
        }

        public ThemeColorName ThemeColorName { get; }

        public ThemeColorType ThemeColorType { get; }

        public string Value { get; }

        private static ThemeColorName InitializeThemeColorName(XmlNode node)
        {
            switch (node.LocalName)
            {
                case "dk1": return ThemeColorName.Dark1;
                case "lt1": return ThemeColorName.Light1;
                case "dk2": return ThemeColorName.Dark2;
                case "lt2": return ThemeColorName.Light2;
                case "accent1": return ThemeColorName.Accent1;
                case "accent2": return ThemeColorName.Accent2;
                case "accent3": return ThemeColorName.Accent3;
                case "accent4": return ThemeColorName.Accent4;
                case "accent5": return ThemeColorName.Accent5;
                case "accent6": return ThemeColorName.Accent6;
                case "hlink": return ThemeColorName.Hyperlink;
                case "folHlink": return ThemeColorName.FollowedHyperlink;
                default: throw new ArgumentOutOfRangeException(node.LocalName);
            }
        }

        private static ThemeColorType InitializeThemeColorType(XmlNode node)
        {
            switch (node.LocalName)
            {
                case "sysClr": return ThemeColorType.SystemColor;
                case "srgbClr": return ThemeColorType.SrgbColor;
                default: throw new ArgumentOutOfRangeException(node?.LocalName);
            }
        }
    }
}
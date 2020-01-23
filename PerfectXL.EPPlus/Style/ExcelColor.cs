/*******************************************************************************
 * You may amend and distribute as you like, but don't remove this header!
 *
 * EPPlus provides server-side generation of Excel 2007/2010 spreadsheets.
 * See https://github.com/JanKallman/EPPlus for details.
 *
 * Copyright (C) 2011  Jan Källman
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.

 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  
 * See the GNU Lesser General Public License for more details.
 *
 * The GNU Lesser General Public License can be viewed at http://www.opensource.org/licenses/lgpl-license.php
 * If you unfamiliar with this license or have questions about it, here is an http://www.gnu.org/licenses/gpl-faq.html
 *
 * All code and executables are provided "as is" with no warranty either express or implied. 
 * The author accepts no liability for any damage or loss of business that this product may cause.
 *
 * Code change notes:
 * 
 * Author							Change						Date
 * ******************************************************************************
 * Jan Källman		                Initial Release		        2009-10-01
 * Jan Källman		License changed GPL-->LGPL 2011-12-16
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Text;
using OfficeOpenXml.Style.XmlAccess;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml;
using OfficeOpenXml.Theme;

namespace OfficeOpenXml.Style
{
    /// <summary>
    /// Color for cellstyling
    /// </summary>
    public sealed class ExcelColor :  StyleBase, IColor
    {
        eStyleClass _cls;
        StyleBase _parent;
        internal ExcelColor(ExcelStyles styles, OfficeOpenXml.XmlHelper.ChangedEventHandler ChangedEvent, int worksheetID, string address, eStyleClass cls, StyleBase parent) : 
            base(styles, ChangedEvent, worksheetID, address)
        {
            _parent = parent;
            _cls = cls;
        }
        /// <summary>
        /// The theme color
        /// </summary>
        public string Theme
        {
            get
            {
                return GetSource().Theme;
            }
        }
        /// <summary>
        /// The tint value
        /// </summary>
        public decimal Tint
        {
            get
            {
                return GetSource().Tint;
            }
            set
            {
                if (value > 1 || value < -1)
                {
                    throw (new ArgumentOutOfRangeException("Value must be between -1 and 1"));
                }
                _ChangedEvent(this, new StyleChangeEventArgs(_cls, eStyleProperty.Tint, value, _positionID, _address));
            }
        }
        /// <summary>
        /// The RGB value
        /// </summary>
        public string Rgb
        {
            get
            {
                return GetSource().Rgb;
            }
            internal set
            {
                _ChangedEvent(this, new StyleChangeEventArgs(_cls, eStyleProperty.Color, value, _positionID, _address));
            }
        }
        /// <summary>
        /// The indexed color number.
        /// </summary>
        public int? Indexed
        {
            get
            {
                return GetSource().Indexed;
            }
            set
            {
                _ChangedEvent(this, new StyleChangeEventArgs(_cls, eStyleProperty.IndexedColor, value, _positionID, _address));
            }
        }
        /// <summary>
        /// Set the color of the object
        /// </summary>
        /// <param name="color">The color</param>
        public void SetColor(Color color)
        {
            Rgb = color.ToArgb().ToString("X");       
        }
        /// <summary>
        /// Set the color of the object
        /// </summary>
        /// <param name="alpha">Alpha component value</param>
        /// <param name="red">Red component value</param>
        /// <param name="green">Green component value</param>
        /// <param name="blue">Blue component value</param>
        public void SetColor(int alpha, int red, int green, int blue)
        {
            if(alpha < 0 || red < 0 || green < 0 ||blue < 0 ||
               alpha > 255 || red > 255 || green > 255 || blue > 255)
            {
                throw (new ArgumentException("Argument range must be from 0 to 255"));
            }
            Rgb = alpha.ToString("X2") + red.ToString("X2") + green.ToString("X2") + blue.ToString("X2");
        }
        internal override string Id
        {
            get 
            {
                return Theme + Tint + Rgb + Indexed;
            }
        }
        private ExcelColorXml GetSource()
        {
            Index = _parent.Index < 0 ? 0 : _parent.Index;
            switch(_cls)
            {
                case eStyleClass.FillBackgroundColor:
                    return _styles.Fills[Index].BackgroundColor;
                case eStyleClass.FillPatternColor:
                    return _styles.Fills[Index].PatternColor;
                case eStyleClass.Font:
                    return _styles.Fonts[Index].Color;
                case eStyleClass.BorderLeft:
                    return _styles.Borders[Index].Left.Color;
                case eStyleClass.BorderTop:
                    return _styles.Borders[Index].Top.Color;
                case eStyleClass.BorderRight:
                    return _styles.Borders[Index].Right.Color;
                case eStyleClass.BorderBottom:
                    return _styles.Borders[Index].Bottom.Color;
                case eStyleClass.BorderDiagonal:
                    return _styles.Borders[Index].Diagonal.Color;
                default:
                    throw(new Exception("Invalid style-class for Color"));
            }
        }

        /// <summary>
        /// Return the RGB value for the Indexed or Tint property
        /// </summary>
        /// <param name="schemeColors">The list of colors for the current color scheme</param>
        /// <returns>The RGB color starting with a #</returns>
        public string LookupColor(ICollection<SchemeColor> schemeColors = null)
        {
            return LookupColor(this, schemeColors);
        }

        #region RgbLookup
        // reference extracted from ECMA-376, Part 4, Section 3.8.26 or 18.8.27 SE Part 1
        private static readonly string[] RgbLookup =
        {
            "#FF000000", // 0
            "#FFFFFFFF",
            "#FFFF0000",
            "#FF00FF00",
            "#FF0000FF",
            "#FFFFFF00",
            "#FFFF00FF",
            "#FF00FFFF",
            "#FF000000", // 8
            "#FFFFFFFF",
            "#FFFF0000",
            "#FF00FF00",
            "#FF0000FF",
            "#FFFFFF00",
            "#FFFF00FF",
            "#FF00FFFF",
            "#FF800000",
            "#FF008000",
            "#FF000080",
            "#FF808000",
            "#FF800080",
            "#FF008080",
            "#FFC0C0C0",
            "#FF808080",
            "#FF9999FF",
            "#FF993366",
            "#FFFFFFCC",
            "#FFCCFFFF",
            "#FF660066",
            "#FFFF8080",
            "#FF0066CC",
            "#FFCCCCFF",
            "#FF000080",
            "#FFFF00FF",
            "#FFFFFF00",
            "#FF00FFFF",
            "#FF800080",
            "#FF800000",
            "#FF008080",
            "#FF0000FF",
            "#FF00CCFF",
            "#FFCCFFFF",
            "#FFCCFFCC",
            "#FFFFFF99",
            "#FF99CCFF",
            "#FFFF99CC",
            "#FFCC99FF",
            "#FFFFCC99",
            "#FF3366FF",
            "#FF33CCCC",
            "#FF99CC00",
            "#FFFFCC00",
            "#FFFF9900",
            "#FFFF6600",
            "#FF666699",
            "#FF969696",
            "#FF003366",
            "#FF339966",
            "#FF003300",
            "#FF333300",
            "#FF993300",
            "#FF993366",
            "#FF333399",
            "#FF333333" // 63
        };
        #endregion

        /// <summary>
        /// Return the ARGB value for the color object that uses the Indexed or Tint property
        /// </summary>
        /// <param name="theColor">The color object</param>
        /// <param name="schemeColors">The list of colors for the current color scheme</param>
        /// <returns>The ARGB color starting with a "#". Or, if the color is not set: null.</returns>
        public static string LookupColor(ExcelColor theColor, ICollection<SchemeColor> schemeColors = null)
        {
            string rawColorString;
            if (!string.IsNullOrEmpty(theColor.Rgb))
            {
                rawColorString = PrefixColorString(theColor.Rgb);
            }
            else if (!string.IsNullOrEmpty(theColor.Theme) && Regex.IsMatch(theColor.Theme, @"^\d+$"))
            {
                var index = int.Parse(theColor.Theme);
                rawColorString = Enum.IsDefined(typeof(ThemeColorName), index)
                    ? PrefixColorString(schemeColors?.FirstOrDefault(x => x.ThemeColorName == (ThemeColorName) index)?.Value)
                    : null;
            }
            else if (theColor.Indexed == null)
            {
                rawColorString = null;
            }
            else
            {
                switch (theColor.Indexed.Value)
                {
                    case 64:
                        // System Foreground, get from theme color scheme, otherwise assume black
                        rawColorString = PrefixColorString(schemeColors?.FirstOrDefault(x => x.ThemeColorName == ThemeColorName.Dark1)?.Value ?? "000000");
                        break;
                    case 65:
                        // System Background, get from theme color scheme, otherwise assume white
                        rawColorString = PrefixColorString(schemeColors?.FirstOrDefault(x => x.ThemeColorName == ThemeColorName.Light1)?.Value ?? "FFFFFF");
                        break;
                    default:
                        rawColorString = RgbLookup.ElementAtOrDefault(theColor.Indexed.Value);
                        break;
                }
            }

            return ApplyTint(rawColorString, theColor.Tint);
        }

        private static string ApplyTint(string argbColor, decimal tint)
        {
            if (string.IsNullOrEmpty(argbColor) || tint == 0)
            {
                return argbColor;
            }

            Color color = HexValueToColor(argbColor);
            var (hue, li, sat) = ColorConversion.RgbToHls(color.R, color.G, color.B);
            li += tint < 0 ? li * (double) tint : (1.0 - li) * (double) tint;
            var (r, g, b) = ColorConversion.HlsToRgb(hue, li, sat);
            return PrefixColorString($"{r:X2}{g:X2}{b:X2}");
        }

        private static Color HexValueToColor(string hexColor)
        {
            return Color.FromArgb(int.Parse(hexColor.TrimStart('#'), NumberStyles.AllowHexSpecifier));
        }

        // Standardize to 32-bits color value that starts with a hash.
        private static string PrefixColorString(string hexValue)
        {
            if (hexValue == null)
            {
                return null;
            }

            if (hexValue.StartsWith("#"))
            {
                return hexValue;
            }

            return hexValue.Length == 8 ? $"#{hexValue}" : $"#FF{hexValue}";
        }

        #region ColorConversion
        private static class ColorConversion
        {
            public static (double hue, double li, double sat) RgbToHls(byte r, byte g, byte b)
            {
                var doubleR = r / 255.0;
                var doubleG = g / 255.0;
                var doubleB = b / 255.0;

                var max = new[] {doubleR, doubleB, doubleG}.Max();
                var min = new[] {doubleR, doubleB, doubleG}.Min();

                var diff = max - min;
                var li = (max + min) / 2;
                if (Math.Abs(diff) < 0.00001)
                {
                    return (0, li, 0);
                }

                double hue;
                double sat;
                if (li <= 0.5)
                {
                    sat = diff / (max + min);
                }
                else
                {
                    sat = diff / (2 - max - min);
                }

                var rDist = (max - doubleR) / diff;
                var gDist = (max - doubleG) / diff;
                var bDist = (max - doubleB) / diff;

                if (Math.Abs(doubleR - max) < 0.00001)
                {
                    hue = bDist - gDist;
                }
                else if (Math.Abs(doubleG - max) < 0.00001)
                {
                    hue = 2 + rDist - bDist;
                }
                else
                {
                    hue = 4 + gDist - rDist;
                }

                hue *= 60;
                if (hue < 0)
                {
                    hue += 360;
                }

                return (hue, li, sat);
            }

            public static (byte r, byte g, byte b) HlsToRgb(double hue, double li, double sat)
            {
                var p2 = li <= 0.5 ? li * (1 + sat) : li + sat - li * sat;
                var p1 = 2 * li - p2;

                if (Math.Abs(sat) < 0.00001)
                {
                    var rgb = (byte) (li * 255.0);
                    return (rgb, rgb, rgb);
                }

                var doubleR = QqhToRgb(p1, p2, hue + 120);
                var doubleG = QqhToRgb(p1, p2, hue);
                var doubleB = QqhToRgb(p1, p2, hue - 120);
                var r = (byte) (doubleR * 255.0);
                var g = (byte) (doubleG * 255.0);
                var b = (byte) (doubleB * 255.0);
                return (r, g, b);
            }

            private static double QqhToRgb(double q1, double q2, double hue)
            {
                if (hue > 360)
                {
                    hue -= 360;
                }

                if (hue < 0)
                {
                    hue += 360;
                }

                if (hue < 60)
                {
                    return q1 + (q2 - q1) * hue / 60;
                }

                if (hue < 180)
                {
                    return q2;
                }

                if (hue < 240)
                {
                    return q1 + (q2 - q1) * (240 - hue) / 60;
                }

                return q1;
            }
        }
        #endregion
    }
}

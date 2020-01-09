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
using OfficeOpenXml.Style.XmlAccess;

namespace OfficeOpenXml.Style
{
    /// <summary>
    /// Toplevel class for cell styling
    /// </summary>
    public sealed class ExcelStyle : StyleBase
    {
        private const string xfIdPath = "@xfId";
        private readonly ExcelXfs _xfs;

        internal ExcelStyle(ExcelStyles styles, XmlHelper.ChangedEventHandler ChangedEvent, int positionID, string Address, int xfsId) : base(styles,
            ChangedEvent, positionID, Address)
        {
            Index = xfsId;
            if (positionID > -1)
            {
                _xfs = _styles.CellXfs[xfsId];
            }
            else
            {
                _xfs = _styles.CellStyleXfs[xfsId];
            }

            Styles = styles;
            PositionID = positionID;
            Numberformat = new ExcelNumberFormat(styles, ChangedEvent, PositionID, Address, _xfs.NumberFormatId);
            Font = new ExcelFont(styles, ChangedEvent, PositionID, Address, _xfs.FontId);
            Fill = new ExcelFill(styles, ChangedEvent, PositionID, Address, _xfs.FillId);
            Border = new Border(styles, ChangedEvent, PositionID, Address, _xfs.BorderId);
        }

        /// <summary>
        /// Border
        /// </summary>
        public Border Border { get; set; }

        /// <summary>
        /// Fill Styling
        /// </summary>
        public ExcelFill Fill { get; set; }

        /// <summary>
        /// Font styling
        /// </summary>
        public ExcelFont Font { get; set; }

        /// <summary>
        /// If true the formula is hidden when the sheet is protected.
        /// <seealso cref="ExcelWorksheet.Protection" />
        /// </summary>
        public bool Hidden
        {
            get => _xfs.Hidden;
            set => _ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.Style, eStyleProperty.Hidden, value, _positionID, _address));
        }

        /// <summary>
        /// The horizontal alignment in the cell
        /// </summary>
        public ExcelHorizontalAlignment HorizontalAlignment
        {
            get => _xfs.HorizontalAlignment;
            set => _ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.Style, eStyleProperty.HorizontalAlign, value, _positionID, _address));
        }

        internal override string Id => Numberformat.Id + "|" + Font.Id + "|" + Fill.Id + "|" + Border.Id + "|" + VerticalAlignment + "|" + HorizontalAlignment +
                                       "|" + WrapText + "|" + ReadingOrder + "|" + XfId + "|" + QuotePrefix;

        /// <summary>
        /// The margin between the border and the text
        /// </summary>
        public int Indent
        {
            get => _xfs.Indent;
            set
            {
                if (value < 0 || value > 250)
                {
                    throw new ArgumentOutOfRangeException("Indent must be between 0 and 250");
                }

                _ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.Style, eStyleProperty.Indent, value, _positionID, _address));
            }
        }

        /// <summary>
        /// If true the cell is locked for editing when the sheet is protected
        /// <seealso cref="ExcelWorksheet.Protection" />
        /// </summary>
        public bool Locked
        {
            get => _xfs.Locked;
            set => _ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.Style, eStyleProperty.Locked, value, _positionID, _address));
        }

        /// <summary>
        /// Number format
        /// </summary>
        public ExcelNumberFormat Numberformat { get; set; }

        internal int PositionID { get; set; }

        /// <summary>
        /// If true the cell has a quote prefix, which indicates the value of the cell is prefixed with a single quote.
        /// </summary>
        public bool QuotePrefix
        {
            get => _xfs.QuotePrefix;
            set => _ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.Style, eStyleProperty.QuotePrefix, value, _positionID, _address));
        }

        /// <summary>
        /// Reading order
        /// </summary>
        public ExcelReadingOrder ReadingOrder
        {
            get => _xfs.ReadingOrder;
            set => _ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.Style, eStyleProperty.ReadingOrder, value, _positionID, _address));
        }

        /// <summary>
        /// Shrink the text to fit
        /// </summary>
        public bool ShrinkToFit
        {
            get => _xfs.ShrinkToFit;
            set => _ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.Style, eStyleProperty.ShrinkToFit, value, _positionID, _address));
        }

        internal ExcelStyles Styles { get; set; }

        /// <summary>
        /// Text orientation in degrees. Values range from 0 to 180.
        /// </summary>
        public int TextRotation
        {
            get => _xfs.TextRotation;
            set
            {
                if (value < 0 || value > 180)
                {
                    throw new ArgumentOutOfRangeException("TextRotation out of range.");
                }

                _ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.Style, eStyleProperty.TextRotation, value, _positionID, _address));
            }
        }

        /// <summary>
        /// The vertical alignment in the cell
        /// </summary>
        public ExcelVerticalAlignment VerticalAlignment
        {
            get => _xfs.VerticalAlignment;
            set => _ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.Style, eStyleProperty.VerticalAlign, value, _positionID, _address));
        }

        /// <summary>
        /// Wrap the text
        /// </summary>
        public bool WrapText
        {
            get => _xfs.WrapText;
            set => _ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.Style, eStyleProperty.WrapText, value, _positionID, _address));
        }

        /// <summary>
        /// The index in the style collection
        /// </summary>
        public int XfId
        {
            get => _xfs.XfId;
            set => _ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.Style, eStyleProperty.XfId, value, _positionID, _address));
        }
    }
}
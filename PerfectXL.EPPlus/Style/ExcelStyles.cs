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
 * Jan Källman		    Initial Release		        2009-10-01
 * Jan Källman		    License changed GPL-->LGPL 2011-12-27
 *******************************************************************************/

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Xml;
using OfficeOpenXml.ConditionalFormatting;
using OfficeOpenXml.ConditionalFormatting.Contracts;
using OfficeOpenXml.Style.Dxf;
using OfficeOpenXml.Style.XmlAccess;

namespace OfficeOpenXml.Style
{
    /// <summary>
    /// Contains all shared cell styles for a workbook
    /// </summary>
    internal sealed class ExcelStyles : XmlHelper
    {
        private const string NumberFormatsPath = "d:styleSheet/d:numFmts";
        private const string FontsPath = "d:styleSheet/d:fonts";
        private const string FillsPath = "d:styleSheet/d:fills";
        private const string BordersPath = "d:styleSheet/d:borders";
        private const string CellStyleXfsPath = "d:styleSheet/d:cellStyleXfs";
        private const string CellXfsPath = "d:styleSheet/d:cellXfs";
        private const string CellStylesPath = "d:styleSheet/d:cellStyles";
        private const string DxfsPath = "d:styleSheet/d:dxfs";
        private readonly XmlNamespaceManager _namespaceManager;
        internal int NextDfxNumFmtId = 164;

        //internal Dictionary<int, ExcelXfs> Styles = new Dictionary<int, ExcelXfs>();
        private readonly XmlDocument _styleXml;
        private readonly ExcelWorkbook _wb;
        internal ExcelStyleCollection<ExcelBorderXml> Borders = new ExcelStyleCollection<ExcelBorderXml>();
        internal ExcelStyleCollection<ExcelXfs> CellStyleXfs = new ExcelStyleCollection<ExcelXfs>();
        internal ExcelStyleCollection<ExcelXfs> CellXfs = new ExcelStyleCollection<ExcelXfs>();
        internal ExcelStyleCollection<ExcelDxfStyleConditionalFormatting> Dxfs = new ExcelStyleCollection<ExcelDxfStyleConditionalFormatting>();
        internal ExcelStyleCollection<ExcelFillXml> Fills = new ExcelStyleCollection<ExcelFillXml>();
        internal ExcelStyleCollection<ExcelFontXml> Fonts = new ExcelStyleCollection<ExcelFontXml>();
        internal ExcelStyleCollection<ExcelNamedStyleXml> NamedStyles = new ExcelStyleCollection<ExcelNamedStyleXml>();
        internal ExcelStyleCollection<ExcelNumberFormatXml> NumberFormats = new ExcelStyleCollection<ExcelNumberFormatXml>();

        internal ExcelStyles(XmlNamespaceManager namespaceManager, XmlDocument xml, ExcelWorkbook wb) : base(namespaceManager, xml)
        {
            _styleXml = xml;
            _wb = wb;
            _namespaceManager = namespaceManager;
            SchemaNodeOrder = new[] {"numFmts", "fonts", "fills", "borders", "cellStyleXfs", "cellXfs", "cellStyles", "dxfs"};
            LoadFromDocument();
        }

        internal string Id => "";

        public ExcelNamedStyle CreateNamedStyle(string name)
        {
            return CreateNamedStyle(name, null);
        }

        public ExcelNamedStyle CreateNamedStyle(string name, ExcelStyle template)
        {
            if (_wb.Styles.NamedStyles.ExistsKey(name))
            {
                throw new Exception(string.Format("Key {0} already exists in collection", name));
            }

            var style = new ExcelNamedStyleXml(NameSpaceManager, this);
            int xfIdCopy, positionID;
            ExcelStyles styles;
            if (template == null)
            {
                // style.Style = new ExcelStyle(this, NamedStylePropertyChange, -1, name, 0);
                xfIdCopy = 0;
                positionID = -1;
                styles = this;
            }
            else
            {
                xfIdCopy = template.PositionID < 0 ? template.Index : template.XfId;
                positionID = -1;
                styles = template.Styles;
            }

            //Clone named style
            var styleXfId = CloneStyle(styles, xfIdCopy, true);
            //Clone cells style
            CellStyleXfs[styleXfId].XfId = CellStyleXfs.Count - 1;
            var xfId = CloneStyle(styles, xfIdCopy, true, true); //Always add a new style (We create a new named style here)
            CellXfs[xfId].XfId = styleXfId;
            style.Style = new ExcelStyle(this, NamedStylePropertyChange, positionID, name, styleXfId);
            style.StyleXfId = styleXfId;

            style.Name = name;
            var ix = _wb.Styles.NamedStyles.Add(style.Name, style);
            style.Style.Index = ix;
            return new ExcelNamedStyle(style);
        }

        public void UpdateXml()
        {
            RemoveUnusedStyles();

            //NumberFormat
            XmlNode nfNode = _styleXml.SelectSingleNode(NumberFormatsPath, _namespaceManager);
            if (nfNode == null)
            {
                CreateNode(NumberFormatsPath, true);
                nfNode = _styleXml.SelectSingleNode(NumberFormatsPath, _namespaceManager);
            }
            else
            {
                nfNode.RemoveAll();
            }

            Debug.Assert(nfNode != null, nameof(nfNode) + " != null");

            var count = 0;
            var normalIx = NamedStyles.FindIndexById("Normal");
            if (NamedStyles.Count > 0 && normalIx >= 0 && NamedStyles[normalIx].Style.Numberformat.NumFmtID >= 164)
            {
                ExcelNumberFormatXml nf = NumberFormats[NumberFormats.FindIndexById(NamedStyles[normalIx].Style.Numberformat.Id)];
                nfNode.AppendChild(nf.CreateXmlNode(_styleXml.CreateElement("numFmt", ExcelPackage.schemaMain)));
                nf.newID = count++;
            }

            foreach (ExcelNumberFormatXml nf in NumberFormats)
            {
                if (!nf.BuildIn /*&& nf.newID<0*/) //Buildin formats are not updated.
                {
                    nfNode.AppendChild(nf.CreateXmlNode(_styleXml.CreateElement("numFmt", ExcelPackage.schemaMain)));
                    nf.newID = count;
                    count++;
                }
            }

            ((XmlElement) nfNode).SetAttribute("count", count.ToString());

            //Font
            count = 0;
            XmlNode fntNode = _styleXml.SelectSingleNode(FontsPath, _namespaceManager);
            Debug.Assert(fntNode != null, nameof(fntNode) + " != null");
            fntNode.RemoveAll();

            //Normal should be first in the collection
            if (NamedStyles.Count > 0 && normalIx >= 0 && NamedStyles[normalIx].Style.Font.Index > 0)
            {
                ExcelFontXml fnt = Fonts[NamedStyles[normalIx].Style.Font.Index];
                fntNode.AppendChild(fnt.CreateXmlNode(_styleXml.CreateElement("font", ExcelPackage.schemaMain)));
                fnt.newID = count++;
            }

            foreach (ExcelFontXml fnt in Fonts)
            {
                if (fnt.useCnt > 0 /* && fnt.newID<0*/)
                {
                    fntNode.AppendChild(fnt.CreateXmlNode(_styleXml.CreateElement("font", ExcelPackage.schemaMain)));
                    fnt.newID = count;
                    count++;
                }
            }

            ((XmlElement) fntNode).SetAttribute("count", count.ToString());

            //Fills
            count = 0;
            XmlNode fillsNode = _styleXml.SelectSingleNode(FillsPath, _namespaceManager);
            Debug.Assert(fillsNode != null, nameof(fillsNode) + " != null");
            fillsNode.RemoveAll();
            Fills[0].useCnt = 1; //Must exist (none);  
            Fills[1].useCnt = 1; //Must exist (gray125);
            foreach (ExcelFillXml fill in Fills)
            {
                if (fill.useCnt > 0)
                {
                    fillsNode.AppendChild(fill.CreateXmlNode(_styleXml.CreateElement("fill", ExcelPackage.schemaMain)));
                    fill.newID = count;
                    count++;
                }
            }

            ((XmlElement) fillsNode).SetAttribute("count", count.ToString());

            //Borders
            count = 0;
            XmlNode bordersNode = _styleXml.SelectSingleNode(BordersPath, _namespaceManager);
            Debug.Assert(bordersNode != null, nameof(bordersNode) + " != null");
            bordersNode.RemoveAll();
            Borders[0].useCnt = 1; //Must exist blank;
            foreach (ExcelBorderXml border in Borders)
            {
                if (border.useCnt > 0)
                {
                    bordersNode.AppendChild(border.CreateXmlNode(_styleXml.CreateElement("border", ExcelPackage.schemaMain)));
                    border.newID = count;
                    count++;
                }
            }

            ((XmlElement) bordersNode).SetAttribute("count", count.ToString());

            XmlNode styleXfsNode = _styleXml.SelectSingleNode(CellStyleXfsPath, _namespaceManager);
            if (styleXfsNode == null && NamedStyles.Count > 0)
            {
                CreateNode(CellStyleXfsPath);
                styleXfsNode = _styleXml.SelectSingleNode(CellStyleXfsPath, _namespaceManager);
            }

            if (NamedStyles.Count > 0)
            {
                styleXfsNode?.RemoveAll();
            }

            //NamedStyles
            count = normalIx > -1 ? 1 : 0; //If we have a normal style, we make sure it's added first.

            XmlNode cellStyleNode = _styleXml.SelectSingleNode(CellStylesPath, _namespaceManager);
            cellStyleNode?.RemoveAll();
            XmlNode cellXfsNode = _styleXml.SelectSingleNode(CellXfsPath, _namespaceManager);
            cellXfsNode?.RemoveAll();

            if (NamedStyles.Count > 0 && normalIx >= 0)
            {
                NamedStyles[normalIx].newID = 0;
                AddNamedStyle(0, styleXfsNode, cellXfsNode, NamedStyles[normalIx]);
            }

            foreach (ExcelNamedStyleXml style in NamedStyles)
            {
                if (!style.Name.Equals("normal", StringComparison.OrdinalIgnoreCase))
                {
                    AddNamedStyle(count++, styleXfsNode, cellXfsNode, style);
                }
                else
                {
                    style.newID = 0;
                }

                cellStyleNode.AppendChild(style.CreateXmlNode(_styleXml.CreateElement("cellStyle", ExcelPackage.schemaMain)));
            }

            if (cellStyleNode != null)
            {
                (cellStyleNode as XmlElement).SetAttribute("count", count.ToString());
            }

            if (styleXfsNode != null)
            {
                (styleXfsNode as XmlElement).SetAttribute("count", count.ToString());
            }

            //CellStyle
            var xfix = 0;
            foreach (ExcelXfs xf in CellXfs)
            {
                if (xf.useCnt > 0 && !(normalIx >= 0 && NamedStyles[normalIx].StyleXfId == xfix))
                {
                    cellXfsNode.AppendChild(xf.CreateXmlNode(_styleXml.CreateElement("xf", ExcelPackage.schemaMain)));
                    xf.newID = count;
                    count++;
                }

                xfix++;
            }

            (cellXfsNode as XmlElement).SetAttribute("count", count.ToString());

            //Set dxf styling for conditional Formatting
            XmlNode dxfsNode = _styleXml.SelectSingleNode(DxfsPath, _namespaceManager);
            foreach (ExcelWorksheet ws in _wb.Worksheets)
            {
                if (ws is ExcelChartsheet)
                {
                    continue;
                }

                foreach (IExcelConditionalFormattingRuleStyle cf in ws.ConditionalFormatting.Cast<IExcelConditionalFormattingRuleStyle>())
                {
                    if (cf.Style.HasValue)
                    {
                        var ix = Dxfs.FindIndexById(cf.Style.Id);
                        if (ix < 0)
                        {
                            ((ExcelConditionalFormattingRule) cf).DxfId = Dxfs.Count;
                            Dxfs.Add(cf.Style.Id, cf.Style);
                            XmlElement elem = ((XmlDocument) TopNode).CreateElement("d", "dxf", ExcelPackage.schemaMain);
                            cf.Style.CreateNodes(new XmlHelperInstance(NameSpaceManager, elem), "");
                            dxfsNode.AppendChild(elem);
                        }
                        else
                        {
                            ((ExcelConditionalFormattingRule) cf).DxfId = ix;
                        }
                    }
                }
            }

            if (dxfsNode != null)
            {
                (dxfsNode as XmlElement).SetAttribute("count", Dxfs.Count.ToString());
            }
        }

        internal int CloneStyle(ExcelStyles style, int styleId)
        {
            return CloneStyle(style, styleId, false, false);
        }

        internal int CloneStyle(ExcelStyles style, int styleId, bool isNamedStyle)
        {
            return CloneStyle(style, styleId, isNamedStyle, false);
        }

        internal int CloneStyle(ExcelStyles style, int styleId, bool isNamedStyle, bool alwaysAddCellXfs)
        {
            ExcelXfs xfs;
            lock (style)
            {
                if (isNamedStyle)
                {
                    xfs = style.CellStyleXfs[styleId];
                }
                else
                {
                    xfs = style.CellXfs[styleId];
                }

                ExcelXfs newXfs = xfs.Copy(this);
                //Number format
                if (xfs.NumberFormatId > 0)
                {
                    //rake36: Two problems here...
                    //rake36:  1. the first time through when format stays equal to String.Empty, it adds a string.empty to the list of Number Formats
                    //rake36:  2. when adding a second sheet, if the numberformatid == 164, it finds the 164 added by previous sheets but was using the array index
                    //rake36:      for the numberformatid

                    var format = string.Empty;
                    foreach (ExcelNumberFormatXml fmt in style.NumberFormats)
                    {
                        if (fmt.NumFmtId == xfs.NumberFormatId)
                        {
                            format = fmt.Format;
                            break;
                        }
                    }

                    //rake36: Don't add another format if it's blank
                    if (!string.IsNullOrEmpty(format))
                    {
                        var ix = NumberFormats.FindIndexById(format);
                        if (ix < 0)
                        {
                            var item = new ExcelNumberFormatXml(NameSpaceManager) {Format = format, NumFmtId = NumberFormats.NextId++};
                            NumberFormats.Add(format, item);
                            //rake36: Use the just added format id
                            newXfs.NumberFormatId = item.NumFmtId;
                        }
                        else
                        {
                            //rake36: Use the format id defined by the index... not the index itself
                            newXfs.NumberFormatId = NumberFormats[ix].NumFmtId;
                        }
                    }
                }

                //Font
                if (xfs.FontId > -1)
                {
                    var ix = Fonts.FindIndexById(xfs.Font.Id);
                    if (ix < 0)
                    {
                        ExcelFontXml item = style.Fonts[xfs.FontId].Copy();
                        ix = Fonts.Add(xfs.Font.Id, item);
                    }

                    newXfs.FontId = ix;
                }

                //Border
                if (xfs.BorderId > -1)
                {
                    var ix = Borders.FindIndexById(xfs.Border.Id);
                    if (ix < 0)
                    {
                        ExcelBorderXml item = style.Borders[xfs.BorderId].Copy();
                        ix = Borders.Add(xfs.Border.Id, item);
                    }

                    newXfs.BorderId = ix;
                }

                //Fill
                if (xfs.FillId > -1)
                {
                    var ix = Fills.FindIndexById(xfs.Fill.Id);
                    if (ix < 0)
                    {
                        ExcelFillXml item = style.Fills[xfs.FillId].Copy();
                        ix = Fills.Add(xfs.Fill.Id, item);
                    }

                    newXfs.FillId = ix;
                }

                //Named style reference
                if (xfs.XfId > 0)
                {
                    var id = style.CellStyleXfs[xfs.XfId].Id;
                    var newId = CellStyleXfs.FindIndexById(id);
                    if (newId >= 0)
                    {
                        newXfs.XfId = newId;
                    }
                    else if (style._wb != _wb && alwaysAddCellXfs == false) //Not the same workbook, copy the namedstyle to the workbook or match the id
                    {
                        var nsFind = style.NamedStyles.ToDictionary(d => d.StyleXfId);
                        if (nsFind.ContainsKey(xfs.XfId))
                        {
                            ExcelNamedStyleXml st = nsFind[xfs.XfId];
                            if (NamedStyles.ExistsKey(st.Name))
                            {
                                newXfs.XfId = NamedStyles.FindIndexById(st.Name);
                            }
                            else
                            {
                                ExcelNamedStyle ns = CreateNamedStyle(st.Name, st.Style);
                                newXfs.XfId = NamedStyles.Count - 1;
                            }
                        }
                    }
                }

                int index;
                if (isNamedStyle && alwaysAddCellXfs == false)
                {
                    index = CellStyleXfs.Add(newXfs.Id, newXfs);
                }
                else
                {
                    if (alwaysAddCellXfs)
                    {
                        index = CellXfs.Add(newXfs.Id, newXfs);
                    }
                    else
                    {
                        index = CellXfs.FindIndexById(newXfs.Id);
                        if (index < 0)
                        {
                            index = CellXfs.Add(newXfs.Id, newXfs);
                        }
                    }
                }

                return index;
            }
        }

        internal int GetStyleId(ExcelWorksheet ws, int row, int col)
        {
            var v = 0;
            if (ws.ExistsStyleInner(row, col, ref v))
            {
                return v;
            }

            if (ws.ExistsStyleInner(row, 0, ref v)) //First Row
            {
                return v;
            }

            if (ws.ExistsStyleInner(0, col, ref v))
            {
                return v;
            }

            int r = 0, c = col;
            if (ws._values.PrevCell(ref r, ref c))
            {
                //var column=ws.GetValueInner(0,c) as ExcelColumn;
                ExcelCoreValue val = ws._values.GetValue(0, c);
                var column = (ExcelColumn) val._value;
                if (column != null && column.ColumnMax >= col) //Fixes issue 15174
                {
                    //return ws.GetStyleInner(0, c);
                    return val._styleId;
                }

                return 0;
            }

            return 0;
        }

        internal int GetStyleIdFromName(string name)
        {
            var i = NamedStyles.FindIndexById(name);
            if (i >= 0)
            {
                var id = NamedStyles[i].XfId;
                if (id < 0)
                {
                    var styleXfId = NamedStyles[i].StyleXfId;
                    ExcelXfs newStyle = CellStyleXfs[styleXfId].Copy();
                    newStyle.XfId = styleXfId;
                    id = CellXfs.FindIndexById(newStyle.Id);
                    if (id < 0)
                    {
                        id = CellXfs.Add(newStyle.Id, newStyle);
                    }

                    NamedStyles[i].XfId = id;
                }

                return id;
            }

            return 0;
            //throw(new Exception("Named style does not exist"));        	         
        }

        internal ExcelStyle GetStyleObject(int id, int positionId, string address)
        {
            if (id < 0)
            {
                id = 0;
            }

            return new ExcelStyle(this, PropertyChange, positionId, address, id);
        }

        /// <summary>
        ///     Handles property changes on Named styles.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// <returns></returns>
        internal int NamedStylePropertyChange(StyleBase sender, StyleChangeEventArgs e)
        {
            var index = NamedStyles.FindIndexById(e.Address);
            if (index >= 0)
            {
                var newId = CellStyleXfs[NamedStyles[index].StyleXfId].GetNewID(CellStyleXfs, sender, e.StyleClass, e.StyleProperty, e.Value);
                var prevIx = NamedStyles[index].StyleXfId;
                NamedStyles[index].StyleXfId = newId;
                NamedStyles[index].Style.Index = newId;

                NamedStyles[index].XfId = int.MinValue;
                foreach (ExcelXfs style in CellXfs)
                {
                    if (style.XfId == prevIx)
                    {
                        style.XfId = newId;
                    }
                }
            }

            return 0;
        }

        /// <summary>
        ///     Handels changes of properties on the style objects
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// <returns></returns>
        internal int PropertyChange(StyleBase sender, StyleChangeEventArgs e)
        {
            var address = new ExcelAddressBase(e.Address);
            ExcelWorksheet ws = _wb.Worksheets[e.PositionID];
            var styleCashe = new Dictionary<int, int>();
            //Set single address
            lock (ws._values)
            {
                SetStyleAddress(sender, e, address, ws, ref styleCashe);
                if (address.Addresses != null)
                {
                    //Handle multi-addresses
                    foreach (ExcelAddress innerAddress in address.Addresses)
                    {
                        SetStyleAddress(sender, e, innerAddress, ws, ref styleCashe);
                    }
                }
            }

            return 0;
        }

        private void AddNamedStyle(int id, XmlNode styleXfsNode, XmlNode cellXfsNode, ExcelNamedStyleXml style)
        {
            ExcelXfs styleXfs = CellStyleXfs[style.StyleXfId];
            styleXfsNode.AppendChild(styleXfs.CreateXmlNode(_styleXml.CreateElement("xf", ExcelPackage.schemaMain), true));
            styleXfs.newID = id;
            styleXfs.XfId = style.StyleXfId;

            var ix = CellXfs.FindIndexById(styleXfs.Id);
            if (ix < 0)
            {
                cellXfsNode.AppendChild(styleXfs.CreateXmlNode(_styleXml.CreateElement("xf", ExcelPackage.schemaMain)));
            }
            else
            {
                if (id < 0)
                {
                    CellXfs[ix].XfId = id;
                }

                cellXfsNode.AppendChild(CellXfs[ix].CreateXmlNode(_styleXml.CreateElement("xf", ExcelPackage.schemaMain)));
                CellXfs[ix].useCnt = 0;
                CellXfs[ix].newID = id;
            }

            if (style.XfId >= 0)
            {
                style.XfId = CellXfs[style.XfId].newID;
            }
            else
            {
                style.XfId = 0;
            }
        }

        private void AddNewStyleColumn(StyleBase sender, StyleChangeEventArgs e, ExcelWorksheet ws, Dictionary<int, int> styleCashe, ExcelColumn column, int s)
        {
            if (styleCashe.ContainsKey(s))
            {
                ws.SetStyleInner(0, column.ColumnMin, styleCashe[s]);
            }
            else
            {
                ExcelXfs st = CellXfs[s];
                var newId = st.GetNewID(CellXfs, sender, e.StyleClass, e.StyleProperty, e.Value);
                styleCashe.Add(s, newId);
                ws.SetStyleInner(0, column.ColumnMin, newId);
            }
        }

        /// <summary>
        ///     Loads the style XML to memory
        /// </summary>
        private void LoadFromDocument()
        {
            //NumberFormats
            ExcelNumberFormatXml.AddBuildIn(NameSpaceManager, NumberFormats);
            XmlNode numNode = _styleXml.SelectSingleNode(NumberFormatsPath, _namespaceManager);
            if (numNode != null)
            {
                foreach (XmlNode n in numNode)
                {
                    var nf = new ExcelNumberFormatXml(_namespaceManager, n);
                    NumberFormats.Add(nf.Id, nf);
                    if (nf.NumFmtId >= NumberFormats.NextId)
                    {
                        NumberFormats.NextId = nf.NumFmtId + 1;
                    }
                }
            }

            //Fonts
            XmlNode fontNode = _styleXml.SelectSingleNode(FontsPath, _namespaceManager);
            foreach (XmlNode n in fontNode)
            {
                var f = new ExcelFontXml(_namespaceManager, n);
                Fonts.Add(f.Id, f);
            }

            //Fills
            XmlNode fillNode = _styleXml.SelectSingleNode(FillsPath, _namespaceManager);
            foreach (XmlNode n in fillNode)
            {
                ExcelFillXml f;
                if (n.FirstChild != null && n.FirstChild.LocalName == "gradientFill")
                {
                    f = new ExcelGradientFillXml(_namespaceManager, n);
                }
                else
                {
                    f = new ExcelFillXml(_namespaceManager, n);
                }

                Fills.Add(f.Id, f);
            }

            //Borders
            XmlNode borderNode = _styleXml.SelectSingleNode(BordersPath, _namespaceManager);
            foreach (XmlNode n in borderNode)
            {
                var b = new ExcelBorderXml(_namespaceManager, n);
                Borders.Add(b.Id, b);
            }

            //cellStyleXfs
            XmlNodeList styleXfsNodes = _styleXml.SelectNodes($"{CellStyleXfsPath}/d:xf", _namespaceManager);
            if (styleXfsNodes != null && styleXfsNodes.Count > 0)
            {
                for (var i = 0; i < styleXfsNodes.Count; i++)
                {
                    XmlNode n = styleXfsNodes[i];
                    var item = new ExcelXfs(_namespaceManager, n, this, i);
                    CellStyleXfs.Add(item.Id, item);
                }
            }

            XmlNodeList styleNodes = _styleXml.SelectNodes($"{CellXfsPath}/d:xf", _namespaceManager);
            if (styleNodes != null && styleNodes.Count > 0)
            {
                for (var i = 0; i < styleNodes.Count; i++)
                {
                    XmlNode n = styleNodes[i];
                    var item = new ExcelXfs(_namespaceManager, n, this, i);
                    CellXfs.Add(item.Id, item);
                }
            }

            //cellStyle
            XmlNode namedStyleNode = _styleXml.SelectSingleNode(CellStylesPath, _namespaceManager);
            if (namedStyleNode != null)
            {
                foreach (XmlNode n in namedStyleNode)
                {
                    var item = new ExcelNamedStyleXml(_namespaceManager, n, this);
                    NamedStyles.Add(item.Name, item);
                }
            }

            //dxfsPath
            XmlNode dxfsNode = _styleXml.SelectSingleNode(DxfsPath, _namespaceManager);
            if (dxfsNode != null)
            {
                foreach (XmlNode x in dxfsNode)
                {
                    var item = new ExcelDxfStyleConditionalFormatting(_namespaceManager, x, this);
                    Dxfs.Add(item.Id, item);
                }
            }
        }

        private void RemoveUnusedStyles()
        {
            CellXfs[0].useCnt = 1; //First item is always used.
            foreach (ExcelWorksheet sheet in _wb.Worksheets)
            {
                var cse = new CellsStoreEnumerator<ExcelCoreValue>(sheet._values);
                while (cse.Next())
                {
                    var v = cse.Value._styleId;
                    if (v >= 0)
                    {
                        CellXfs[v].useCnt++;
                    }
                }
            }

            foreach (ExcelNamedStyleXml ns in NamedStyles)
            {
                CellStyleXfs[ns.StyleXfId].useCnt++;
            }

            foreach (ExcelXfs xf in CellXfs)
            {
                if (xf.useCnt > 0)
                {
                    if (xf.FontId >= 0)
                    {
                        Fonts[xf.FontId].useCnt++;
                    }

                    if (xf.FillId >= 0)
                    {
                        Fills[xf.FillId].useCnt++;
                    }

                    if (xf.BorderId >= 0)
                    {
                        Borders[xf.BorderId].useCnt++;
                    }
                }
            }

            foreach (ExcelXfs xf in CellStyleXfs)
            {
                if (xf.useCnt > 0)
                {
                    if (xf.FontId >= 0)
                    {
                        Fonts[xf.FontId].useCnt++;
                    }

                    if (xf.FillId >= 0)
                    {
                        Fills[xf.FillId].useCnt++;
                    }

                    if (xf.BorderId >= 0)
                    {
                        Borders[xf.BorderId].useCnt++;
                    }
                }
            }
        }

        private void SetStyleAddress(StyleBase sender, StyleChangeEventArgs e, ExcelAddressBase address, ExcelWorksheet ws, ref Dictionary<int, int> styleCashe)
        {
            if (address.Start.Column == 0 || address.Start.Row == 0)
            {
                throw new Exception("error address");
            }
            //Columns

            if (address.Start.Row == 1 && address.End.Row == ExcelPackage.MaxRows)
            {
                ExcelColumn column;
                int col = address.Start.Column, row = 0;
                bool isNew;
                //Get the startcolumn
                object o = null;
                if (!ws.ExistsValueInner(0, address.Start.Column, ref o))
                {
                    column = ws.Column(address.Start.Column);
                    isNew = true;
                }
                else
                {
                    //column = (ExcelColumn)ws.GetValueInner(0, address.Start.Column);
                    column = (ExcelColumn) o;
                    isNew = false;
                }

                var prevColumnMax = column.ColumnMax;
                while (column.ColumnMin <= address.End.Column)
                {
                    if (column.ColumnMin > prevColumnMax + 1)
                    {
                        ExcelColumn newColumn = ws.Column(prevColumnMax + 1);
                        newColumn.ColumnMax = column.ColumnMin - 1;
                        AddNewStyleColumn(sender, e, ws, styleCashe, newColumn, newColumn.StyleID);
                    }

                    if (column.ColumnMax > address.End.Column)
                    {
                        ExcelColumn newCol = ws.CopyColumn(column, address.End.Column + 1, column.ColumnMax);
                        column.ColumnMax = address.End.Column;
                    }

                    var s = ws.GetStyleInner(0, column.ColumnMin);
                    AddNewStyleColumn(sender, e, ws, styleCashe, column, s);

                    //index++;
                    prevColumnMax = column.ColumnMax;
                    if (!ws._values.NextCell(ref row, ref col) || row > 0)
                    {
                        if (column._columnMax == address.End.Column)
                        {
                            break;
                        }

                        if (isNew)
                        {
                            column._columnMax = address.End.Column;
                        }
                        else
                        {
                            ExcelColumn newColumn = ws.Column(column._columnMax + 1);
                            newColumn.ColumnMax = address.End.Column;
                            AddNewStyleColumn(sender, e, ws, styleCashe, newColumn, newColumn.StyleID);
                            column = newColumn;
                        }

                        break;
                    }

                    column = ws.GetValueInner(0, col) as ExcelColumn;
                }

                if (column._columnMax < address.End.Column)
                {
                    ExcelColumn newCol = ws.Column(column._columnMax + 1);
                    newCol._columnMax = address.End.Column;

                    var s = ws.GetStyleInner(0, column.ColumnMin);
                    if (styleCashe.ContainsKey(s))
                    {
                        ws.SetStyleInner(0, column.ColumnMin, styleCashe[s]);
                    }
                    else
                    {
                        ExcelXfs st = CellXfs[s];
                        var newId = st.GetNewID(CellXfs, sender, e.StyleClass, e.StyleProperty, e.Value);
                        styleCashe.Add(s, newId);
                        ws.SetStyleInner(0, column.ColumnMin, newId);
                    }

                    column._columnMax = address.End.Column;
                }

                //Set for individual cells in the span. We loop all cells here since the cells are sorted with columns first.
                var cse = new CellsStoreEnumerator<ExcelCoreValue>(ws._values, 1, address._fromCol, address._toRow, address._toCol);
                while (cse.Next())
                {
                    if (cse.Column >= address.Start.Column && cse.Column <= address.End.Column && cse.Value._styleId != 0)
                    {
                        if (styleCashe.ContainsKey(cse.Value._styleId))
                        {
                            ws.SetStyleInner(cse.Row, cse.Column, styleCashe[cse.Value._styleId]);
                        }
                        else
                        {
                            ExcelXfs st = CellXfs[cse.Value._styleId];
                            var newId = st.GetNewID(CellXfs, sender, e.StyleClass, e.StyleProperty, e.Value);
                            styleCashe.Add(cse.Value._styleId, newId);
                            ws.SetStyleInner(cse.Row, cse.Column, newId);
                        }
                    }
                }

                if (!(address._fromCol == 1 && address._toCol == ExcelPackage.MaxColumns))
                {
                    //Update cells with styled columns
                    cse = new CellsStoreEnumerator<ExcelCoreValue>(ws._values, 1, 0, address._toRow, 0);
                    while (cse.Next())
                    {
                        if (cse.Value._styleId == 0)
                        {
                            continue;
                        }

                        for (var c = address._fromCol; c <= address._toCol; c++)
                        {
                            if (!ws.ExistsStyleInner(cse.Row, c))
                            {
                                if (styleCashe.ContainsKey(cse.Value._styleId))
                                {
                                    ws.SetStyleInner(cse.Row, c, styleCashe[cse.Value._styleId]);
                                }
                                else
                                {
                                    ExcelXfs st = CellXfs[cse.Value._styleId];
                                    var newId = st.GetNewID(CellXfs, sender, e.StyleClass, e.StyleProperty, e.Value);
                                    styleCashe.Add(cse.Value._styleId, newId);
                                    ws.SetStyleInner(cse.Row, c, newId);
                                }
                            }
                        }
                    }
                }
            }

            //Rows
            else if (address.Start.Column == 1 && address.End.Column == ExcelPackage.MaxColumns)
            {
                for (var rowNum = address.Start.Row; rowNum <= address.End.Row; rowNum++)
                {
                    var s = ws.GetStyleInner(rowNum, 0);
                    if (s == 0)
                    {
                        //iterate all columns and set the row to the style of the last column
                        var cse = new CellsStoreEnumerator<ExcelCoreValue>(ws._values, 0, 1, 0, ExcelPackage.MaxColumns);
                        while (cse.Next())
                        {
                            s = cse.Value._styleId;
                            if (s == 0)
                            {
                                continue;
                            }

                            var c = ws.GetValueInner(cse.Row, cse.Column) as ExcelColumn;
                            if (c != null && c.ColumnMax < ExcelPackage.MaxColumns)
                            {
                                for (var col = c.ColumnMin; col < c.ColumnMax; col++)
                                {
                                    if (!ws.ExistsStyleInner(rowNum, col))
                                    {
                                        ws.SetStyleInner(rowNum, col, s);
                                    }
                                }
                            }
                        }

                        ws.SetStyleInner(rowNum, 0, s);
                        cse.Dispose();
                    }

                    if (styleCashe.ContainsKey(s))
                    {
                        ws.SetStyleInner(rowNum, 0, styleCashe[s]);
                    }
                    else
                    {
                        ExcelXfs st = CellXfs[s];
                        var newId = st.GetNewID(CellXfs, sender, e.StyleClass, e.StyleProperty, e.Value);
                        styleCashe.Add(s, newId);
                        ws.SetStyleInner(rowNum, 0, newId);
                    }
                }

                //Update individual cells 
                var cse2 = new CellsStoreEnumerator<ExcelCoreValue>(ws._values, address._fromRow, address._fromCol, address._toRow, address._toCol);
                while (cse2.Next())
                {
                    var s = cse2.Value._styleId;
                    if (s == 0)
                    {
                        continue;
                    }

                    if (styleCashe.ContainsKey(s))
                    {
                        ws.SetStyleInner(cse2.Row, cse2.Column, styleCashe[s]);
                    }
                    else
                    {
                        ExcelXfs st = CellXfs[s];
                        var newId = st.GetNewID(CellXfs, sender, e.StyleClass, e.StyleProperty, e.Value);
                        styleCashe.Add(s, newId);
                        ws.SetStyleInner(cse2.Row, cse2.Column, newId);
                    }
                }

                //Update cells with styled rows
                cse2 = new CellsStoreEnumerator<ExcelCoreValue>(ws._values, 0, 1, 0, address._toCol);
                while (cse2.Next())
                {
                    if (cse2.Value._styleId == 0)
                    {
                        continue;
                    }

                    for (var r = address._fromRow; r <= address._toRow; r++)
                    {
                        if (!ws.ExistsStyleInner(r, cse2.Column))
                        {
                            var s = cse2.Value._styleId;
                            if (styleCashe.ContainsKey(s))
                            {
                                ws.SetStyleInner(r, cse2.Column, styleCashe[s]);
                            }
                            else
                            {
                                ExcelXfs st = CellXfs[s];
                                var newId = st.GetNewID(CellXfs, sender, e.StyleClass, e.StyleProperty, e.Value);
                                styleCashe.Add(s, newId);
                                ws.SetStyleInner(r, cse2.Column, newId);
                            }
                        }
                    }
                }
            }
            else //Cell range
            {
                var tmpCache = styleCashe;
                var rowCache = new Dictionary<int, int>(address.End.Row - address.Start.Row + 1);
                var colCache = new Dictionary<int, ExcelCoreValue>(address.End.Column - address.Start.Column + 1);
                ws._values.SetRangeValueSpecial(address.Start.Row, address.Start.Column, address.End.Row, address.End.Column,
                    (list, index, row, column, args) =>
                    {
                        // Optimized GetStyleID
                        var s = list[index]._styleId;
                        if (s == 0 && !ws.ExistsStyleInner(row, 0, ref s))
                        {
                            // get row styleId with cache
                            if (!rowCache.ContainsKey(row))
                            {
                                rowCache.Add(row, ws._values.GetValue(row, 0)._styleId);
                            }

                            s = rowCache[row];
                            if (s == 0)
                            {
                                // get column styleId with cache
                                if (!colCache.ContainsKey(column))
                                {
                                    colCache.Add(column, ws._values.GetValue(0, column));
                                }

                                s = colCache[column]._styleId;
                                if (s == 0)
                                {
                                    int r = 0, c = column;
                                    if (ws._values.PrevCell(ref r, ref c))
                                    {
                                        //var val = ws._values.GetValue(0, c);
                                        if (!colCache.ContainsKey(c))
                                        {
                                            colCache.Add(c, ws._values.GetValue(0, c));
                                        }

                                        ExcelCoreValue val = colCache[c];
                                        var colObj = (ExcelColumn) val._value;
                                        if (colObj != null && colObj.ColumnMax >= column) //Fixes issue 15174
                                        {
                                            s = val._styleId;
                                        }
                                    }
                                }
                            }
                        }

                        if (tmpCache.ContainsKey(s))
                        {
                            //ws.SetStyleInner(row, column, tmpCache[s]);
                            list[index] = new ExcelCoreValue {_value = list[index]._value, _styleId = tmpCache[s]};
                        }
                        else
                        {
                            ExcelXfs st = CellXfs[s];
                            var newId = st.GetNewID(CellXfs, sender, e.StyleClass, e.StyleProperty, e.Value);
                            tmpCache.Add(s, newId);
                            //ws.SetStyleInner(row, column, newId);
                            list[index] = new ExcelCoreValue {_value = list[index]._value, _styleId = newId};
                        }
                    }, e);
            }
        }

        #region XmlHelpFunctions
        private static int GetXmlNodeInt(XmlNode node)
        {
            return int.TryParse(GetXmlNode(node), out var i) ? i : 0;
        }

        private static string GetXmlNode(XmlNode node)
        {
            return node?.Value ?? "";
        }
        #endregion
    }
}
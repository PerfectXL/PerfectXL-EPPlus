﻿/*******************************************************************************
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
 * Mats Alm   		                Added       		        2013-03-01 (Prior file history on https://github.com/swmal/ExcelFormulaParser)
 *******************************************************************************/
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;

namespace OfficeOpenXml.FormulaParsing.ExcelUtilities
{
    public class CellReferenceProvider
    {
        public virtual IEnumerable<string> GetReferencedAddresses(string cellFormula, ParsingContext context)
        {
            var resultCells = new List<string>();
            var r = context.Configuration.Lexer.Tokenize(cellFormula, context.Scopes.Current.Address.Worksheet);
            var toAddresses = r.Where(x => x.TokenType == TokenType.ExcelAddress);
            foreach (var toAddress in toAddresses)
            {
                var rangeAddress = context.RangeAddressFactory.Create(toAddress.Value);
                var rangeCells = new List<string>();
                if (rangeAddress.FromRow < rangeAddress.ToRow || rangeAddress.FromCol < rangeAddress.ToCol)
                {
                    for (var col = rangeAddress.FromCol; col <= rangeAddress.ToCol; col++)
                    {
                        for (var row = rangeAddress.FromRow; row <= rangeAddress.ToRow; row++)
                        {
                            resultCells.Add(context.RangeAddressFactory.Create(col, row).Address);
                        }
                    }
                }
                else
                {
                    rangeCells.Add(toAddress.Value);
                }
                resultCells.AddRange(rangeCells);
            }
            return resultCells;
        }
    }
}

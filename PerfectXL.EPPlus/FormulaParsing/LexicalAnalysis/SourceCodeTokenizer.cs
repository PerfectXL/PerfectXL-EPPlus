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
 * Mats Alm   		                Added       		        2013-03-01 (Prior file history on https://github.com/swmal/ExcelFormulaParser)
 *******************************************************************************/

using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.Utilities;

namespace OfficeOpenXml.FormulaParsing.LexicalAnalysis
{
    public class SourceCodeTokenizer : ISourceCodeTokenizer
    {
        private static readonly TokenType[] _sheetReferenceTokens = { TokenType.ExcelAddressR1C1, TokenType.ExcelAddress, TokenType.NameValue, TokenType.Unrecognized };
        public static ISourceCodeTokenizer Default
        {
            get { return new SourceCodeTokenizer(FunctionNameProvider.Empty, NameValueProvider.Empty, false); }
        }
        public static ISourceCodeTokenizer R1C1
        {
            get { return new SourceCodeTokenizer(FunctionNameProvider.Empty, NameValueProvider.Empty, true); }
        }


        public SourceCodeTokenizer(IFunctionNameProvider functionRepository, INameValueProvider nameValueProvider, bool r1c1=false)
            : this(new TokenFactory(functionRepository, nameValueProvider, r1c1), new TokenSeparatorProvider())
        {

        }
        public SourceCodeTokenizer(ITokenFactory tokenFactory, ITokenSeparatorProvider tokenProvider)
        {
            _tokenFactory = tokenFactory;
            _separatorProvider = tokenProvider;
        }

        private readonly ITokenSeparatorProvider _separatorProvider;
        private readonly ITokenFactory _tokenFactory;

        public IEnumerable<Token> Tokenize(string input)
        {
            return Tokenize(input, null);
        }
        public IEnumerable<Token> Tokenize(string input, string worksheet)
        {
            if (string.IsNullOrEmpty(input))
            {
                return Enumerable.Empty<Token>();
            }
            var context = new TokenizerContext(input);
            var handler = new TokenHandler(context, _tokenFactory, _separatorProvider);
            handler.Worksheet = worksheet;
            while(handler.HasMore())
            {
                handler.Next();
            }
            if (context.CurrentTokenHasValue)
            {
                context.AddToken(CreateToken(context, worksheet));
            }

            HandleUnrecognizedTokens(context, _separatorProvider.Tokens);

            return context.Result;
        }

        


        private static void HandleUnrecognizedTokens(TokenizerContext context, IDictionary<string, Token>  tokens)
        {
            int i = 0;
            while (i < context.Result.Count)
            {
                var token=context.Result[i];
                
                if (token.TokenType == TokenType.Unrecognized)
                {
                    //Check 3-D sheet reference, Note: 3-D reference with single sheet name will be seen as a single sheet name
                    if (i < context.Result.Count - 4 && context.Result[i+1].TokenType == TokenType.Colon && context.Result[i + 2].TokenType == TokenType.Unrecognized 
                            && context.Result[i + 3].TokenType == TokenType.ExclamationMark && _sheetReferenceTokens.Contains(context.Result[i + 4].TokenType))
                    {
                        token.TokenType = TokenType.WorksheetName;
                        context.Result[i + 2].TokenType = TokenType.WorksheetName;
                        i += 4;
                        continue;
                    }

                    //Check regular sheet reference
                    if (i < context.Result.Count - 2 && context.Result[i + 1].TokenType == TokenType.ExclamationMark && _sheetReferenceTokens.Contains(context.Result[i + 2].TokenType))
                    {
                        token.TokenType = TokenType.WorksheetName;
                        i += 2;
                        continue;
                    }

                    //Check for function
                    if (i < context.Result.Count - 1 && context.Result[i + 1].TokenType == TokenType.OpeningParenthesis)
                    {
                        token.TokenType = TokenType.Function;
                        i += 2;
                        continue;
                    }

                    //Check for Column / Row reference
                    if (i < context.Result.Count - 2 && context.Result[i + 1].TokenType == TokenType.Colon && context.Result[i + 2].TokenType == TokenType.Unrecognized 
                        && IsColumnOrRowReference($"{token.Value}:{context.Result[i + 2].Value}"))
                    {
                        token.TokenType = TokenType.ExcelAddress;
                        context.Result[i + 2].TokenType = TokenType.ExcelAddress;
                        i += 3;
                        continue;
                    }
                    
                    token.TokenType = TokenType.NameValue;
                }

                i++;
            }
        }

        private static bool IsColumnOrRowReference(string address)
        {
            return Regex.IsMatch(address, RegexConstants.ColumnReferencePattern, RegexOptions.IgnorePatternWhitespace) 
                   || Regex.IsMatch(address, RegexConstants.RowReferencePattern, RegexOptions.IgnorePatternWhitespace);
        }

        private Token CreateToken(TokenizerContext context, string worksheet)
        {
            if (context.CurrentToken == "-")
            {
                if (context.LastToken == null && context.LastToken.TokenType == TokenType.Operator)
                {
                    return new Token("-", TokenType.Negator);
                }
            }
            return _tokenFactory.Create(context.Result, context.CurrentToken, worksheet);
        }
    }
}

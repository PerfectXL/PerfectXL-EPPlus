using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;

namespace PerfectXL.EPPlus.UnitTests
{
    [TestClass]
    public class TableReferenceTest
    {
        private Lexer _lexer;

        [TestInitialize]
        public void Initialize()
        {
            _lexer = new Lexer(SourceCodeTokenizer.R1C1, new SyntacticAnalyzer());
        }

        [DataTestMethod]
        [DataRow("Table1[Data1]")]
        [DataRow("Table1['#Headers]")]
        [DataRow("Table1[[Data1]:['#Headers]]")]
        [DataRow("Table1[   [Data1]:['#Headers]   ]")]
        [DataRow("Table1[[#This Row],[Data1]:[Data3]]")]
        [DataRow("Table1[   [#This Row]  ,  [Data1]:[Data3]   ]")]
        [DataRow("Table1[#Totals]")]
        [DataRow("Table1[[#Data],[Data1]]")]
        [DataRow("Table1[[#Data]  , [Data1]]")]
        [DataRow("Table1[#This Row]")]
        public void ValidTableReferences(string reference)
        {
            var tokens = _lexer.Tokenize(reference, null).ToList();
            Assert.IsTrue(tokens.Count == 1 && tokens[0].TokenType == TokenType.TableReference, $"{reference} should be a valid table reference");
        }

        [DataTestMethod]
        [DataRow("Table1[[Data1]: ['#Headers]]")]
        [DataRow("Table1[Da ta1]")]
        [DataRow("Table1 [Data1]")]
        [DataRow("Table1[[Data1]:[#This Row]]")]
        public void InvalidTableReferences(string reference)
        {
            var tokens = _lexer.Tokenize(reference, null).ToList();
            Assert.IsTrue(tokens.All(x => x.TokenType != TokenType.TableReference), $"{reference} should not be a valid table reference");
        }
    }
}

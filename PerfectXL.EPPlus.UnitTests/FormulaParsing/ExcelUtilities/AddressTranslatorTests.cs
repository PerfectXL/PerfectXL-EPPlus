using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing;
using FakeItEasy;

namespace EPPlusTest.ExcelUtilities
{
    [TestClass]
    public class AddressTranslatorTests
    {
        private AddressTranslator _addressTranslator;
        private ExcelDataProvider _excelDataProvider;
        private const int ExcelMaxRows = 1356;

        [TestInitialize]
        public void Setup()
        {
            _excelDataProvider = A.Fake<ExcelDataProvider>();
            A.CallTo(() => _excelDataProvider.ExcelMaxRows).Returns(ExcelMaxRows);
            _addressTranslator = new AddressTranslator(_excelDataProvider);
        }

        [TestMethod, ExpectedException(typeof(ArgumentNullException))]
        public void ConstructorShouldThrowIfProviderIsNull()
        {
            new AddressTranslator(null);
        }

        [TestMethod]
        public void ShouldTranslateRowNumber()
        {
            _addressTranslator.ToColAndRow("A2", out var col, out var row);
            Assert.AreEqual(2, row);
        }

        [TestMethod]
        public void ShouldTranslateLettersToColumnIndex()
        {
            _addressTranslator.ToColAndRow("C1", out var col, out var row);
            Assert.AreEqual(3, col);
            _addressTranslator.ToColAndRow("AA2", out col, out row);
            Assert.AreEqual(27, col);
            _addressTranslator.ToColAndRow("BC1", out col, out row);
            Assert.AreEqual(55, col);
        }

        [TestMethod]
        public void ShouldTranslateLetterAddressUsingMaxRowsFromProviderLower()
        {
            _addressTranslator.ToColAndRow("A", out var col, out var row);
            Assert.AreEqual(1, row);
        }

        [TestMethod]
        public void ShouldTranslateLetterAddressUsingMaxRowsFromProviderUpper()
        {
            _addressTranslator.ToColAndRow("A", out var col, out var row, AddressTranslator.RangeCalculationBehaviour.LastPart);
            Assert.AreEqual(ExcelMaxRows, row);
        }
    }
}

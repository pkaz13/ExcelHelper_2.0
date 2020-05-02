using ExcelHelper_2._0.Importers;
using ExcelHelper_2._0.Wrappers;
using ExcelHelper_2.Creators;
using NUnit.Framework;
using System.Collections.Generic;
using System.Linq;

namespace ExcelHelperTests.Integration
{
    [TestFixture]
    class ExcelImporterTests
    {
        IExcelCreator _excelCreator;
        IFileStreamWrapper _fileStreamWrapper;
        IExcelImporter<TestExcelObject> _excelImporter;

        [SetUp]
        public void SetUp()
        {
            _excelCreator = new ExcelCreator();
            _fileStreamWrapper = new FileStreamWrapper();
            _excelImporter = new ExcelImporter<TestExcelObject>(_excelCreator, _fileStreamWrapper);
        }

        [Test]
        public void Test_Should_Import_File()
        {
            string path = @"..\..\..\..\test.xlsx";

            IEnumerable<TestExcelObject> result = _excelImporter.Import(path).ToList();

            Assert.IsNotNull(result);
        }
    }
}

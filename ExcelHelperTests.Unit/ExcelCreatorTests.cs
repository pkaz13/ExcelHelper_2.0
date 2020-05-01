using ExcelHelper_2;
using ExcelHelper_2._0.Exceptions;
using ExcelHelper_2.Creators;
using NUnit.Framework;
using System.Collections.Generic;

namespace ExcelHelperTests.Unit
{
    [TestFixture]
    class ExcelCreatorTests
    {
        ExcelCreator _excelCreator;

        [SetUp]
        public void SetUp()
        {
            _excelCreator = new ExcelCreator();
        }

        [Test]
        public void Test_Create_Should_Create_ExcelFile()
        {
            List<TestExcelObject> testExcelObjects = new List<TestExcelObject>
            {
                new TestExcelObject("name1", 1),
                new TestExcelObject("name2", 2)
            };
            List<string> columnNames = new List<string>
            {
                "name",
                "value"
            };
            string sheetName = "sheetName";

            IExcelFile excelFile = _excelCreator.Create(testExcelObjects, sheetName, columnNames);

            Assert.IsNotNull(excelFile);
        }

        [Test]
        public void Test_Create_Should_Throw_ExcelCreationException()
        {
            List<TestExcelObject> testExcelObjects = null;
            List<string> columnNames = new List<string>
            {
                "name",
                "value"
            };
            string sheetName = "sheetName";

            Assert.Throws<ExcelCreationException>(() => _excelCreator.Create(testExcelObjects, sheetName, columnNames));
        }
    }
}

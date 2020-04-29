using ExcelHelper_2.Creators;
using ExcelHelper_2.Exporters;
using NUnit.Framework;
using System.Collections.Generic;

namespace ExcelHelperTests.Integration
{
    [TestFixture]
    class ExcelExporterTests
    {
        IExcelCreator _excelCreator;
        IExcelExporter _excelExporter;

        const string path = @"..\..\..\..\test.xlsx";

        [SetUp]
        public void SetUp()
        {
            _excelCreator = new ExcelCreator();
            _excelExporter = new ExcelExporter(_excelCreator);
        }

        [Test]
        public void Test_Should_Export_Object_To_Excel()
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

            _excelExporter.Export(testExcelObjects, columnNames, path);
        }
    }
}

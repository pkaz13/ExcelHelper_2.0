using ExcelHelper_2;
using ExcelHelper_2.Creators;
using ExcelHelper_2.Exporters;
using Moq;
using NUnit.Framework;
using System.Collections.Generic;
using System.IO;

namespace ExcelHelperTests.Unit
{
    [TestFixture]
    class ExcelExporterTests
    {
        Mock<IExcelCreator> _excelCreatorMock;
        Mock<IExcelFile> _excelFileMock;
        ExcelExporter _excelExporter;

        [SetUp]
        public void SetUp()
        {
            _excelCreatorMock = new Mock<IExcelCreator>();
            _excelFileMock = new Mock<IExcelFile>();
            _excelExporter = new ExcelExporter(_excelCreatorMock.Object);
        }

        [Test]
        public void Test_Should_Export_Excel_File()
        {
            _excelCreatorMock.Setup(x => x.Create(It.IsAny<IEnumerable<TestExcelObject>>(), It.IsAny<string>(), It.IsAny<List<string>>())).Returns(_excelFileMock.Object);
            List<TestExcelObject> testExcelObjects = new List<TestExcelObject>
            {
                new TestExcelObject("name1", 1),
                new TestExcelObject("name2", 2)
            };

            _excelExporter.Export(testExcelObjects, new List<string> { "name", "value" }, @"test\path\file.xlsx");

            _excelFileMock.Verify(x => x.Save(It.IsAny<FileInfo>()), Times.Once);
        }
    }
}

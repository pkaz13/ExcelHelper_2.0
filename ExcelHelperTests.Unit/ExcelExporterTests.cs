using ExcelHelper_2;
using ExcelHelper_2._0.Exceptions;
using ExcelHelper_2.Creators;
using ExcelHelper_2.Exporters;
using Moq;
using NUnit.Framework;
using System;
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

        [Test]
        public void Test_Export_Should_Throw_ArgumentException_When_Collection_To_Export_Is_Null()
        {
            List<TestExcelObject> testExcelObjects = null;
            List<string> columnNames = new List<string>
            {
                "name",
                "value"
            };
            string path = @"test\path\file.xlsx";

            _excelFileMock.Verify(x => x.Save(It.IsAny<FileInfo>()), Times.Never);
            Assert.Throws<ArgumentException>(() => _excelExporter.Export(testExcelObjects, columnNames, path));
        }

        [Test]
        public void Test_Export_Should_Throw_ArgumentException_When_Collection_To_Export_Is_Empty()
        {
            List<TestExcelObject> testExcelObjects = new List<TestExcelObject>();
            List<string> columnNames = new List<string>
            {
                "name",
                "value"
            };
            string path = @"test\path\file.xlsx";

            _excelFileMock.Verify(x => x.Save(It.IsAny<FileInfo>()), Times.Never);
            Assert.Throws<ArgumentException>(() => _excelExporter.Export(testExcelObjects, columnNames, path));
        }

        [Test]
        public void Test_Export_Should_Throw_ArgumentException_When_ColumnNames_Are_Null()
        {
            List<TestExcelObject> testExcelObjects = new List<TestExcelObject>
            {
                new TestExcelObject("name1", 1),
                new TestExcelObject("name2", 2)
            };
            List<string> columnNames = null;
            string path = @"test\path\file.xlsx";

            _excelFileMock.Verify(x => x.Save(It.IsAny<FileInfo>()), Times.Never);
            Assert.Throws<ArgumentException>(() => _excelExporter.Export(testExcelObjects, columnNames, path));
        }

        [Test]
        public void Test_Export_Should_Throw_ArgumentException_When_ColumnNames_Are_Empty()
        {
            List<TestExcelObject> testExcelObjects = new List<TestExcelObject>
            {
                new TestExcelObject("name1", 1),
                new TestExcelObject("name2", 2)
            };
            List<string> columnNames = new List<string>();
            string path = @"test\path\file.xlsx";

            _excelFileMock.Verify(x => x.Save(It.IsAny<FileInfo>()), Times.Never);
            Assert.Throws<ArgumentException>(() => _excelExporter.Export(testExcelObjects, columnNames, path));
        }

        [Test]
        public void Test_Export_Should_Throw_ArgumentException_When_Path_Is_Null()
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
            string path = null;

            _excelFileMock.Verify(x => x.Save(It.IsAny<FileInfo>()), Times.Never);
            Assert.Throws<ArgumentException>(() => _excelExporter.Export(testExcelObjects, columnNames, path));
        }

        [Test]
        public void Test_Export_Should_Throw_ArgumentException_When_Path_Is_Empty()
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
            string path = string.Empty;

            _excelFileMock.Verify(x => x.Save(It.IsAny<FileInfo>()), Times.Never);
            Assert.Throws<ArgumentException>(() => _excelExporter.Export(testExcelObjects, columnNames, path));
        }

        [Test]
        public void Test_Export_Should_Throw_ExcelCreationException_When_ExcelCreator_Throws_Exception()
        {
            _excelCreatorMock.Setup(x => x.Create(It.IsAny<IEnumerable<TestExcelObject>>(), It.IsAny<string>(), It.IsAny<List<string>>())).Throws(new ExcelCreationException());
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
            string path = @"test\path\file.xlsx";

            Assert.Throws<ExcelCreationException>(() => _excelExporter.Export(testExcelObjects, columnNames, path));
        }

        [Test]
        public void Test_Export_Should_Throw_ExcelCreationException_When_ExcelFile_Throws_Exception()
        {
            _excelFileMock.Setup(x => x.Save(It.IsAny<FileInfo>())).Throws(new ExcelCreationException());
            _excelCreatorMock.Setup(x => x.Create(It.IsAny<IEnumerable<TestExcelObject>>(), It.IsAny<string>(), It.IsAny<List<string>>())).Returns(_excelFileMock.Object);
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
            string path = @"test\path\file.xlsx";

            Assert.Throws<ExcelCreationException>(() => _excelExporter.Export(testExcelObjects, columnNames, path));
        }
    }
}

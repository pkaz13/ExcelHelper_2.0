using ExcelHelper_2;
using ExcelHelper_2._0.Exceptions;
using ExcelHelper_2._0.Importers;
using ExcelHelper_2._0.Wrappers;
using ExcelHelper_2.Creators;
using Moq;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.IO;

namespace ExcelHelperTests.Unit
{
    [TestFixture]
    class ExcelImporterTests
    {
        Mock<IFileStreamWrapper> _fileStreamWrapperMock;
        Mock<IExcelCreator> _excelCreatorMock;
        Mock<IExcelFile> _excelFileMock;
        ExcelImporter<TestExcelObject> _excelImporter;

        [SetUp]
        public void SetUp()
        {
            _fileStreamWrapperMock = new Mock<IFileStreamWrapper>();
            _excelCreatorMock = new Mock<IExcelCreator>();
            _excelFileMock = new Mock<IExcelFile>();
            _excelImporter = new ExcelImporter<TestExcelObject>(_excelCreatorMock.Object, _fileStreamWrapperMock.Object);
        }

        [Test]
        public void Test_Import_Should_Create_TestExcelObject()
        {
            _excelFileMock.Setup(x => x.ConvertToObjects<TestExcelObject>()).Returns(new List<TestExcelObject>());
            _excelCreatorMock.Setup(x => x.Create(It.IsAny<FileStream>())).Returns(_excelFileMock.Object);
            string path = @"fake\path.xlsx";

            IEnumerable<TestExcelObject> testExcelObjects = _excelImporter.Import(path);

            Assert.IsNotNull(testExcelObjects);
            _excelFileMock.Verify(x => x.ConvertToObjects<TestExcelObject>(), Times.Once);
        }

        [Test]
        public void Test_Import_Should_Throw_ArgumentException_When_Path_Is_Null()
        {
            string path = null;

            Assert.Throws<ArgumentException>(() => _excelImporter.Import(path));
        }

        [Test]
        public void Test_Import_Should_Throw_ArgumentException_When_Path_Is_Empty()
        {
            string path = string.Empty;

            Assert.Throws<ArgumentException>(() => _excelImporter.Import(path));
        }

        [Test]
        public void Test_Should_Throw_ExcelCreationException_When_ExcelCreator_Throws_Exception()
        {
            _excelCreatorMock.Setup(x => x.Create(It.IsAny<FileStream>())).Throws(new ExcelCreationException());
            string path = @"fake\path.xlsx";

            Assert.Throws<ExcelCreationException>(() => _excelImporter.Import(path));
        }

        [Test]
        public void Test_Should_Throw_ExcelConversionException_When_ExcelFile_Throws_Exception()
        {
            _excelFileMock.Setup(x => x.ConvertToObjects<TestExcelObject>()).Throws(new ExcelConvertionException());
            _excelCreatorMock.Setup(x => x.Create(It.IsAny<FileStream>())).Returns(_excelFileMock.Object);
            string path = @"fake\path.xlsx";

            Assert.Throws<ExcelConvertionException>(() => _excelImporter.Import(path));
        }
    }
}

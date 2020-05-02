using ExcelHelper_2._0.Wrappers;
using ExcelHelper_2.Creators;
using System;
using System.Collections.Generic;
using System.IO;

namespace ExcelHelper_2._0.Importers
{
    public class ExcelImporter<T> : IExcelImporter<T> where T : new()
    {
        private readonly IExcelCreator _excelCreator;
        private readonly IFileStreamWrapper _fileStreamWrapper;

        public ExcelImporter(IExcelCreator excelCreator, IFileStreamWrapper fileStreamWrapper)
        {
            _excelCreator = excelCreator;
            _fileStreamWrapper = fileStreamWrapper;
        }

        /// <summary>
        /// Returns list of T based on *.xlsx worksheet.
        /// </summary>
        public IEnumerable<T> Import(string path)
        {
            if (string.IsNullOrEmpty(path))
            {
                throw new ArgumentException("Path cannot be null or empty");
            }

            IEnumerable<T> collectionToReturn = new List<T>();
            using (FileStream fileStream = _fileStreamWrapper.Init(path, FileMode.Open))
            using (IExcelFile excelFile = _excelCreator.Create(fileStream))
            {
                collectionToReturn = new List<T>(excelFile.ConvertToObjects<T>());
            }
            return collectionToReturn;
        }
    }
}

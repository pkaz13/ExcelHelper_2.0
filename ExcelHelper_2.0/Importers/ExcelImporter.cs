using ExcelHelper_2.Creators;
using System;
using System.Collections.Generic;
using System.IO;

namespace ExcelHelper_2._0.Importers
{
    public class ExcelImporter<T> : IExcelImporter<T> where T : new()
    {
        private readonly IExcelCreator _excelCreator;

        public ExcelImporter(IExcelCreator excelCreator)
        {
            _excelCreator = excelCreator;
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

            using (FileStream fileStream = new FileStream(path, FileMode.Open))
            using (IExcelFile excelFile = _excelCreator.Create(fileStream))
            {
                return excelFile.ConvertToObjects<T>();
            }
        }
    }
}

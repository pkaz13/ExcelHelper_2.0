using ExcelHelper_2.Creators;
using ExcelHelper_2.Utils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ExcelHelper_2.Exporters
{
    public class ExcelExporter : IExcelExporter
    {
        private const string sheetName = "Sheet1";

        private readonly IExcelCreator _excelCreator;

        public ExcelExporter(IExcelCreator excelCreator)
        {
            _excelCreator = excelCreator;
        }

        /// <summary>
        /// Creates Excel file based on T and saves it to specified path.
        /// </summary>
        public void Export<T>(IEnumerable<T> collection, List<string> columnNames, string path)
        {
            if (IsCollectionNullOrEmpty(collection) || AreColumnNamesNullOrEmpty(columnNames) || string.IsNullOrEmpty(path))
            {
                throw new ArgumentException("One of the arguments is null or empty.");
            }

            FileInfo fileInfo = CreateFileInfo(path);
            using (IExcelFile excelFile = _excelCreator.Create(collection, sheetName, columnNames))
            {
                excelFile.Save(fileInfo);
            }
        }

        private FileInfo CreateFileInfo(string path)
        {
            FileInfo file = new FileInfo(path);
            return FileHelper.DeleteFileIfExistAndCreateNewFile(file);
        }

        private bool IsCollectionNullOrEmpty<T>(IEnumerable<T> collection)
        {
            return collection == null || collection.Count() == 0 ? true : false;
        }

        private bool AreColumnNamesNullOrEmpty(List<string> columnNames)
        {
            return columnNames == null || columnNames.Count == 0 ? true : false;
        }
    }
}

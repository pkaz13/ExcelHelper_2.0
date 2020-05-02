using ExcelHelper_2._0.Exceptions;
using ExcelHelper_2.Utils;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;

namespace ExcelHelper_2.Creators
{
    public class ExcelCreator : IExcelCreator
    {
        private const string startCell = "A1";

        /// <summary>
        /// Creates Excel package based on T.
        /// </summary>
        public IExcelFile Create<T>(IEnumerable<T> collection, string worksheetName, List<string> columnNames)
        {
            try
            {
                ExcelPackage excelFile = new ExcelPackage();
                ExcelWorksheet worksheet = excelFile.Workbook.Worksheets.Add(worksheetName);
                worksheet.Cells[startCell].LoadFromCollectionFiltered(collection);

                SetColumnNames(worksheet, columnNames);

                double minimumWidth = 0;
                worksheet.Cells.AutoFitColumns(minimumWidth);
                return new ExcelFile(excelFile);
            }
            catch (Exception ex)
            {
                throw new ExcelCreationException(ex.Message, ex);
            }
        }

        public IExcelFile Create(FileStream fileStream)
        {
            try
            {
                ExcelPackage excelFile = new ExcelPackage(fileStream);
                return new ExcelFile(excelFile);
            }
            catch (Exception ex)
            {
                throw new ExcelCreationException(ex.Message, ex);
            }
        }

        private void SetColumnNames(ExcelWorksheet worksheet, List<string> columnNames)
        {
            for (int i = 0; i < columnNames.Count; i++)
            {
                worksheet.Cells[1, i + 1].Value = columnNames[i];
            }
        }
    }
}

﻿using ExcelHelper_2._0.Utils;
using OfficeOpenXml;
using System.Collections.Generic;

namespace ExcelHelper_2._0
{
    class ExcelCreator : IExcelCreator
    {
        private const string startCell = "A1";

        public IExcelFile Create<T>(IEnumerable<T> collection, string worksheetName, List<string> columnNames)
        {
            using (ExcelPackage excelFile = new ExcelPackage())
            {
                ExcelWorksheet worksheet = excelFile.Workbook.Worksheets.Add(worksheetName);
                worksheet.Cells[startCell].LoadFromCollectionFiltered(collection);

                SetColumnNames(worksheet, columnNames);

                double minimumWidth = 0;
                worksheet.Cells.AutoFitColumns(minimumWidth);
                return new ExcelFie(excelFile);
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
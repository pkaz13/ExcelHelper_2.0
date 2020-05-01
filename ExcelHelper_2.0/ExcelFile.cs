using ExcelHelper_2._0.Exceptions;
using OfficeOpenXml;
using System;
using System.IO;

namespace ExcelHelper_2
{
    class ExcelFile : IExcelFile
    {
        private readonly ExcelPackage _excel;

        public ExcelFile(ExcelPackage excel)
        {
            _excel = excel;
        }

        public void Save(FileInfo fileInfo)
        {
            try
            {
                _excel.SaveAs(fileInfo);
            }
            catch (Exception ex)
            {
                throw new ExcelCreationException(ex.Message, ex);
            }

        }

        public void Dispose()
        {
            _excel.Dispose();
        }
    }
}

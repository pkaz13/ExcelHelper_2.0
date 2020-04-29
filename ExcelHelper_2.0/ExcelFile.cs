using OfficeOpenXml;
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
            _excel.SaveAs(fileInfo);
        }

        public void Dispose()
        {
            _excel.Dispose();
        }
    }
}

using OfficeOpenXml;
using System.IO;

namespace ExcelHelper_2
{
    class ExcelFie : IExcelFile
    {
        private readonly ExcelPackage _excel;

        public ExcelFie(ExcelPackage excel)
        {
            _excel = excel;
        }

        public void Save(FileInfo fileInfo)
        {
            _excel.SaveAs(fileInfo);
        }
    }
}

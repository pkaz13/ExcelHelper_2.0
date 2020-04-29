using OfficeOpenXml;

namespace ExcelHelper_2
{
    class ExcelFie : IExcelFile
    {
        public ExcelPackage Excel { get; }

        public ExcelFie(ExcelPackage excel)
        {
            Excel = excel;
        }
    }
}

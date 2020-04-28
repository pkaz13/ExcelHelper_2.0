using OfficeOpenXml;

namespace ExcelHelper_2._0
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

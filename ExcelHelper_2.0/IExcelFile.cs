using OfficeOpenXml;

namespace ExcelHelper_2._0
{
    public interface IExcelFile
    {
        ExcelPackage Excel { get; }
    }
}

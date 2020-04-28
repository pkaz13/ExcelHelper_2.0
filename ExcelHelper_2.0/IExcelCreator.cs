using System.Collections.Generic;

namespace ExcelHelper_2._0
{
    public interface IExcelCreator
    {
        IExcelFile Create<T>(IEnumerable<T> collection, string worksheetName, List<string> columnNames);
    }
}

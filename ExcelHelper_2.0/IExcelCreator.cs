using System.Collections.Generic;

namespace ExcelHelper_2
{
    public interface IExcelCreator
    {
        IExcelFile Create<T>(IEnumerable<T> collection, string worksheetName, List<string> columnNames);
    }
}

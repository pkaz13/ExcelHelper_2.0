using System.Collections.Generic;

namespace ExcelHelper_2.Creators
{
    public interface IExcelCreator
    {
        IExcelFile Create<T>(IEnumerable<T> collection, string worksheetName, List<string> columnNames);
    }
}

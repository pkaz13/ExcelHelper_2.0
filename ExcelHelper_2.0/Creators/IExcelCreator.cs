using System.Collections.Generic;
using System.IO;

namespace ExcelHelper_2.Creators
{
    public interface IExcelCreator
    {
        IExcelFile Create<T>(IEnumerable<T> collection, string worksheetName, List<string> columnNames);
        IExcelFile Create(FileStream fileStream);
    }
}

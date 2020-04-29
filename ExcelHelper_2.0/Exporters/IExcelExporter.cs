using System.Collections.Generic;

namespace ExcelHelper_2.Exporters
{
    public interface IExcelExporter
    {
        void Export<T>(IEnumerable<T> collection, List<string> columnNames, string path);
    }
}

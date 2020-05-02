using System.Collections.Generic;

namespace ExcelHelper_2._0.Importers
{
    public interface IExcelImporter<T>
    {
        IEnumerable<T> Import(string path);
    }
}

using System;
using System.Collections.Generic;
using System.IO;

namespace ExcelHelper_2
{
    public interface IExcelFile : IDisposable
    {
        void Save(FileInfo fileInfo);
        IEnumerable<T> ConvertToObjects<T>() where T : new();
    }
}

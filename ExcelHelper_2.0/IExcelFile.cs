using System;
using System.IO;

namespace ExcelHelper_2
{
    public interface IExcelFile : IDisposable
    {
        void Save(FileInfo fileInfo);
    }
}

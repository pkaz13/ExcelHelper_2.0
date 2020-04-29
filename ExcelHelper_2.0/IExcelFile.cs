using System.IO;

namespace ExcelHelper_2
{
    public interface IExcelFile
    {
        void Save(FileInfo fileInfo);
    }
}

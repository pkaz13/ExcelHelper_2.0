using System.IO;

namespace ExcelHelper_2._0.Wrappers
{
    public interface IFileStreamWrapper
    {
        FileStream Init(string path, FileMode fileMode);
    }
}

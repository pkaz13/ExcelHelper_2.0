using System.IO;

namespace ExcelHelper_2.Utils
{
    class FileHelper
    {
        public static FileInfo DeleteFileIfExistAndCreateNewFile(FileInfo newFile)
        {
            if (newFile.Exists)
            {
                newFile.Delete();
                newFile = new FileInfo(newFile.FullName);
            }
            return newFile;
        }

        public static string FixFileNameExcel(string fileName)
        {
            if (Path.HasExtension(fileName))
            {
                string extension = Path.GetExtension(fileName);
                if (extension != ".xlsx")
                {
                    fileName = Path.GetFileNameWithoutExtension(fileName) + ".xlsx";
                }
            }
            else
            {
                fileName += ".xlsx";
            }
            return fileName;
        }
    }
}

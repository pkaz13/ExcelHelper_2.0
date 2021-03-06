﻿using System.IO;

namespace ExcelHelper_2._0.Wrappers
{
    public class FileStreamWrapper : IFileStreamWrapper
    {
        public FileStream Init(string path, FileMode fileMode)
        {
            return new FileStream(path, fileMode);
        }
    }
}

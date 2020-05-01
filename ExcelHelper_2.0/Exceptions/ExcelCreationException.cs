using System;

namespace ExcelHelper_2._0.Exceptions
{
    public class ExcelCreationException : Exception
    {
        public ExcelCreationException() { }

        public ExcelCreationException(string message) : base(message) { }

        public ExcelCreationException(string message, Exception inner) : base(message, inner) { }
    }
}

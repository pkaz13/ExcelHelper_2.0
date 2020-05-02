using System;

namespace ExcelHelper_2._0.Exceptions
{
    public class ExcelConvertionException : Exception
    {
        public ExcelConvertionException() { }

        public ExcelConvertionException(string message) : base(message) { }

        public ExcelConvertionException(string message, Exception inner) : base(message, inner) { }
    }
}

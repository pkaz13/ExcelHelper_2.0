﻿using System;

namespace ExcelHelper_2.Attributes
{
    public class ExcelColumn : Attribute
    {
        public int ColumnIndex { get; set; }

        public ExcelColumn(int column)
        {
            ColumnIndex = column;
        }
    }
}

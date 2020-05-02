using ExcelHelper_2.Attributes;

namespace ExcelHelperTests.Integration
{
    class TestExcelObject
    {
        [ExcelColumn(1)]
        public string Name { get; set; }
        [ExcelColumn(2)]
        public int Value { get; set; }

        public TestExcelObject(string name, int value)
        {
            Name = name;
            Value = value;
        }

        public TestExcelObject()
        {

        }
    }
}

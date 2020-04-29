namespace ExcelHelperTests.Integration
{
    class TestExcelObject
    {
        public string Name { get; set; }
        public int Value { get; set; }

        public TestExcelObject(string name, int value)
        {
            Name = name;
            Value = value;
        }
    }
}

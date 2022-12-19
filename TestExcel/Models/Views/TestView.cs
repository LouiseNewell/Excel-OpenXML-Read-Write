namespace TestExcel.Models.Views
{
    /// <summary>
    /// TestView data row specific to this app
    /// </summary>
    public class TestView
    {
        public bool ABool { get; set; }
        public DateTime? ADateTimeNullable { get; set; }
        public decimal? ADecimalNullable { get; set; }
        public double ADouble { get; set; }
        public int AnInt { get; set; }
        public string ANumberAsText { get; set; } = string.Empty;
        public string AString { get; set; } = string.Empty;
    }
}
namespace TestExcel.Models.ExcelOpenXML
{
    /// <summary>
    /// CellView cell view is not application specific
    /// </summary>
    public class CellView
    {
        /// <summary>
        /// Value dynamic? cell value so can be any data type
        /// </summary>
        public dynamic? Value { get; set; }
        /// <summary>
        /// Style index should match the CellFormat index in stylesheet for app
        /// </summary>
        public uint Style { get; set; }
        /// <summary>
        /// IsNumberAsText flags if number should be stored as text for example if leading zeros are needed
        /// </summary>
        public bool IsNumberAsText { get; set; }
    }
}
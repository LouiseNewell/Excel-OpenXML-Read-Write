using DocumentFormat.OpenXml.Spreadsheet;
using TestExcel.Models.ExcelOpenXML;

namespace TestExcel.InterfacesServices
{
    /// <summary>
    /// IExcelOpenXML core interface to read and write to excel files using Open XML
    /// This interface is not application specific
    /// uses CellView generic cell model so not application specific here
    /// </summary>
    public interface IExcelOpenXML
    {
        bool CreateSpreadsheet(string fnap, string sheetName, Stylesheet stylesheet, Columns columns, List<List<CellView>> listRows);
        List<List<CellView>> ReadSpreadsheet(string fnap);
    }
}
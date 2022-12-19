using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using TestExcel.Models.ExcelOpenXML;

namespace TestExcel.InterfacesServices
{
    /// <summary>
    /// SExcelOpenXML core service to read and write to excel files using Open XML
    /// Uses CellView model so is not application specific
    /// </summary>
    public class SExcelOpenXML : IExcelOpenXML
    {

        #region public in interface
        /// <summary>
        /// CreateSpreadsheet
        /// </summary>
        /// <param name="fnap">file name and path</param>
        /// <param name="sheetName">sheet name</param>
        /// <param name="stylesheet">stylesheet</param>
        /// <param name="columns">columns widths</param>
        /// <param name="listRows">list of rows</param>
        /// <returns>ok or not</returns>
        public bool CreateSpreadsheet(string fnap, string sheetName, Stylesheet stylesheet, Columns columns, List<List<CellView>> listRows)
        {
            bool bOK;
            // create spreadsheet document and use it to
            using SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(fnap, SpreadsheetDocumentType.Workbook);
            // create workbook part with new workbook
            WorkbookPart workbookPart = spreadsheetDocument.AddWorkbookPart();
            workbookPart.Workbook = new Workbook();
            // add workbook styles part to workbook part and add passed stylesheet
            WorkbookStylesPart workbookStylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
            workbookStylesPart.Stylesheet = stylesheet;
            // add worksheet part to workbook part and add new worksheet
            WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet();
            // add columns widths to worksheet
            worksheetPart.Worksheet.Append(columns);
            // set sheet data with headers rows and data rows
            SheetData sheetData = SetSheetData(listRows);
            // add sheet data to worksheet after column widths
            worksheetPart.Worksheet.Append(sheetData);
            // add new sheets to workbook
            Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());
            // add sheet with worksheet part id and name to sheets
            Sheet sheet = new()
            {
                Id = spreadsheetDocument.WorkbookPart?.GetIdOfPart(worksheetPart) ?? string.Empty,
                SheetId = 1,
                Name = sheetName
            };
            sheets.Append(sheet);
            // save WorkbookPart
            workbookPart.Workbook.Save();
            // close SpreadsheetDocument
            spreadsheetDocument.Close();
            bOK = true;
            return bOK;
        }
        /// <summary>
        /// ReadSpreadsheet
        /// </summary>
        /// <param name="fnap">file name and path</param>
        /// <returns>list of rows</returns>
        public List<List<CellView>> ReadSpreadsheet(string fnap)
        {
            List<List<CellView>> listRows = new();
            // open spreadsheet document and use it to read
            using SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(fnap, false);
            // workbook part
            WorkbookPart? workbookPart = spreadsheetDocument.WorkbookPart;
            if (workbookPart is not null)
            {
                // read first sheet data from worksheet part
                listRows = ReadSheetData(workbookPart);
            }
            return listRows;
        }
        #endregion public in interface
        #region private not in interface
        /// <summary>
        /// ReadSheetData read sheet data from workbook part
        /// </summary>
        /// <param name="workbookPart">workbook part</param>
        /// <returns>list of rows</returns>
        private static List<List<CellView>> ReadSheetData(WorkbookPart workbookPart)
        {
            //todo could adapt this to read given sheet rather than first
            List<List<CellView>> listRows = new();
            // read first sheet data from first worksheet part from workbook part
            WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
            SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();
            // get cells for each row
            foreach (Row row in sheetData.Elements<Row>())
            {
                List<CellView> dataRow = new();
                foreach (Cell cell in row.Elements<Cell>())
                {
                    CellView cellView = GetCellView(workbookPart, cell);
                    dataRow.Add(cellView);
                }
                listRows.Add(dataRow);
            }
            return listRows;
        }
        /// <summary>
        /// GetCellView get cell view
        /// converts value to correct data type based on cell data type and style 
        /// but stores as dynamic? so do not need app specific data model here
        /// </summary>
        /// <param name="workbookPart">workbook part</param>
        /// <param name="cell">cell</param>
        /// <returns>cell view</returns>
        private static CellView GetCellView(WorkbookPart workbookPart, Cell cell)
        {
            CellView cellView = new();
            if (cell is not null)
            {
                string cellText = cell.CellValue?.InnerText.ToString() ?? string.Empty;
                // get cell style
                cellView.Style = uint.TryParse(cell?.StyleIndex?.Value.ToString(), out uint sivid) ? sivid : 0;
                // if cell has no data type then its a datetime or number
                if (cell?.DataType is null)
                {
                    // get cell number format from cell format from stylesheet for workbook for this cell style
                    CellFormat? cellFormat = (CellFormat?)workbookPart?.WorkbookStylesPart?.Stylesheet?.CellFormats?.ElementAt((int)cellView.Style) ?? null;
                    uint cellNumberFormat = cellFormat?.NumberFormatId?.Value ?? 0;
                    // Guide to number formats where data type is null see https://github.com/closedxml/closedxml/wiki/NumberFormatId-Lookup-Table
                    // General = 0, Number = 1, Decimal = 2, Percentage = 10, Scientific = 11, Fraction = 12, Accounting = 44, Currency = 164
                    // DateShort = 14, DateLong = 165, Time = 166
                    // Text = 49 useful for number as text normally text should have data type
                    switch (cellNumberFormat)
                    {
                        // dates
                        case 14:
                        case 165:
                        case 166:
                            // dates are held as numbers so convert to number first then datetime
                            cellView.Value = DateTime.FromOADate(double.TryParse(cellText, out double ddt) ? ddt : 0);
                            break;
                        // take anything else as a number
                        default:
                            cellView.Value = decimal.TryParse(cellText, out decimal dec) ? dec : 0;
                            cellView.Value = cellText;
                            break;
                    }
                }
                // if cell has data type then its a boolean or string
                else
                {
                    switch (cell.DataType.Value)
                    {
                        // boolean convert from text boolean to bool
                        case CellValues.Boolean:
                            cellView.Value = (cell?.CellValue?.InnerText) == "true";
                            break;
                        // strings are stored in shared string table so get index first to get value from there
                        case CellValues.SharedString:
                            int id = int.TryParse(cell.CellValue?.InnerText, out int cvid) ? cvid : 0;
                            SharedStringItem? ssi = workbookPart?.SharedStringTablePart?.SharedStringTable?.Elements<SharedStringItem>()?.ElementAt(id) ?? null;
                            cellView.Value = ssi?.InnerText ?? string.Empty;
                            break;
                        // take anything else as is
                        default:
                            cellView.Value = cell?.CellValue?.InnerText;
                            break;
                    }
                }
            }
            return cellView;
        }
        /// <summary>
        /// SetSheetData set sheet data with list of rows
        /// parses dynamic? data to convert to correct value and data type
        /// so do not need app specific data model here
        /// </summary>
        /// <param name="listRows">list of rows</param>
        /// <returns>sheet data</returns>
        private static SheetData SetSheetData(List<List<CellView>> listRows)
        {
            SheetData sheetData = new();
            foreach (List<CellView> rowView in listRows)
            {
                // create row
                Row row = new();
                foreach (CellView cellView in rowView)
                {
                    // create cell with style required
                    Cell cell = new()
                    {
                        StyleIndex = cellView.Style
                    };
                    // set text for cell value so can parse data type
                    string cellText = cellView.Value?.ToString() ?? string.Empty;
                    // if flagged as number as text set data type for string and value as text for cell
                    if (cellView.IsNumberAsText)
                    {
                        cell.DataType = CellValues.String;
                        cell.CellValue = new CellValue(cellText);
                    }
                    // otherwise parse text for cell and set data type and value accordingly
                    else
                    {
                        if (bool.TryParse(cellText, out bool oBool))
                        {
                            cell.DataType = CellValues.Boolean;
                            cell.CellValue = new CellValue(cellView.Value);
                        }
                        else
                        if (DateTime.TryParse(cellText, out DateTime oDateTime))
                        {
                            cell.DataType = CellValues.Date;
                            cell.CellValue = new CellValue(oDateTime);
                        }
                        else
                        if (decimal.TryParse(cellText, out decimal oDecimal))
                        {
                            cell.DataType = CellValues.Number;
                            cell.CellValue = new CellValue(oDecimal);
                        }
                        else
                        if (double.TryParse(cellText, out double oDouble))
                        {
                            cell.DataType = CellValues.Number;
                            cell.CellValue = new CellValue(oDouble);
                        }
                        else
                        {
                            cell.DataType = CellValues.String;
                            cell.CellValue = new CellValue(cellView.Value);
                        }
                    }
                    // add cell to row
                    row.Append(cell);
                }
                // add row to sheet data
                sheetData.Append(row);
            }
            return sheetData;
        }
        #endregion private not in interface
    }
}
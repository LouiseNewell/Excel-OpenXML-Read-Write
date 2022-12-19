using DocumentFormat.OpenXml.Spreadsheet;
using TestExcel.Models.ExcelOpenXML;
using TestExcel.Models.Views;

namespace TestExcel.InterfacesServices
{
    public class SExcel : IExcel
    {
        private readonly IExcelOpenXML _iExcelOpenXML;

        /// <summary>
        /// SExcel core service to read and write to excel files
        /// Uses data models specific to this app
        /// Uses excel Columns and Stylesheet specific to this app
        /// Uses CellView model to call IExcelOpenXML interface methods 
        /// </summary>
        /// <param name="iExcelOpenXML"></param>
        public SExcel(IExcelOpenXML iExcelOpenXML)
        {
            _iExcelOpenXML = iExcelOpenXML;
        }
        #region public in interface

        /// <summary>
        /// Read read spreadsheet
        /// </summary>
        /// <param name="fnap">filename and path</param>
        /// <returns>Tuple of message and list of data rows</returns>
        public Tuple<string, List<TestView>> Read(string fnap)
        {
            Tuple<string, List<TestView>> results;
            try
            {
                // read spreadsheet to get all excel rows using non application specific CellView model
                List<List<CellView>> excelRows = _iExcelOpenXML.ReadSpreadsheet(fnap);
                // get data rows converting to data model specific to app
                results = GetDataRows(excelRows);
            }
            catch (Exception)
            {
                results = new("Spreadsheet could not be read", new List<TestView>());
            }
            return results;
        }
        /// <summary>
        /// Write write spreadsheet
        /// </summary>
        /// <param name="dataRows">list of data rows</param>
        /// <param name="fnap">file name and path</param>
        /// <returns>message</returns>
        public string Write(List<TestView> dataRows, string fnap)
        {
            string message;
            try
            {
                // set sheet name, stylesheet, and column widths specific to app
                string sheetName = $"TEST {DateTime.Now:MMMM} {DateTime.Now.Year}";
                Stylesheet stylesheet = SetStylesheet();
                Columns columns = SetColumns();
                // set excel rows using non application specific CellView model
                List<List<CellView>> excelRows = new();
                // add header rows specific to app
                excelRows.AddRange(SetHeaderRows());
                // add data rows converting from data model specific to app
                excelRows.AddRange(SetDataRows(dataRows));
                // create spreadsheet using non application specific excel rows
                message = _iExcelOpenXML.CreateSpreadsheet(fnap, sheetName, stylesheet, columns, excelRows) ? $"Spreadsheet {fnap} created" : $"Unable to create spreadsheet {fnap}";
            }
            catch (Exception)
            {
                message = "spreadsheet could not be written";
            }
            return message;
        }
        #endregion public in interface
        #region private not in interface
        /// <summary>
        /// GetDataRows get data rows from rows of excel cells
        /// converts from non application specific CellView model to data model specific to app 
        /// </summary>
        /// <param name="excelRows">rows of excel cells</param>
        /// <returns>Tuple of message and data rows</returns>
        private static Tuple<string, List<TestView>> GetDataRows(List<List<CellView>> excelRows)
        {
            string sMess = string.Empty;
            List<TestView> imports = new();
            try
            {
                // here we can check spreadsheet is in the format expected specific to app
                // check spreadsheet in correct format with 3 header rows then data rows with 7 cells
                if (excelRows.Count > 3)
                {
                    List<List<CellView>> excelSkipHeaders = excelRows.Skip(3).ToList();
                    List<List<CellView>> excelDataRows = excelSkipHeaders.Where(w => w.Count == 7).ToList();
                    if (excelDataRows.Any())
                    {
                        // convert rows of excel cells to data model specific to app
                        foreach (List<CellView> row in excelDataRows)
                        {
                            // set details for row
                            TestView dataRow = new()
                            {
                                ABool = bool.TryParse(row[0].Value?.ToString() ?? false, out bool b) ? b : false,
                                ADateTimeNullable = DateTime.TryParse(row[1].Value?.ToString() ?? string.Empty, out DateTime dat) ? dat : null,
                                ADecimalNullable = decimal.TryParse(row[2].Value?.ToString() ?? string.Empty, out decimal dec) ? dec : null,
                                ADouble = double.TryParse(row[3].Value?.ToString() ?? string.Empty, out double dou) ? dou : 0,
                                AnInt = int.TryParse(row[4].Value?.ToString() ?? string.Empty, out int i) ? i : 0,
                                ANumberAsText = row[5].Value?.ToString() ?? string.Empty,
                                AString = row[6].Value?.ToString() ?? string.Empty
                            };
                            imports.Add(dataRow);
                        }
                    }
                    sMess = $"{imports.Count} rows read from spreadsheet";
                }
                else
                {
                    sMess = "Spreadsheet not in correct format";
                }
            }
            catch (Exception)
            {
                sMess = "Read spreadsheet failed";
            }
            return new Tuple<string, List<TestView>>(sMess, imports);
        }
        /// <summary>
        /// SetColumns set columns as column ranges from min to max with width specific to app
        /// </summary>
        /// <returns>Columns</returns>
        private static Columns SetColumns()
        {
            Columns columns = new();
            columns.Append(new Column() { Min = 1, Max = 2, Width = 20, CustomWidth = true });
            columns.Append(new Column() { Min = 3, Max = 4, Width = 15, CustomWidth = true });
            columns.Append(new Column() { Min = 5, Max = 5, Width = 10, CustomWidth = true });
            columns.Append(new Column() { Min = 6, Max = 7, Width = 25, CustomWidth = true });
            return columns;
        }
        /// <summary>
        /// SetDataRows set data rows of excel cells with style for each cell
        /// converts from data model specific to app to non application specific CellView model
        /// </summary>
        /// <param name="dataRows">data rows</param>
        /// <returns>rows of excel cells</returns>
        private static List<List<CellView>> SetDataRows(IEnumerable<TestView> dataRows)
        {
            List<List<CellView>> excelRows = new();
            // convert each data row to row of excel cells with style for each cell according to data type
            foreach (TestView dataRow in dataRows)
            {
                List<CellView> excelRow = new() {
                    new() { Value = dataRow.ABool, Style = 0 },
                    new() { Value = dataRow.ADateTimeNullable, Style = 4 },
                    new() { Value = dataRow.ADecimalNullable, Style = 6 },
                    new() { Value = dataRow.ADouble, Style = 6 },
                    new() { Value = dataRow.AnInt, Style = 5 },
                    new() { Value = dataRow.ANumberAsText, Style = 7, IsNumberAsText = true },
                    new() { Value = dataRow.AString, Style = 0 }
                };
                excelRows.Add(excelRow);
            }
            return excelRows;
        }
        /// <summary>
        /// SetHeaderRows set header rows of excel cells with style for each cell
        /// </summary>
        /// <returns>rows of excel cells</returns>
        private static List<List<CellView>> SetHeaderRows()
        {
            // set header rows of excel cells with style for each cell as required specific to app
            List<List<CellView>> listHeaderRows = new() {
                new() { new() { Value = "Test Excel", Style = 0 } },
                new() { new() { Value = $"Run {DateTime.Now:MMMM} {DateTime.Now.Year}", Style = 2 } },
                new() {
                    new() { Value = "True or False", Style = 3 },
                    new() { Value = "Date", Style = 3 },
                    new() { Value = "Decimal", Style = 3 },
                    new() { Value = "Double", Style = 3 },
                    new() { Value = "Integer", Style = 3 },
                    new() { Value = "Number as Text", Style = 3 },
                    new() { Value = "Text", Style = 3 }
                }
            };
            return listHeaderRows;
        }
        /// <summary>
        /// SetStylesheet set stylesheet for cell format styles specific to app
        /// </summary>
        /// <returns>Stylesheet</returns>
        private static Stylesheet SetStylesheet()
        {
            // set stylesheet with CellFormats to use index of these for style id on cells
            // note order is important here and some lists have reserved items that must be included first
            Stylesheet stylesheet = new(
                new Fonts(
                    new Font(new FontName() { Val = "Calibri" }, new FontSize() { Val = 12D }),// Index 0 default font
                    new Font(new FontName() { Val = "Calibri" }, new FontSize() { Val = 12D }, new Bold()) // Index 1 bold
                )
                { Count = 2 },
                new Fills(
                    new Fill(new PatternFill() { PatternType = PatternValues.None }), // Index 0 reserved default fill
                    new Fill(new PatternFill() { PatternType = PatternValues.Gray125 }), // Index 1 reserved grey fill
                    new Fill(new PatternFill() { PatternType = PatternValues.Solid, ForegroundColor = new() { Rgb = "FFFF99" } }), // Index 2 yellow fill
                    new Fill(new PatternFill() { PatternType = PatternValues.Solid, ForegroundColor = new() { Rgb = "99CCFF" } })  // Index 3 blue fill
                )
                { Count = 4 },
                new Borders(
                    new Border(new LeftBorder(), new RightBorder(), new TopBorder(), new BottomBorder(), new DiagonalBorder())// Index 0 default border
                )
                { Count = 1 },
                new CellFormats(
                    // Use CellFormat index as style id on cells
                    new CellFormat() { BorderId = 0, FontId = 0, ApplyFont = true }, // Index 0 reserved defaults
                    new CellFormat() { BorderId = 0, FontId = 0, ApplyFont = true }, // Index 1 reserved grey fill
                    new CellFormat() { BorderId = 0, FillId = 2, FontId = 1, ApplyFill = true, ApplyFont = true, }, // Index 2 yellow fill bold used in header rows
                    new CellFormat() { BorderId = 0, FillId = 3, FontId = 1, ApplyFill = true, ApplyFont = true },  // Index 3 blue fill bold used in header rows
                    new CellFormat() { BorderId = 0, FillId = 0, FontId = 0, NumberFormatId = 14, FormatId = 0, ApplyNumberFormat = true }, // Index 4 number format for date "d/m/yyyy" used in data rows
                    new CellFormat() { BorderId = 0, FillId = 0, FontId = 0, NumberFormatId = 1, FormatId = 0, ApplyNumberFormat = true }, // Index 5 number format for number used in data rows
                    new CellFormat() { BorderId = 0, FillId = 0, FontId = 0, NumberFormatId = 4, FormatId = 0, ApplyNumberFormat = true }, // Index 6 number format for decimal "#,##0.00" used in data rows
                    new CellFormat() { BorderId = 0, FillId = 0, FontId = 0, NumberFormatId = 49, FormatId = 0, ApplyNumberFormat = true } // Index 7 number format for number as text "@" used in data rows
                )
                { Count = 8 }
                // Guide to NumberFormatId where data type is null see https://github.com/closedxml/closedxml/wiki/NumberFormatId-Lookup-Table
                // General = 0, Number = 1, Decimal = 2, Percentage = 10, Scientific = 11, Fraction = 12, Accounting = 44, Currency = 164
                // DateShort = 14, DateLong = 165, Time = 166
                // Text = 49 useful for number as text normally text should have data type
            );
            return stylesheet;
        }
        #endregion private not in interface
    }
}
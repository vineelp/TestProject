using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;

namespace ExcelHelper.Helpers
{
    public class ExcelHelper
    {
        private readonly ILogger<ExcelHelper> _logger;
        public ExcelHelper(ILogger<ExcelHelper> logger)
        {
            _logger = logger;
        }
        public void ReadExcelFile(string fileName)
        {
            try
            {
                string curDIR = Directory.GetCurrentDirectory();
                //Lets open the existing excel file and read through its content . Open the excel using openxml sdk
                using (SpreadsheetDocument doc = SpreadsheetDocument.Open(Path.Combine(curDIR, fileName), false))
                {
                    //create the object for workbook part  
                    WorkbookPart workbookPart = doc.WorkbookPart;
                    Sheets thesheetcollection = workbookPart.Workbook.GetFirstChild<Sheets>();
                    StringBuilder excelResult = new StringBuilder();

                    //using for each loop to get the sheet from the sheetcollection  
                    foreach (Sheet thesheet in thesheetcollection)
                    {
                        excelResult.AppendLine("Excel Sheet Name : " + thesheet.Name);
                        excelResult.AppendLine("----------------------------------------------- ");
                        //statement to get the worksheet object by using the sheet id  
                        Worksheet theWorksheet = ((WorksheetPart)workbookPart.GetPartById(thesheet.Id)).Worksheet;

                        SheetData thesheetdata = (SheetData)theWorksheet.GetFirstChild<SheetData>();
                        foreach (Row thecurrentrow in thesheetdata)
                        {
                            foreach (Cell thecurrentcell in thecurrentrow)
                            {
                                //statement to take the integer value  
                                string currentcellvalue = string.Empty;
                                if (thecurrentcell.DataType != null)
                                {
                                    if (thecurrentcell.DataType == CellValues.SharedString)
                                    {
                                        int id;
                                        if (Int32.TryParse(thecurrentcell.InnerText, out id))
                                        {
                                            SharedStringItem item = workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(id);
                                            if (item.Text != null)
                                            {
                                                //code to take the string value  
                                                excelResult.Append(item.Text.Text + " ");
                                            }
                                            else if (item.InnerText != null)
                                            {
                                                currentcellvalue = item.InnerText;
                                            }
                                            else if (item.InnerXml != null)
                                            {
                                                currentcellvalue = item.InnerXml;
                                            }
                                        }
                                    }
                                }
                                else
                                {
                                    excelResult.Append(Convert.ToInt16(thecurrentcell.InnerText) + " ");
                                }
                            }
                            excelResult.AppendLine();
                        }
                        excelResult.Append("");
                        _logger.LogInformation(excelResult.ToString());
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "There was a problem processing the request");
            }
        }
        public DataTable ToDataTable<T>(IList<T> data)
        {
            PropertyDescriptorCollection props =
            TypeDescriptor.GetProperties(typeof(T));
            DataTable table = new DataTable();
            for (int i = 0; i < props.Count; i++)
            {
                PropertyDescriptor prop = props[i];
                table.Columns.Add(prop.Name, Nullable.GetUnderlyingType(
            prop.PropertyType) ?? prop.PropertyType);
            }
            object[] values = new object[props.Count];
            foreach (T item in data)
            {
                for (int i = 0; i < values.Length; i++)
                {
                    values[i] = props[i].GetValue(item);
                }
                table.Rows.Add(values);
            }
            return table;
        }
        public void WriteExcelFile<T>(List<T> obj, string fileName, string[] plHeaders=null, string[] dHeaders=null)
        {
            string curDIR = Directory.GetCurrentDirectory();
            // Lets converts our object data to Datatable for a simplified logic.
            DataTable table = ToDataTable<T>(obj);

            using (SpreadsheetDocument document = SpreadsheetDocument.Create(Path.Combine(curDIR, fileName), SpreadsheetDocumentType.Workbook))
            {
                WorkbookPart workbookPart = document.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();

                WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                var sheetData = new SheetData();
                worksheetPart.Worksheet = new Worksheet(sheetData);

                // Ignoring errors for parsing
                IgnoredErrors ignoredErrors = new IgnoredErrors();
                IgnoredError ignoredError = new IgnoredError()
                {
                    SequenceOfReferences = new ListValue<StringValue> { InnerText = "A:Z" },
                    NumberStoredAsText = true
                };
                ignoredErrors.Append(ignoredError);
                worksheetPart.Worksheet.Append(ignoredErrors);

                // Adding style
                WorkbookStylesPart stylePart = workbookPart.AddNewPart<WorkbookStylesPart>();
                stylePart.Stylesheet = GenerateStylesheet();
                stylePart.Stylesheet.Save();

                // Create custom widths for columns
                Columns lstColumns = worksheetPart.Worksheet.GetFirstChild<Columns>();
                Boolean needToInsertColumns = false;
                if (lstColumns == null)
                {
                    lstColumns = new Columns();
                    needToInsertColumns = true;
                }
                // Min = 1, Max = 1 ==> Apply this to column 1 (A)
                // Min = 2, Max = 2 ==> Apply this to column 2 (B)
                // Width = 25 ==> Set the width to 25
                // CustomWidth = true ==> Tell Excel to use the custom width
                lstColumns.Append(new Column() { Min = 1, Max = 1, Width = 35, CustomWidth = true });
                lstColumns.Append(new Column() { Min = 2, Max = 2, Width = 15, CustomWidth = true });
                lstColumns.Append(new Column() { Min = 3, Max = 3, Width = 15, CustomWidth = true });
                lstColumns.Append(new Column() { Min = 4, Max = 4, Width = 20, CustomWidth = true });
                lstColumns.Append(new Column() { Min = 5, Max = 5, Width = 25, CustomWidth = true });
                lstColumns.Append(new Column() { Min = 6, Max = 6, Width = 20, CustomWidth = true });
                lstColumns.Append(new Column() { Min = 7, Max = 7, Width = 30, CustomWidth = true });
                // Only insert the columns if we had to create a new columns element
                if (needToInsertColumns)
                    worksheetPart.Worksheet.InsertAt(lstColumns, 0);

                Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());
                Sheet sheet = new Sheet() { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Sheet1" };

                sheets.Append(sheet);

                Row headerRow = new Row();

                List<String> columns = new List<string>();
                foreach (System.Data.DataColumn column in table.Columns)
                {
                    columns.Add(column.ColumnName);

                    Cell cell = new Cell();
                    cell.DataType = CellValues.String;
                    cell.CellValue = new CellValue(column.ColumnName);
                    cell.StyleIndex = 3;
                    headerRow.AppendChild(cell);
                }               

                //sheetData.AppendChild(headerRow);

                foreach (DataRow dsrow in table.Rows)
                {                   
                    Row newRow = new Row();
                    foreach (String col in columns)
                    {
                        if (col.Equals("ID"))
                            continue;
                        Cell cell = new Cell();
                        cell.DataType = CellValues.String;
                        cell.CellValue = new CellValue(dsrow[col].ToString());
                        cell.StyleIndex = 6;
                        newRow.AppendChild(cell);
                    }

                    sheetData.AppendChild(newRow);
                }

                workbookPart.Workbook.Save();
            }
        }
        private Stylesheet GenerateStylesheet()
        {
            Fonts fonts = new Fonts(
                new Font(// Index 0 - default
                    new FontSize() { Val = 12 }
                    ),
                new Font( // Index 1 - blue
                    new FontSize() { Val = 12 },
                    new Color() { Rgb = "6699FF" }
                ), 
                new Font( // Index 2 - green
                    new FontSize() { Val = 12 },
                    new Color() { Rgb = "3BF744" }
                ),
                new Font( // Index 3 - red
                    new FontSize() { Val = 12 },
                    new Color() { Rgb = "E96B49" }
                ),
                new Font( // Index 4 
                    new FontSize() { Val = 16 },
                    new Bold()
                ));

            Fills fills = new Fills(
                    new Fill(new PatternFill() { PatternType = PatternValues.None }), // Index 0 - default
                    new Fill(new PatternFill() { PatternType = PatternValues.Gray125 }), // Index 1 - default
                    new Fill(new PatternFill(new ForegroundColor { Rgb = new HexBinaryValue() { Value = "6699FF" } })
                    { PatternType = PatternValues.Solid }), // Index 2 - header blue
                    new Fill(new PatternFill(new ForegroundColor { Rgb = new HexBinaryValue() { Value = "57B050" } })
                    { PatternType = PatternValues.Solid }), // Index 3 - header green
                    new Fill(new PatternFill(new ForegroundColor { Rgb = new HexBinaryValue() { Value = "D9D9D9" } })
                    { PatternType = PatternValues.Solid }) // Index 4 - body grey
                );

            Borders borders = new Borders(
                    new Border(), // index 0 default
                    new Border( // index 1 black border
                        new LeftBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                        new RightBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                        new TopBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                        new BottomBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                        new DiagonalBorder())
                );

            CellFormats cellFormats = new CellFormats(
                    new CellFormat(), // default - 0
                    new CellFormat { FontId = 4, FillId = 2, BorderId = 1, ApplyFill = true }, // header blue - 1
                    new CellFormat { FontId = 4, FillId = 3, BorderId = 1, ApplyFill = true }, // header green - 2
                    new CellFormat { FontId = 1, FillId = 0, BorderId = 1, ApplyBorder = true }, // body blue - 3
                    new CellFormat { FontId = 2, FillId = 0, BorderId = 1, ApplyBorder = true }, // body green - 4
                    new CellFormat { FontId = 3, FillId = 0, BorderId = 1, ApplyBorder = true }, // body red - 5
                    new CellFormat { FontId = 0, FillId = 4, BorderId = 1, ApplyBorder = true }, // body grey - 6
                    new CellFormat { FontId = 0, FillId = 0, BorderId = 1, ApplyBorder = true } // body default - 7
                    
                );

            return new Stylesheet(fonts, fills, borders, cellFormats);
        }
    }
}

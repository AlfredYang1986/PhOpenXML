using System;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Linq;
using System.Text.RegularExpressions;

namespace OMT01.phXlsx {
    public class phExcelFacade {
        private static phExcelFacade _instance = null;

        public static phExcelFacade getInstance() {
            if (_instance == null) {
                _instance = new phExcelFacade();
            }
            return _instance;
        }

        private string getOutputDir() {

            var tmp = "D:\\alfredyang\\";
            if (!Directory.Exists(tmp)) {
                Directory.CreateDirectory(tmp);
            }

            return tmp;
        }

        public void CreateNewExcel(string name, phXlsEnum tp = phXlsEnum.XLSX) {
            var filepath = getOutputDir() + name;

            SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(filepath, SpreadsheetDocumentType.Workbook);

            WorkbookPart workbookpart = spreadsheetDocument.AddWorkbookPart();
            workbookpart.Workbook = new Workbook();

            // Font 
            spreadsheetDocument.WorkbookPart.AddNewPart<WorkbookStylesPart>();
            spreadsheetDocument.WorkbookPart.WorkbookStylesPart.Stylesheet = CreateStylesheet();
            phXlsxFormatConf.getInstance().PushCellFormatsToStylesheet(spreadsheetDocument.WorkbookPart.WorkbookStylesPart.Stylesheet);
            spreadsheetDocument.WorkbookPart.WorkbookStylesPart.Stylesheet.Save();

            WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet();

            // Column
            Columns columns = new Columns();
            columns.Append(new Column() { Min = 1, Max = 3, Width = 20, CustomWidth = true });
            columns.Append(new Column() { Min = 4, Max = 4, Width = 30, CustomWidth = true });
            worksheetPart.Worksheet.Append(columns);

            // Data
            worksheetPart.Worksheet.Append(new SheetData());

            Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());
            Sheet sheet = new Sheet() { Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "pharbers" };
            sheets.Append(sheet);

            worksheetPart.Worksheet.Save();
            workbookpart.Workbook.Save();
            spreadsheetDocument.Close();
        }

        private string GetColumnName(string cellName) {
            // Create a regular expression to match the column name portion of the cell name.
            Regex regex = new Regex("[A-Za-z]+");
            Match match = regex.Match(cellName);

            return match.Value;
        }

        private uint GetRowIndex(string cellName) {
            // Create a regular expression to match the row index portion the cell name.
            Regex regex = new Regex(@"\d+");
            Match match = regex.Match(cellName);

            return uint.Parse(match.Value);
        }

        public void MergeCell(string name, string c1, string c2) {
            var filepath = getOutputDir() + name;
            using (SpreadsheetDocument spreadSheet = SpreadsheetDocument.Open(filepath, true)) {
                var iter = spreadSheet.WorkbookPart.WorksheetParts.GetEnumerator();
                iter.MoveNext();
                var workSheetPart = iter.Current;
                var worksheet = workSheetPart.Worksheet;
                SheetData sheetData = workSheetPart.Worksheet.GetFirstChild<SheetData>();

                InsertCellInWorksheet(GetColumnName(c1), GetRowIndex(c2), workSheetPart);
                InsertCellInWorksheet(GetColumnName(c1), GetRowIndex(c2), workSheetPart);

                MergeCells mergeCells;
                if (worksheet.Elements<MergeCells>().Count() > 0) {
                    mergeCells = worksheet.Elements<MergeCells>().First();
                } else {
                    mergeCells = new MergeCells();

                    // Insert a MergeCells object into the specified position.
                    if (worksheet.Elements<CustomSheetView>().Count() > 0) {
                        worksheet.InsertAfter(mergeCells, worksheet.Elements<CustomSheetView>().First());
                    } else if (worksheet.Elements<DataConsolidate>().Count() > 0) {
                        worksheet.InsertAfter(mergeCells, worksheet.Elements<DataConsolidate>().First());
                    } else if (worksheet.Elements<SortState>().Count() > 0) {
                        worksheet.InsertAfter(mergeCells, worksheet.Elements<SortState>().First());
                    } else if (worksheet.Elements<AutoFilter>().Count() > 0) {
                        worksheet.InsertAfter(mergeCells, worksheet.Elements<AutoFilter>().First());
                    } else if (worksheet.Elements<Scenarios>().Count() > 0) {
                        worksheet.InsertAfter(mergeCells, worksheet.Elements<Scenarios>().First());
                    } else if (worksheet.Elements<ProtectedRanges>().Count() > 0) {
                        worksheet.InsertAfter(mergeCells, worksheet.Elements<ProtectedRanges>().First());
                    } else if (worksheet.Elements<SheetProtection>().Count() > 0) {
                        worksheet.InsertAfter(mergeCells, worksheet.Elements<SheetProtection>().First());
                    } else if (worksheet.Elements<SheetCalculationProperties>().Count() > 0) {
                        worksheet.InsertAfter(mergeCells, worksheet.Elements<SheetCalculationProperties>().First());
                    } else {
                        worksheet.InsertAfter(mergeCells, worksheet.Elements<SheetData>().First());
                    }
                }

                // Create the merged cell and append it to the MergeCells collection.
                MergeCell mergeCell = new MergeCell() { Reference = new StringValue(c1 + ":" + c2) };
                mergeCells.Append(mergeCell);

                worksheet.Save();
            }
        }

        public void PushValueInCell(string name, string value, string c) {
            var filepath = getOutputDir() + name;
            using (SpreadsheetDocument spreadSheet = SpreadsheetDocument.Open(filepath, true)) {
                var iter = spreadSheet.WorkbookPart.WorksheetParts.GetEnumerator();
                iter.MoveNext();
                var workSheetPart = iter.Current;
                SheetData sheetData = workSheetPart.Worksheet.GetFirstChild<SheetData>();

                Cell cell = InsertCellInWorksheet(GetColumnName(c), GetRowIndex(c), workSheetPart);
                cell.CellValue = new CellValue(value);
                cell.DataType = new EnumValue<CellValues>(CellValues.String);

                workSheetPart.Worksheet.Save();
            }
        }

        private Cell InsertCellInWorksheet(string columnName, uint rowIndex, WorksheetPart worksheetPart) {
            Worksheet worksheet = worksheetPart.Worksheet;
            SheetData sheetData = worksheet.GetFirstChild<SheetData>();
            string cellReference = columnName + rowIndex;

            // If the worksheet does not contain a row with the specified row index, insert one.
            Row row;
            if (sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).Count() != 0) {
                row = sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).First();
            } else {
                row = new Row() { RowIndex = rowIndex };
                sheetData.Append(row);
            }

             // If there is not a cell with the specified column name, insert one.  
            if (row.Elements<Cell>().Where(c => c.CellReference.Value == columnName + rowIndex).Count() > 0) {
                return row.Elements<Cell>().Where(c => c.CellReference.Value == cellReference).First();
            } else {
                // Cells must be in sequential order according to CellReference. Determine where to insert the new cell.
                Cell refCell = null;
                foreach (Cell cell in row.Elements<Cell>()) {
                    if (cell.CellReference.Value.Length == cellReference.Length) {
                        if (string.Compare(cell.CellReference.Value, cellReference, true) > 0) {
                            refCell = cell;
                            break;
                        }
                    }
                }

                Cell newCell = new Cell() { CellReference = cellReference };
                row.InsertBefore(newCell, refCell);

                worksheet.Save();
                return newCell;
            }
        }

        private Stylesheet CreateStylesheet() {
            var ss = new Stylesheet();

            var fts = new Fonts();
            var ftn = new FontName { Val = "Arial" };
            var ftsz = new FontSize { Val = 11 };
            var ft = new DocumentFormat.OpenXml.Spreadsheet.Font { FontName = ftn, FontSize = ftsz };
            fts.Append(ft);
            fts.Count = (uint)fts.ChildElements.Count;

            var fills = new Fills();
            var fill = new Fill();
            var patternFill = new PatternFill { PatternType = PatternValues.None };
            fill.PatternFill = patternFill;
            fills.Append(fill);

            fill = new Fill();
            patternFill = new PatternFill { PatternType = PatternValues.Gray125 };
            fill.PatternFill = patternFill;
            fills.Append(fill);

            fills.Count = (uint)fills.ChildElements.Count;

            var borders = new Borders();
            var border = new Border {
                LeftBorder = new LeftBorder(),
                RightBorder = new RightBorder(),
                TopBorder = new TopBorder(),
                BottomBorder = new BottomBorder(),
                DiagonalBorder = new DiagonalBorder()
            };
            borders.Append(border);
            borders.Count = (uint)borders.ChildElements.Count;

            var csfs = new CellStyleFormats();
            var cf = new CellFormat { NumberFormatId = 0, FontId = 0, FillId = 0, BorderId = 0 };
            csfs.Append(cf);
            csfs.Count = (uint)csfs.ChildElements.Count;

            // dd/mm/yyyy is also Excel style index 14

            uint iExcelIndex = 164;
            var nfs = new NumberingFormats();
            var cfs = new CellFormats();

            cf = new CellFormat { NumberFormatId = 0, FontId = 0, FillId = 0, BorderId = 0, FormatId = 0 };
            cfs.Append(cf);

            var nf = new NumberingFormat { NumberFormatId = iExcelIndex, FormatCode = "dd/mm/yyyy hh:mm:ss" };
            nfs.Append(nf);

            cf = new CellFormat {
                NumberFormatId = nf.NumberFormatId,
                FontId = 0,
                FillId = 0,
                BorderId = 0,
                FormatId = 0,
                ApplyNumberFormat = true
            };
            cfs.Append(cf);

            iExcelIndex = 165;
            nfs = new NumberingFormats();
            cfs = new CellFormats();

            cf = new CellFormat { NumberFormatId = 0, FontId = 0, FillId = 0, BorderId = 0, FormatId = 0 };
            cfs.Append(cf);

            nf = new NumberingFormat { NumberFormatId = iExcelIndex, FormatCode = "MMM yyyy" };
            nfs.Append(nf);

            cf = new CellFormat {
                NumberFormatId = nf.NumberFormatId,
                FontId = 0,
                FillId = 0,
                BorderId = 0,
                FormatId = 0,
                ApplyNumberFormat = true
            };
            cfs.Append(cf);

            iExcelIndex = 170;
            nf = new NumberingFormat { NumberFormatId = iExcelIndex, FormatCode = "#,##0.0000" };
            nfs.Append(nf);
            cf = new CellFormat {
                NumberFormatId = nf.NumberFormatId,
                FontId = 0,
                FillId = 0,
                BorderId = 0,
                FormatId = 0,
                ApplyNumberFormat = true
            };
            cfs.Append(cf);

            // #,##0.00 is also Excel style index 4
            iExcelIndex = 171;
            nf = new NumberingFormat { NumberFormatId = iExcelIndex, FormatCode = "#,##0.00" };
            nfs.Append(nf);
            cf = new CellFormat {
                NumberFormatId = nf.NumberFormatId,
                FontId = 0,
                FillId = 0,
                BorderId = 0,
                FormatId = 0,
                ApplyNumberFormat = true
            };
            cfs.Append(cf);

            // @ is also Excel style index 49
            iExcelIndex = 172;
            nf = new NumberingFormat { NumberFormatId = iExcelIndex, FormatCode = "@" };
            nfs.Append(nf);
            cf = new CellFormat {
                NumberFormatId = nf.NumberFormatId,
                FontId = 0,
                FillId = 0,
                BorderId = 0,
                FormatId = 0,
                ApplyNumberFormat = true
            };
            cfs.Append(cf);

            nfs.Count = (uint)nfs.ChildElements.Count;
            cfs.Count = (uint)cfs.ChildElements.Count;

            ss.Append(nfs);
            ss.Append(fts);
            ss.Append(fills);
            ss.Append(borders);
            ss.Append(csfs);
            ss.Append(cfs);

            var css = new CellStyles();
            var cs = new CellStyle { Name = "Normal", FormatId = 0, BuiltinId = 0 };
            css.Append(cs);
            css.Count = (uint)css.ChildElements.Count;
            ss.Append(css);

            var dfs = new DifferentialFormats { Count = 0 };
            ss.Append(dfs);

            var tss = new TableStyles {
                Count = 0,
                DefaultTableStyle = "TableStyleMedium9",
                DefaultPivotStyle = "PivotStyleLight16"
            };
            ss.Append(tss);

            return ss;
        }

        public void AddFont(string name, string font_family, int font_size, string color, bool isBold, string c) {
            var filepath = getOutputDir() + name;
            using (SpreadsheetDocument spreadSheet = SpreadsheetDocument.Open(filepath, true)) {
                var iter = spreadSheet.WorkbookPart.WorksheetParts.GetEnumerator();
                iter.MoveNext();
                var workSheetPart = iter.Current;
                Cell cell = InsertCellInWorksheet(GetColumnName(c), GetRowIndex(c), workSheetPart);

                var fs = spreadSheet.WorkbookPart.WorkbookStylesPart.Stylesheet.Fonts;
                var cf = spreadSheet.WorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats;

                Font font2 = new Font();
                Bold bold1 = new Bold();
                FontSize fontSize2 = new FontSize() { Val = (Double)font_size };
                Color color2 = new Color() { Rgb = new HexBinaryValue(color) };
                FontName fontName2 = new FontName() { Val = font_family };
                FontFamilyNumbering fontFamilyNumbering2 = new FontFamilyNumbering() { Val = 2 };
                FontScheme fontScheme2 = new FontScheme() { Val = FontSchemeValues.Minor };

                if (isBold) {
                    font2.Append(bold1);
                }
                font2.Append(fontSize2);
                font2.Append(color2);
                font2.Append(fontName2);
                font2.Append(fontFamilyNumbering2);
                font2.Append(fontScheme2);

                fs.Append(font2);

                CellFormat cellFormat2 = new CellFormat() { NumberFormatId = 0, FontId = (UInt32)(fs.Elements<Font>().Count() - 1), FillId = 0, BorderId = 0, FormatId = 0, ApplyFill = true };
                cf.Append(cellFormat2);
                spreadSheet.WorkbookPart.WorkbookStylesPart.Stylesheet.Save();

                cell.StyleIndex = (UInt32)(spreadSheet.WorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats.Elements<CellFormat>().Count() - 1);
                workSheetPart.Worksheet.Save();
            }
        }

        public void AddFill(string name, string color, string c) {
            var filepath = getOutputDir() + name;
            using (SpreadsheetDocument spreadSheet = SpreadsheetDocument.Open(filepath, true)) {
                var iter = spreadSheet.WorkbookPart.WorksheetParts.GetEnumerator();
                iter.MoveNext();
                var workSheetPart = iter.Current;
                Cell cell = InsertCellInWorksheet(GetColumnName(c), GetRowIndex(c), workSheetPart);

                Fills fs = spreadSheet.WorkbookPart.WorkbookStylesPart.Stylesheet.Fills;
                var cf = spreadSheet.WorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats;

                Fill fill1 = new Fill();

                PatternFill patternFill1 = new PatternFill() { PatternType = PatternValues.Solid };
                ForegroundColor foregroundColor1 = new ForegroundColor() { Rgb = color };
                BackgroundColor backgroundColor1 = new BackgroundColor() { Indexed = (UInt32Value)64U };

                patternFill1.Append(foregroundColor1);
                patternFill1.Append(backgroundColor1);

                fill1.Append(patternFill1);
                fs.Append(fill1);

                CellFormat cellFormat2 = new CellFormat() { NumberFormatId = 0, FontId = 0, FillId = (UInt32)(fs.Elements<Fill>().Count() - 1), BorderId = 0, FormatId = 0, ApplyFill = true };
                cf.Append(cellFormat2);
                spreadSheet.WorkbookPart.WorkbookStylesPart.Stylesheet.Save();

                cell.StyleIndex = (UInt32)(spreadSheet.WorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats.Elements<CellFormat>().Count() - 1);
                workSheetPart.Worksheet.Save();
            }
        }

        public void AddAlignment(string name, HorizontalAlignmentValues hv, VerticalAlignmentValues vv, string c) {
            var filepath = getOutputDir() + name;
            using (SpreadsheetDocument spreadSheet = SpreadsheetDocument.Open(filepath, true)) {
                var iter = spreadSheet.WorkbookPart.WorksheetParts.GetEnumerator();
                iter.MoveNext();
                var workSheetPart = iter.Current;
                Cell cell = InsertCellInWorksheet(GetColumnName(c), GetRowIndex(c), workSheetPart);

                Fills fs = spreadSheet.WorkbookPart.WorkbookStylesPart.Stylesheet.Fills;
                var cf = spreadSheet.WorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats;

                CellFormat cellFormat2 = new CellFormat() {
                    NumberFormatId = 0, FontId = 0,
                    FillId = (UInt32)(fs.Elements<Fill>().Count() - 1),
                    BorderId = 0,
                    FormatId = 0,
                    Alignment = new Alignment() { Horizontal = hv, Vertical = vv },
                    ApplyFill = true
                };
                cf.Append(cellFormat2);
                spreadSheet.WorkbookPart.WorkbookStylesPart.Stylesheet.Save();

                cell.StyleIndex = (UInt32)(spreadSheet.WorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats.Elements<CellFormat>().Count() - 1);
                workSheetPart.Worksheet.Save();
            }
        }
    }
}

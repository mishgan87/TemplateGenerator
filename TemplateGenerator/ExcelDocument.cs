using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TemplateGenerator
{
    class ExcelDocument
    {
        string FileName { get; set; }
        ExcelDocument()
        {
            FileName = "";
        }
        ExcelDocument(string filename)
        {
            Load(filename);
        }
        public void Load(string filename)
        {
            FileName = filename;
        }
        private static string GetExcelColumnName(int number)
        {
            string name = "";
            while (number > 0)
            {
                int mod = (number - 1) % 26;
                name = Convert.ToChar('A' + mod) + name;
                number = (number - mod) / 26;
            }
            return name;
        }
        static Cell AddCell(Row row, int columnIndex, string text, uint styleIndex)
        {
            Cell refCell = null;
            Cell newCell = new Cell()
            {
                CellReference = $"{GetExcelColumnName(columnIndex + 1)}{row.RowIndex}",
                StyleIndex = styleIndex
            };
            row.InsertBefore(newCell, refCell);

            newCell.CellValue = new CellValue(text);
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);

            return newCell;
        }
        public static void ImportDataTable(DataTable table, string filename)
        {
            using (SpreadsheetDocument document = SpreadsheetDocument.Create(filename, SpreadsheetDocumentType.Workbook))
            {

                WorkbookPart workbookpart = document.AddWorkbookPart();
                WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
                workbookpart.Workbook = new Workbook();
                worksheetPart.Worksheet = new Worksheet(new SheetData());
                Sheets sheets = document.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());
                Sheet sheet = new Sheet()
                {
                    Id = document.WorkbookPart.GetIdOfPart(worksheetPart),
                    SheetId = 1,
                    Name = $"Sheet1"
                };
                sheets.Append(sheet);

                WorkbookStylesPart wbsp = workbookpart.AddNewPart<WorkbookStylesPart>();
                wbsp.Stylesheet = GenerateStyleSheet();
                wbsp.Stylesheet.Save();

                SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

                Columns lstColumns = worksheetPart.Worksheet.GetFirstChild<Columns>();
                if (lstColumns == null)
                {
                    lstColumns = new Columns();
                }

                for (int col = 0; col < table.Columns.Count; col++)
                {
                    lstColumns.Append(new Column() { Min = 1, Max = 15, Width = 15, CustomWidth = true, BestFit = true });
                }
                worksheetPart.Worksheet.InsertAt(lstColumns, 0);

                MergeCells mergeCells = new MergeCells(); // Массив объединённых ячеек
                List<Cell> cells = new List<Cell>();
                Cell cell = null;

                bool isHorizontalMergeStarted = false;
                bool isVerticalMergeStarted = false;

                int verticalMergeColumn = -1;

                uint rowIndex = 1;
                foreach(object tableRow in table.Rows)
                {
                    Row row = new Row { RowIndex = rowIndex, Height = 21.25, CustomHeight = true };
                    sheetData.Append(row);
                    DataRow dataRow = (DataRow)tableRow;
                    int col = 0;

                    foreach (object dataRowCell in dataRow.ItemArray)
                    {
                        string text = dataRowCell.ToString();
                        cell = AddCell(row, col, text, 1);

                        if (col == verticalMergeColumn && isVerticalMergeStarted)
                        {
                            cells.Add(cell);
                        }

                        if (text.Contains("[HM]"))
                        {
                            if (!isHorizontalMergeStarted)
                            {
                                cells.Add(cell);
                                isHorizontalMergeStarted = true;
                            }
                            else
                            {
                                cells.Add(cell);
                                isHorizontalMergeStarted = false;
                                mergeCells.Append(new MergeCell() { Reference = new StringValue($"{cells.First().CellReference.Value}:{cells.Last().CellReference.Value}") });
                                cells.Clear();
                            }
                        }

                        if (col == dataRow.ItemArray.Count() && isHorizontalMergeStarted)
                        {
                            cells.Add(cell);
                            isHorizontalMergeStarted = false;
                            mergeCells.Append(new MergeCell() { Reference = new StringValue($"{cells.First().CellReference.Value}:{cells.Last().CellReference.Value}") });
                            cells.Clear();
                        }

                        if (text.Contains("[VM]"))
                        {
                            if (!isVerticalMergeStarted)
                            {
                                cells.Add(cell);
                                isVerticalMergeStarted = true;
                                verticalMergeColumn = col;
                            }
                            else
                            {
                                cells.Remove(cells.Last());
                                verticalMergeColumn = -1;
                                isVerticalMergeStarted = false;
                                mergeCells.Append(new MergeCell() { Reference = new StringValue($"{cells.First().CellReference.Value}:{cells.Last().CellReference.Value}") });
                                cells.Clear();

                                cells.Add(cell);
                                isVerticalMergeStarted = true;
                                verticalMergeColumn = col;
                            }
                        }

                        if (rowIndex == table.Rows.Count && isVerticalMergeStarted)
                        {
                            verticalMergeColumn = -1;
                            isVerticalMergeStarted = false;
                            mergeCells.Append(new MergeCell() { Reference = new StringValue($"{cells.First().CellReference.Value}:{cells.Last().CellReference.Value}") });
                            cells.Clear();
                        }

                        col++;
                    }

                    rowIndex++;
                }

                worksheetPart.Worksheet.InsertAfter(mergeCells, worksheetPart.Worksheet.Elements<SheetData>().First()); // Добавляем к документу список объединённых ячеек
                worksheetPart.Worksheet.Save();

                workbookpart.Workbook.Save();
                document.Close();
            }

            MessageBox.Show("Done");
        }
        public static void ImportTreeView(TreeView treeView)
        {
            string fileName = $"{treeView.Nodes[0].Text}.xlsx";

            using (SpreadsheetDocument document = SpreadsheetDocument.Create(fileName, SpreadsheetDocumentType.Workbook))
            {

                WorkbookPart workbookpart = document.AddWorkbookPart();
                WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
                workbookpart.Workbook = new Workbook();
                worksheetPart.Worksheet = new Worksheet(new SheetData());
                Sheets sheets = document.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());
                Sheet sheet = new Sheet()
                {
                    Id = document.WorkbookPart.GetIdOfPart(worksheetPart),
                    SheetId = 1,
                    Name = $"Sheet1"
                };
                sheets.Append(sheet);

                WorkbookStylesPart wbsp = workbookpart.AddNewPart<WorkbookStylesPart>();
                wbsp.Stylesheet = GenerateStyleSheet();
                wbsp.Stylesheet.Save();

                SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

                TreeNodeCollection nodes = treeView.Nodes[0].Nodes;

                int max = 0;
                int maxCount = 0;
                foreach (TreeNode group in nodes)
                {
                    foreach (TreeNode attribute in group.Nodes)
                    {
                        int count = attribute.GetNodeCount(false);
                        if (count > maxCount)
                        {
                            maxCount = count;
                        }
                        if (count * 2 > max)
                        {
                            max = count * 2;
                        }
                    }
                }

                max += 2;

                Columns lstColumns = worksheetPart.Worksheet.GetFirstChild<Columns>();
                if (lstColumns == null)
                {
                    lstColumns = new Columns();
                }

                for (int col = 0; col <= max; col++)
                {
                    DoubleValue width = 15;
                    if (col == 0)
                    {
                        width = 5;
                    }
                    lstColumns.Append(new Column() { Min = 1, Max = (UInt32)width, Width = width, CustomWidth = true, BestFit = true });
                }
                worksheetPart.Worksheet.InsertAt(lstColumns, 0);

                uint rowIndex = 1;
                int attributeIndex = 1;

                string mergeRange = "";
                MergeCells mergeCells = new MergeCells(); // Массив объединённых ячеек
                Cell cell = null;

                foreach (TreeNode group in nodes)
                {
                    rowIndex = (uint)(sheetData.Descendants<Row>().Count() + 1);
                    Row rowGroup = new Row { RowIndex = rowIndex, Height = 21.25, CustomHeight = true };
                    sheetData.Append(rowGroup);

                    cell = AddCell(rowGroup, 0, group.Text, 1);
                    mergeRange = cell.CellReference.Value;
                    for (int index = 1; index < max; index++)
                    {
                        cell = AddCell(rowGroup, index, "", 1);
                    }
                    mergeRange += $":{cell.CellReference.Value}";
                    mergeCells.Append(new MergeCell() { Reference = new StringValue(mergeRange) });

                    foreach (TreeNode attribute in group.Nodes)
                    {
                        rowIndex = (uint)(sheetData.Descendants<Row>().Count() + 1);
                        Row row = new Row { RowIndex = rowIndex, Height = 21.25, CustomHeight = true };
                        sheetData.Append(row);

                        AddCell(row, 0, $"{attributeIndex}", 1);
                        cell = AddCell(row, 1, attribute.Text, 1);
                        mergeRange = cell.CellReference.Value;

                        int count = attribute.GetNodeCount(false);
                        int nodeIndex = 0;

                        for (int index = 2; index < max; index++)
                        {
                            string text = "";

                            if (nodeIndex < count && count > 0)
                            {
                                text = attribute.Nodes[nodeIndex].Text;
                            }

                            nodeIndex++;

                            cell = AddCell(row, index, text, 2);

                            if (index == maxCount + 1 && count == 0)
                            {
                                mergeRange += $":{cell.CellReference.Value}";
                                mergeCells.Append(new MergeCell() { Reference = new StringValue(mergeRange) });
                            }
                        }
                        attributeIndex++;
                    }
                }

                worksheetPart.Worksheet.InsertAfter(mergeCells, worksheetPart.Worksheet.Elements<SheetData>().First()); // Добавляем к документу список объединённых ячеек
                worksheetPart.Worksheet.Save();
                workbookpart.Workbook.Save();
                document.Close();
            }

            MessageBox.Show("Done");
        }
        static Stylesheet GenerateStyleSheet()
        {
            return new Stylesheet(

                // Шрифты

                new Fonts(

                    // 0 - Arial - 8 - Чёрный
                    new Font(
                        // new Bold(),
                        new FontSize() { Val = 8 },
                        new Color() { Rgb = new HexBinaryValue() { Value = "000000" } },
                        new FontName() { Val = "Arial" })
                ),

                // Заполнение цветом

                new Fills(

                    // 0 - Без заполнения
                    new Fill(
                        new PatternFill() { PatternType = PatternValues.None }
                    )

                ),

                // Границы ячейки
                new Borders(

                    // 0 - Граней нет
                    new Border(
                        new LeftBorder(),
                        new RightBorder(),
                        new TopBorder(),
                        new BottomBorder(),
                        new DiagonalBorder()),

                    // 1 - Грани все тонкие чёрные
                    new Border(
                        new LeftBorder(
                            new Color() { Rgb = new HexBinaryValue() { Value = "000000" } }
                        )
                        { Style = BorderStyleValues.Thin },
                        new RightBorder(
                            new Color() { Rgb = new HexBinaryValue() { Value = "000000" } }
                        )
                        { Style = BorderStyleValues.Thin },
                        new TopBorder(
                            new Color() { Rgb = new HexBinaryValue() { Value = "000000" } }
                        )
                        { Style = BorderStyleValues.Thin },
                        new BottomBorder(
                            new Color() { Rgb = new HexBinaryValue() { Value = "000000" } }
                        )
                        { Style = BorderStyleValues.Thin },
                        new DiagonalBorder()
                    )

                ),

                // Формат ячейки
                new CellFormats(

                    // 0 - The default cell style
                    new CellFormat() { FontId = 0, FillId = 0, BorderId = 0 },

                    // 1 - Left alignment
                    new CellFormat(
                        new Alignment() { Horizontal = HorizontalAlignmentValues.Left, Vertical = VerticalAlignmentValues.Center, WrapText = true }
                    )
                    { FontId = 0, FillId = 0, BorderId = 1, ApplyAlignment = true, ApplyFill = true, ApplyFont = true },

                    // 2 - Center alignment
                    new CellFormat(
                        new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center, WrapText = true }
                    )
                    { FontId = 0, FillId = 0, BorderId = 1, ApplyAlignment = true, ApplyFill = true, ApplyFont = true },

                    // 3 - Vertival Center alignment
                    new CellFormat(
                        new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center, TextRotation = 90, WrapText = true }
                    )
                    { FontId = 0, FillId = 0, BorderId = 1, ApplyAlignment = true, ApplyFill = true, ApplyFont = true }

                )
            );
        }
        private Columns AutoSize(SheetData sheetData)
        {
            var maxColWidth = GetMaxCharacterWidth(sheetData);

            Columns columns = new Columns();
            //this is the width of my font - yours may be different
            double maxWidth = 7;
            foreach (var item in maxColWidth)
            {
                //width = Truncate([{Number of Characters} * {Maximum Digit Width} + {5 pixel padding}]/{Maximum Digit Width}*256)/256
                double width = Math.Truncate((item.Value * maxWidth + 5) / maxWidth * 256) / 256;

                //pixels=Truncate(((256 * {width} + Truncate(128/{Maximum Digit Width}))/256)*{Maximum Digit Width})
                double pixels = Math.Truncate(((256 * width + Math.Truncate(128 / maxWidth)) / 256) * maxWidth);

                //character width=Truncate(({pixels}-5)/{Maximum Digit Width} * 100+0.5)/100
                double charWidth = Math.Truncate((pixels - 5) / maxWidth * 100 + 0.5) / 100;

                Column col = new Column() { BestFit = true, Min = (UInt32)(item.Key + 1), Max = (UInt32)(item.Key + 1), CustomWidth = true, Width = (DoubleValue)width };
                columns.Append(col);
            }

            return columns;
        }
        private Dictionary<int, int> GetMaxCharacterWidth(SheetData sheetData)
        {
            //iterate over all cells getting a max char value for each column
            Dictionary<int, int> maxColWidth = new Dictionary<int, int>();
            var rows = sheetData.Elements<Row>();
            UInt32[] numberStyles = new UInt32[] { 5, 6, 7, 8 }; //styles that will add extra chars
            UInt32[] boldStyles = new UInt32[] { 1, 2, 3, 4, 6, 7, 8 }; //styles that will bold
            foreach (var r in rows)
            {
                var cells = r.Elements<Cell>().ToArray();

                //using cell index as my column
                for (int i = 0; i < cells.Length; i++)
                {
                    var cell = cells[i];
                    var cellValue = cell.CellValue == null ? string.Empty : cell.CellValue.InnerText;
                    var cellTextLength = cellValue.Length;

                    if (cell.StyleIndex != null && numberStyles.Contains(cell.StyleIndex))
                    {
                        int thousandCount = (int)Math.Truncate((double)cellTextLength / 4);

                        //add 3 for '.00' 
                        cellTextLength += (3 + thousandCount);
                    }

                    if (cell.StyleIndex != null && boldStyles.Contains(cell.StyleIndex))
                    {
                        //add an extra char for bold - not 100% acurate but good enough for what i need.
                        cellTextLength += 1;
                    }

                    if (maxColWidth.ContainsKey(i))
                    {
                        var current = maxColWidth[i];
                        if (cellTextLength > current)
                        {
                            maxColWidth[i] = cellTextLength;
                        }
                    }
                    else
                    {
                        maxColWidth.Add(i, cellTextLength);
                    }
                }
            }

            return maxColWidth;
        }
    }
}

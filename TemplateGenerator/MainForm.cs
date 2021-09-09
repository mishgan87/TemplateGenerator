using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Color = DocumentFormat.OpenXml.Spreadsheet.Color;
using Font = DocumentFormat.OpenXml.Spreadsheet.Font;

namespace TemplateGenerator
{
    public partial class MainForm : System.Windows.Forms.Form
    {
        private TreeNode ClipboardNode { get; set; }
        public MainForm()
        {
            InitializeComponent();

            ClipboardNode = null;
        }
        private void NewTemplate(object sender, EventArgs e)
        {
            treeView.Nodes.Clear();
            treeView.Nodes.Add(new TreeNode("New Template"));
            treeView.SelectedNode = treeView.Nodes[0];

            ClipboardNode = null;
        }
        private void AddNode(object sender, EventArgs e)
        {
            if (treeView.SelectedNode != null && treeView.SelectedNode.Level < 3)
            {
                string name = "";
                if (treeView.SelectedNode.Level == 0)
                {
                    name = "New Group";
                }
                if (treeView.SelectedNode.Level == 1)
                {
                    name = "New Attribute";
                }
                if (treeView.SelectedNode.Level == 2)
                {
                    name = "New Parameter";
                }
                TreeNode node = new TreeNode(name);
                treeView.SelectedNode.Nodes.Add(node);
                treeView.SelectedNode = node;
                treeView.LabelEdit = true;
                if (!node.IsEditing)
                {
                    node.BeginEdit();
                }
            }
        }
        private void RemoveNode(object sender, EventArgs e)
        {
            TreeNode node = treeView.SelectedNode;
            treeView.Nodes.Remove(node);
        }
        private void EditNode(object sender, EventArgs e)
        {
            TreeNode node = treeView.SelectedNode;
            if (node != null)
            {
                treeView.SelectedNode = node;
                treeView.LabelEdit = true;
                if (!node.IsEditing)
                {
                    node.BeginEdit();
                }
            }
        }
        private void CopyNode(object sender, EventArgs e)
        {
            ClipboardNode = treeView.SelectedNode;
        }
        private void PasteNode(object sender, EventArgs e)
        {
            if (treeView.SelectedNode != null && treeView.SelectedNode.Level < 3)
            {
                treeView.SelectedNode.Nodes.Add(new TreeNode(ClipboardNode.Text));
            }
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
        private void ButtonGenerate_Click(object sender, EventArgs e)
        {
            string fileName = $"{treeView.Nodes[0].Text}.xlsx";

            // Create a spreadsheet document by using the file name
            using (SpreadsheetDocument document = SpreadsheetDocument.Create(fileName, SpreadsheetDocumentType.Workbook))
            {
                // Add a WorkbookPart and Workbook objects
                WorkbookPart workbookpart = document.AddWorkbookPart();
                workbookpart.Workbook = new Workbook();

                // Add a WorksheetPart
                WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();

                // Create Worksheet and SheetData objects
                worksheetPart.Worksheet = new Worksheet(new SheetData());

                // Add a Sheets object
                Sheets sheets = document.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());

                // Append the new worksheet named "Permissible Grid" and associate it with the workbook
                Sheet sheet = new Sheet()
                {
                    Id = document.WorkbookPart.GetIdOfPart(worksheetPart),
                    SheetId = 1,
                    Name = $"Sheet1"
                };
                sheets.Append(sheet);

                // Append stylesheets
                WorkbookStylesPart wbsp = workbookpart.AddNewPart<WorkbookStylesPart>();
                wbsp.Stylesheet = GenerateStyleSheet();
                wbsp.Stylesheet.Save();

                SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>(); // Get the sheetData cell table

                TreeNodeCollection nodes = treeView.Nodes[0].Nodes;

                int max = 0;
                foreach (TreeNode group in nodes)
                {
                    foreach (TreeNode attribute in group.Nodes)
                    {
                        int count = attribute.GetNodeCount(false);
                        if (max < count * 2)
                        {
                            max = count * 2;
                        }
                    }
                }

                max += 2;

                uint rowIndex = 1;
                int attributeIndex = 1;

                string mergeRange = "";
                MergeCells mergeCells = new MergeCells(); // Массив объединённых ячеек
                Cell cell = null;

                foreach (TreeNode group in nodes)
                {
                    rowIndex = (uint)(sheetData.Descendants<Row>().Count() + 1);
                    Row rowGroup = new Row { RowIndex = rowIndex };
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
                        Row row = new Row { RowIndex = rowIndex };
                        sheetData.Append(row);

                        AddCell(row, 0, $"{attributeIndex}", 1);
                        cell = AddCell(row, 1, attribute.Text, 1);
                        mergeRange = cell.CellReference.Value;

                        int count = attribute.GetNodeCount(false);
                        for (int index = 0; index < count * 2; index++)
                        {
                            string text = "";
                            if (index < count)
                            {
                                text = attribute.Nodes[index].Text;
                            }
                            cell = AddCell(row, index + 2, text, 1);
                        }

                        for (int index = count * 2; index < max; index++)
                        {
                            cell = AddCell(row, index, "", 1);
                        }

                        if (count == 0)
                        {
                            mergeRange += $":{cell.CellReference.Value}";
                //            mergeCells.Append(new MergeCell() { Reference = new StringValue(mergeRange) });
                        }

                        attributeIndex++;
                    }
                }

                // Добавляем к документу список объединённых ячеек

                worksheetPart.Worksheet.InsertAfter(mergeCells, worksheetPart.Worksheet.Elements<SheetData>().First());
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
                    ),

                    // 1 - Серый
                    new Fill(
                        new PatternFill(new ForegroundColor() { Rgb = new HexBinaryValue() { Value = "FFAAAAAA" } })
                        { PatternType = PatternValues.Solid }
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

                    // 1 - Грани все
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
                        new Alignment() { Horizontal = HorizontalAlignmentValues.Left, Vertical = VerticalAlignmentValues.Center }
                    )
                    { FontId = 0, FillId = 0, BorderId = 1, ApplyAlignment = true, ApplyFill = true, ApplyFont = true },

                    // 2 - Center alignment
                    new CellFormat(
                        new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center }
                    )
                    { FontId = 0, FillId = 0, BorderId = 1, ApplyAlignment = true, ApplyFill = true, ApplyFont = true },

                    // 3 - Vertival text
                    new CellFormat(
                        new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center, TextRotation = 90 }
                    )
                    { FontId = 0, FillId = 0, BorderId = 1, ApplyAlignment = true, ApplyFill = true, ApplyFont = true }
                )
            );
        }
        private void TreeView_MouseClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                treeView.SelectedNode = e.Node;
                
                System.Windows.Forms.ContextMenuStrip menu = new System.Windows.Forms.ContextMenuStrip();
                
                ToolStripItem itemEdit = menu.Items.Add("Edit");
                //itemEdit.Image = Bitmap.FromFile("\\Icons\\add.ico");
                itemEdit.Click += new EventHandler(this.EditNode);
                
                ToolStripItem itemAdd = menu.Items.Add("Add");
                //itemEdit.Image = Bitmap.FromFile("\\Icons\\add.ico");
                itemAdd.Click += new EventHandler(this.AddNode);
                
                ToolStripItem itemRemove = menu.Items.Add("Remove");
                //itemEdit.Image = Bitmap.FromFile("\\Icons\\add.ico");
                itemRemove.Click += new EventHandler(this.RemoveNode);

                menu.Items.Add("");
                menu.Items.Add("");

                ToolStripItem itemCopy = menu.Items.Add("Copy");
                //itemEdit.Image = Bitmap.FromFile("\\Icons\\add.ico");
                itemCopy.Click += new EventHandler(this.CopyNode);

                ToolStripItem itemPaste = menu.Items.Add("Paste");
                //itemEdit.Image = Bitmap.FromFile("\\Icons\\add.ico");
                itemPaste.Click += new EventHandler(this.PasteNode);

                menu.Show((System.Windows.Forms.Control)sender, new Point(e.X, e.Y));
            }
        }
    }
}

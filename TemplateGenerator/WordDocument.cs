using DocumentFormat.OpenXml.InkML;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Table = DocumentFormat.OpenXml.Wordprocessing.Table;

namespace TemplateGenerator
{
    class WordDocument
    {
        WordDocument()
        {

        }
        public static void ToGridView(string filename, DataGridView gridView)
        {
            using (WordprocessingDocument doc = WordprocessingDocument.Open(filename, true))
            {
                List<Table> tables = new List<Table>();
                foreach (var tbl in doc.MainDocumentPart.Document.Body.Elements<Table>())
                {
                    tables.Add(tbl);
                }

                DataTable grid = new DataTable();

                int index = 1;

                foreach (Table table in tables)
                {
                    foreach (TableRow row in table.Elements<TableRow>())
                    {
                        int col = 1;
                        DataRow drow = grid.NewRow();

                        foreach (TableCell cell in row.Elements<TableCell>())
                        {
                            string text = $"{cell.InnerText}";

                            var properties = cell.TableCellProperties;
                            var direction = properties.TextDirection;
                            string textDirection = "[]";
                            if (direction != null)
                            {
                                textDirection = $"[{direction.Val.InnerText}]";
                            }

                            if (properties.HorizontalMerge != null)
                            {
                                if (properties.HorizontalMerge.Val != null)
                                {
                                    text = $"[HM][{textDirection}] {text}";
                                }
                            }

                            if (properties.VerticalMerge != null)
                            {
                                if (properties.VerticalMerge.Val != null)
                                {
                                    text = $"[VM]{textDirection} {text}";
                                }
                            }

                            if (grid.Columns.Count < col)
                            {
                                grid.Columns.Add($"{col}");
                            }
                            drow[col - 1] = text;

                            col++;

                            Paragraph paragraph = cell.Elements<Paragraph>().First();
                            // Run run = paragraph.Elements<Run>().First();
                            // Text txt = run.Elements<Text>().First();
                        }

                        index++;
                        grid.Rows.Add(drow);
                    }
                }

                gridView.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
                gridView.DataSource = grid;
                foreach(DataGridViewRow dgvrow in gridView.Rows)
                {
                    dgvrow.Height = 55;
                }

                string xfilename = filename;
                xfilename = xfilename.Remove(xfilename.LastIndexOf("."), xfilename.Length - xfilename.LastIndexOf("."));
                xfilename += ".xlsx";

                ExcelDocument.ImportDataTable(grid, xfilename);
            }
        }
    }
}

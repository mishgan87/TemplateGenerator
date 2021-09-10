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
        public static void ImportExportToTreeView(string filename, TreeView treeView, DataGridView gridView)
        {
            treeView.Nodes.Clear();

            using (WordprocessingDocument doc = WordprocessingDocument.Open(filename, true))
            {
                List<Table> tables = new List<Table>();
                foreach (var tbl in doc.MainDocumentPart.Document.Body.Elements<Table>())
                {
                    tables.Add(tbl);
                }

                Table table = tables[0];

                DataTable grid = new DataTable();

                int index = 1;
                TreeNode rootNode = new TreeNode("Root");
                foreach (TableRow row in table.Elements<TableRow>())
                {
                    int col = 1;
                    DataRow drow = grid.NewRow();

                    TreeNode groupNode = new TreeNode($"Row {index}");
                    TreeNode attributeNode = new TreeNode($"Row {index}");

                    foreach (TableCell cell in row.Elements<TableCell>())
                    {
                        string text = $"{cell.InnerText}";

                        var properties = cell.TableCellProperties;
                        
                        if (properties.HorizontalMerge != null)
                        {
                            if (properties.HorizontalMerge.Val != null)
                            {
                                text = $"[HM.{properties.HorizontalMerge.Val.Value}] {text}";
                            }
                        }

                        if (properties.VerticalMerge != null)
                        {
                            if (properties.VerticalMerge.Val != null)
                            {
                                text = $"[VM.{properties.VerticalMerge.Val.Value}] {text}";
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
                        TreeNode treeNode = new TreeNode(text);
                        attributeNode.Nodes.Add(treeNode);
                    }
                    groupNode.Nodes.Add(attributeNode);
                    rootNode.Nodes.Add(groupNode);
                    index++;
                    grid.Rows.Add(drow);
                }

                treeView.Nodes.Add(rootNode);

                gridView.DataSource = grid;
            }
        }
    }
}

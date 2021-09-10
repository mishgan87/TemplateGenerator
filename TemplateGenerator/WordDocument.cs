using DocumentFormat.OpenXml.InkML;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
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
        public static void ImportExportToTreeView(string filename, TreeView treeView)
        {
            // Open a WordprocessingDocument for editing using the filepath.
            using (WordprocessingDocument doc = WordprocessingDocument.Open(filename, true))
            {
                List<Table> tables = new List<Table>();
                foreach (var tbl in doc.MainDocumentPart.Document.Body.Elements<Table>())
                {
                    tables.Add(tbl);
                }

                // Берем первую таблицу (конечно, нужно чтобы она была)
                Table table = tables[0];

                // Первая строка из таблицы
                TableRow row = table.Elements<TableRow>().ElementAt(0);

                // Первая ячейка из строки
                TableCell cell = row.Elements<TableCell>().ElementAt(0);

                Paragraph paragraph = cell.Elements<Paragraph>().First();
                Run run = paragraph.Elements<Run>().First();
                Text txt = run.Elements<Text>().First();
            }
        }
    }
}

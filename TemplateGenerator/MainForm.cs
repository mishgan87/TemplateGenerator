using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

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
        private void ButtonGenerate_Click(object sender, EventArgs e)
        {
            ExcelDocument.ImportTreeView(treeView);
        }
        private void TreeView_MouseClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                treeView.SelectedNode = e.Node;
                
                System.Windows.Forms.ContextMenuStrip menu = new System.Windows.Forms.ContextMenuStrip();
                menu.Font = treeView.Font;

                string nameAdd = "";
                string nameRemove = "";
                if (treeView.SelectedNode.Level == 0)
                {
                    nameAdd = "New group";
                    nameRemove = "Remove template";
                }
                if (treeView.SelectedNode.Level == 1)
                {
                    nameAdd = "New attribute";
                    nameRemove = "Remove group";
                }
                if (treeView.SelectedNode.Level == 2)
                {
                    nameAdd = "New parameter";
                    nameRemove = "Remove attribute";
                }
                if (treeView.SelectedNode.Level == 3)
                {
                    nameRemove = "Remove paremeter";
                }

                if (treeView.SelectedNode.Level < 3)
                {
                    ToolStripItem itemAdd = menu.Items.Add($"{nameAdd}");
                    itemAdd.Image = global::TemplateGenerator.Properties.Resources.plus;
                    itemAdd.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None;
                    itemAdd.Click += new EventHandler(this.AddNode);
                }

                if (treeView.SelectedNode.Level > 0)
                {
                    ToolStripItem itemRemove = menu.Items.Add($"{nameRemove}");
                    itemRemove.Image = global::TemplateGenerator.Properties.Resources.minus;
                    itemRemove.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None;
                    itemRemove.Click += new EventHandler(this.RemoveNode);
                }
                /*
                ToolStripItem itemCopy = menu.Items.Add("Copy");
                itemCopy.Image = global::TemplateGenerator.Properties.Resources.copy;
                itemCopy.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None;
                itemCopy.Click += new EventHandler(this.CopyNode);
                */
                if (ClipboardNode != null)
                {
                    ToolStripItem itemPaste = menu.Items.Add("Paste");
                    itemPaste.Image = global::TemplateGenerator.Properties.Resources.paste;
                    itemPaste.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None;
                    itemPaste.Click += new EventHandler(this.PasteNode);
                }

                menu.Show((System.Windows.Forms.Control)sender, new Point(e.X, e.Y));
            }
        }

        private void TreeView_NodeMouseDoubleClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            TreeNode node = e.Node;
            
            if (node == null)
            {
                return;
            }

            if (!node.IsExpanded)
            {
                node.Expand();
            }
            else
            {
                node.Collapse();
            }

            treeView.SelectedNode = node;
            treeView.LabelEdit = true;
            if (!node.IsEditing)
            {
                node.BeginEdit();
            }
        }

        private void ButtonImport_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.InitialDirectory = ".";
                openFileDialog.Filter = "MsWord files (*.docx)|*.docx";
                openFileDialog.FilterIndex = 2;
                openFileDialog.RestoreDirectory = true;
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string filename = openFileDialog.FileName;                    
                    string text = filename; // имя файла без расширения и полного пути
                    text = text.Remove(text.LastIndexOf("."), text.Length - text.LastIndexOf("."));
                    text = text.Substring(text.LastIndexOf("\\") + 1, text.Length - text.LastIndexOf("\\") - 1);

                    WordDocument.ImportExportToTreeView(filename, treeView);
                }
            }
        }
    }
}

namespace TemplateGenerator
{
    partial class MainForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.treeView = new System.Windows.Forms.TreeView();
            this.menu = new System.Windows.Forms.ToolStrip();
            this.buttonNewTemplate = new System.Windows.Forms.ToolStripButton();
            this.buttonOpenTemplate = new System.Windows.Forms.ToolStripButton();
            this.buttonSaveTemplate = new System.Windows.Forms.ToolStripButton();
            this.buttonGenerate = new System.Windows.Forms.ToolStripButton();
            this.buttonImport = new System.Windows.Forms.ToolStripButton();
            this.layout = new System.Windows.Forms.SplitContainer();
            this.gridView = new System.Windows.Forms.DataGridView();
            this.menu.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.layout)).BeginInit();
            this.layout.Panel1.SuspendLayout();
            this.layout.Panel2.SuspendLayout();
            this.layout.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.gridView)).BeginInit();
            this.SuspendLayout();
            // 
            // treeView
            // 
            this.treeView.Dock = System.Windows.Forms.DockStyle.Left;
            this.treeView.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.treeView.Location = new System.Drawing.Point(0, 0);
            this.treeView.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.treeView.Name = "treeView";
            this.treeView.Size = new System.Drawing.Size(527, 489);
            this.treeView.TabIndex = 0;
            this.treeView.NodeMouseClick += new System.Windows.Forms.TreeNodeMouseClickEventHandler(this.TreeView_MouseClick);
            this.treeView.NodeMouseDoubleClick += new System.Windows.Forms.TreeNodeMouseClickEventHandler(this.TreeView_NodeMouseDoubleClick);
            // 
            // menu
            // 
            this.menu.Dock = System.Windows.Forms.DockStyle.Fill;
            this.menu.GripStyle = System.Windows.Forms.ToolStripGripStyle.Hidden;
            this.menu.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.buttonNewTemplate,
            this.buttonOpenTemplate,
            this.buttonSaveTemplate,
            this.buttonGenerate,
            this.buttonImport});
            this.menu.LayoutStyle = System.Windows.Forms.ToolStripLayoutStyle.HorizontalStackWithOverflow;
            this.menu.Location = new System.Drawing.Point(0, 0);
            this.menu.Name = "menu";
            this.menu.Size = new System.Drawing.Size(1067, 60);
            this.menu.TabIndex = 1;
            this.menu.Text = "toolStrip1";
            // 
            // buttonNewTemplate
            // 
            this.buttonNewTemplate.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.buttonNewTemplate.Image = global::TemplateGenerator.Properties.Resources._new;
            this.buttonNewTemplate.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None;
            this.buttonNewTemplate.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.buttonNewTemplate.Name = "buttonNewTemplate";
            this.buttonNewTemplate.Size = new System.Drawing.Size(36, 57);
            this.buttonNewTemplate.ToolTipText = "Create";
            this.buttonNewTemplate.Click += new System.EventHandler(this.NewTemplate);
            // 
            // buttonOpenTemplate
            // 
            this.buttonOpenTemplate.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.buttonOpenTemplate.Image = global::TemplateGenerator.Properties.Resources.open;
            this.buttonOpenTemplate.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None;
            this.buttonOpenTemplate.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.buttonOpenTemplate.Name = "buttonOpenTemplate";
            this.buttonOpenTemplate.Size = new System.Drawing.Size(36, 57);
            this.buttonOpenTemplate.ToolTipText = "Open";
            // 
            // buttonSaveTemplate
            // 
            this.buttonSaveTemplate.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.buttonSaveTemplate.Image = global::TemplateGenerator.Properties.Resources.save;
            this.buttonSaveTemplate.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None;
            this.buttonSaveTemplate.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.buttonSaveTemplate.Name = "buttonSaveTemplate";
            this.buttonSaveTemplate.Size = new System.Drawing.Size(36, 57);
            this.buttonSaveTemplate.ToolTipText = "Save";
            // 
            // buttonGenerate
            // 
            this.buttonGenerate.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.buttonGenerate.Image = global::TemplateGenerator.Properties.Resources.excel;
            this.buttonGenerate.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None;
            this.buttonGenerate.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.buttonGenerate.Name = "buttonGenerate";
            this.buttonGenerate.Size = new System.Drawing.Size(36, 57);
            this.buttonGenerate.ToolTipText = "Export";
            this.buttonGenerate.Click += new System.EventHandler(this.ButtonGenerate_Click);
            // 
            // buttonImport
            // 
            this.buttonImport.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.buttonImport.Image = global::TemplateGenerator.Properties.Resources.word;
            this.buttonImport.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None;
            this.buttonImport.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.buttonImport.Name = "buttonImport";
            this.buttonImport.Size = new System.Drawing.Size(36, 57);
            this.buttonImport.ToolTipText = "Import";
            this.buttonImport.Click += new System.EventHandler(this.ButtonImport_Click);
            // 
            // layout
            // 
            this.layout.Dock = System.Windows.Forms.DockStyle.Fill;
            this.layout.FixedPanel = System.Windows.Forms.FixedPanel.Panel1;
            this.layout.IsSplitterFixed = true;
            this.layout.Location = new System.Drawing.Point(0, 0);
            this.layout.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.layout.Name = "layout";
            this.layout.Orientation = System.Windows.Forms.Orientation.Horizontal;
            // 
            // layout.Panel1
            // 
            this.layout.Panel1.Controls.Add(this.menu);
            // 
            // layout.Panel2
            // 
            this.layout.Panel2.Controls.Add(this.gridView);
            this.layout.Panel2.Controls.Add(this.treeView);
            this.layout.Size = new System.Drawing.Size(1067, 554);
            this.layout.SplitterDistance = 60;
            this.layout.SplitterWidth = 5;
            this.layout.TabIndex = 2;
            // 
            // gridView
            // 
            this.gridView.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.gridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.gridView.Location = new System.Drawing.Point(540, 0);
            this.gridView.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.gridView.Name = "gridView";
            this.gridView.Size = new System.Drawing.Size(527, 489);
            this.gridView.TabIndex = 1;
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1067, 554);
            this.Controls.Add(this.layout);
            this.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.Name = "MainForm";
            this.Text = "Template Generator";
            this.menu.ResumeLayout(false);
            this.menu.PerformLayout();
            this.layout.Panel1.ResumeLayout(false);
            this.layout.Panel1.PerformLayout();
            this.layout.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.layout)).EndInit();
            this.layout.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.gridView)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TreeView treeView;
        private System.Windows.Forms.ToolStrip menu;
        private System.Windows.Forms.ToolStripButton buttonNewTemplate;
        private System.Windows.Forms.ToolStripButton buttonOpenTemplate;
        private System.Windows.Forms.ToolStripButton buttonSaveTemplate;
        private System.Windows.Forms.SplitContainer layout;
        private System.Windows.Forms.ToolStripButton buttonGenerate;
        private System.Windows.Forms.ToolStripButton buttonImport;
        private System.Windows.Forms.DataGridView gridView;
    }
}


﻿namespace TemplateGenerator
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
            this.buttonOpenTemplate = new System.Windows.Forms.ToolStripButton();
            this.buttonGenerate = new System.Windows.Forms.ToolStripButton();
            this.buttonSaveTemplate = new System.Windows.Forms.ToolStripButton();
            this.buttonNewTemplate = new System.Windows.Forms.ToolStripButton();
            this.layout = new System.Windows.Forms.SplitContainer();
            this.menu.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.layout)).BeginInit();
            this.layout.Panel1.SuspendLayout();
            this.layout.Panel2.SuspendLayout();
            this.layout.SuspendLayout();
            this.SuspendLayout();
            // 
            // treeView
            // 
            this.treeView.Dock = System.Windows.Forms.DockStyle.Fill;
            this.treeView.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.treeView.Location = new System.Drawing.Point(0, 0);
            this.treeView.Name = "treeView";
            this.treeView.Size = new System.Drawing.Size(800, 386);
            this.treeView.TabIndex = 0;
            this.treeView.NodeMouseClick += new System.Windows.Forms.TreeNodeMouseClickEventHandler(this.TreeView_MouseClick);
            // 
            // menu
            // 
            this.menu.Dock = System.Windows.Forms.DockStyle.Fill;
            this.menu.GripStyle = System.Windows.Forms.ToolStripGripStyle.Hidden;
            this.menu.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.buttonOpenTemplate,
            this.buttonGenerate,
            this.buttonSaveTemplate,
            this.buttonNewTemplate});
            this.menu.LayoutStyle = System.Windows.Forms.ToolStripLayoutStyle.HorizontalStackWithOverflow;
            this.menu.Location = new System.Drawing.Point(0, 0);
            this.menu.Name = "menu";
            this.menu.Size = new System.Drawing.Size(800, 60);
            this.menu.TabIndex = 1;
            this.menu.Text = "toolStrip1";
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
            // layout
            // 
            this.layout.Dock = System.Windows.Forms.DockStyle.Fill;
            this.layout.FixedPanel = System.Windows.Forms.FixedPanel.Panel1;
            this.layout.Location = new System.Drawing.Point(0, 0);
            this.layout.Name = "layout";
            this.layout.Orientation = System.Windows.Forms.Orientation.Horizontal;
            // 
            // layout.Panel1
            // 
            this.layout.Panel1.Controls.Add(this.menu);
            // 
            // layout.Panel2
            // 
            this.layout.Panel2.Controls.Add(this.treeView);
            this.layout.Size = new System.Drawing.Size(800, 450);
            this.layout.SplitterDistance = 60;
            this.layout.TabIndex = 2;
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.layout);
            this.Name = "MainForm";
            this.Text = "Template Generator";
            this.menu.ResumeLayout(false);
            this.menu.PerformLayout();
            this.layout.Panel1.ResumeLayout(false);
            this.layout.Panel1.PerformLayout();
            this.layout.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.layout)).EndInit();
            this.layout.ResumeLayout(false);
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
    }
}


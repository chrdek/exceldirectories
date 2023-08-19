namespace ExcelDirectories
{
    partial class Form1
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
            this.components = new System.ComponentModel.Container();
            this.treeView = new System.Windows.Forms.TreeView();
            this.button1 = new System.Windows.Forms.Button();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.spreadsheet1 = new Spire.Spreadsheet.Forms.Spreadsheet();
            this.SuspendLayout();
            // 
            // treeView
            // 
            this.treeView.Location = new System.Drawing.Point(13, 76);
            this.treeView.Name = "treeView";
            this.treeView.Size = new System.Drawing.Size(272, 567);
            this.treeView.TabIndex = 0;
            // 
            // button1
            // 
            this.button1.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.button1.Font = new System.Drawing.Font("Microsoft Sans Serif", 15F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button1.ForeColor = System.Drawing.SystemColors.ControlDarkDark;
            this.button1.Location = new System.Drawing.Point(13, 13);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(272, 57);
            this.button1.TabIndex = 2;
            this.button1.Text = "Select Directory...";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // folderBrowserDialog1
            // 
            this.folderBrowserDialog1.SelectedPath = "C:\\Users\\standarduser\\Desktop";
            // 
            // spreadsheet1
            // 
            this.spreadsheet1.ActiveSheetIndex = 0;
            this.spreadsheet1.HorizontalScrollBarVisibility = true;
            this.spreadsheet1.Location = new System.Drawing.Point(291, 13);
            this.spreadsheet1.Name = "spreadsheet1";
            this.spreadsheet1.ScrollBarsVisibility = true;
            this.spreadsheet1.SheetTabControlWidth = 400;
            this.spreadsheet1.Size = new System.Drawing.Size(870, 630);
            this.spreadsheet1.TabIndex = 3;
            this.spreadsheet1.VerticalScrollBarVisibility = true;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1173, 655);
            this.Controls.Add(this.spreadsheet1);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.treeView);
            this.Name = "Form1";
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TreeView treeView;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        private Spire.Spreadsheet.Forms.Spreadsheet spreadsheet1;
    }
}


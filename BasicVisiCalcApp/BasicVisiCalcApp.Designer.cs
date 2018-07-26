namespace NS_BasicVisiCalcApp
{
    partial class BasicVisiCalcApp
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            this.dbGridTable = new System.Windows.Forms.DataGridView();
            this.A = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.B = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.C = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.D = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.E = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.F = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.gBoxShortCuts = new System.Windows.Forms.GroupBox();
            this.btnExit = new System.Windows.Forms.Button();
            this.btnEdit = new System.Windows.Forms.Button();
            this.btnHelp = new System.Windows.Forms.Button();
            this.btnNew = new System.Windows.Forms.Button();
            this.btnSave = new System.Windows.Forms.Button();
            this.t1 = new System.Windows.Forms.TextBox();
            this.t3 = new System.Windows.Forms.TextBox();
            this.gBoxCurrCell = new System.Windows.Forms.GroupBox();
            this.btnUpdate = new System.Windows.Forms.Button();
            this.t2 = new System.Windows.Forms.TextBox();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.fileMenu = new System.Windows.Forms.ToolStripMenuItem();
            this.newFileToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.openFileToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.saveFileToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.saveFileAsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.exitToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.editMenu = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem2 = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator3 = new System.Windows.Forms.ToolStripSeparator();
            this.pasteToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem3 = new System.Windows.Forms.ToolStripMenuItem();
            this.copyRangeToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem5 = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem4 = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator2 = new System.Windows.Forms.ToolStripSeparator();
            this.insertRowToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.insertColumnToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.deleteRowToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.deleteColumnToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.excludeRowsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.excludeColumnsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.helpMenu = new System.Windows.Forms.ToolStripMenuItem();
            this.showHelpToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.functionsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.aboutToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            ((System.ComponentModel.ISupportInitialize)(this.dbGridTable)).BeginInit();
            this.groupBox1.SuspendLayout();
            this.gBoxShortCuts.SuspendLayout();
            this.gBoxCurrCell.SuspendLayout();
            this.menuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // dbGridTable
            // 
            this.dbGridTable.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dbGridTable.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.A,
            this.B,
            this.C,
            this.D,
            this.E,
            this.F});
            this.dbGridTable.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dbGridTable.Location = new System.Drawing.Point(3, 16);
            this.dbGridTable.Name = "dbGridTable";
            this.dbGridTable.RowHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.Single;
            this.dbGridTable.RowHeadersWidth = 50;
            this.dbGridTable.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing;
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.Color.Black;
            this.dbGridTable.RowsDefaultCellStyle = dataGridViewCellStyle1;
            this.dbGridTable.Size = new System.Drawing.Size(572, 425);
            this.dbGridTable.TabIndex = 0;
            this.dbGridTable.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dbGridTable_CellContentClick);
            // 
            // A
            // 
            this.A.HeaderText = "A";
            this.A.Name = "A";
            // 
            // B
            // 
            this.B.HeaderText = "B";
            this.B.Name = "B";
            // 
            // C
            // 
            this.C.HeaderText = "C";
            this.C.Name = "C";
            // 
            // D
            // 
            this.D.HeaderText = "D";
            this.D.Name = "D";
            // 
            // E
            // 
            this.E.HeaderText = "E";
            this.E.Name = "E";
            // 
            // F
            // 
            this.F.HeaderText = "F";
            this.F.Name = "F";
            // 
            // groupBox1
            // 
            this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox1.Controls.Add(this.dbGridTable);
            this.groupBox1.Location = new System.Drawing.Point(12, 82);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(578, 444);
            this.groupBox1.TabIndex = 1;
            this.groupBox1.TabStop = false;
            // 
            // gBoxShortCuts
            // 
            this.gBoxShortCuts.Controls.Add(this.btnExit);
            this.gBoxShortCuts.Controls.Add(this.btnEdit);
            this.gBoxShortCuts.Controls.Add(this.btnHelp);
            this.gBoxShortCuts.Controls.Add(this.btnNew);
            this.gBoxShortCuts.Controls.Add(this.btnSave);
            this.gBoxShortCuts.Location = new System.Drawing.Point(376, 27);
            this.gBoxShortCuts.Name = "gBoxShortCuts";
            this.gBoxShortCuts.Size = new System.Drawing.Size(211, 49);
            this.gBoxShortCuts.TabIndex = 2;
            this.gBoxShortCuts.TabStop = false;
            this.gBoxShortCuts.Text = "ShortCuts";
            // 
            // btnExit
            // 
            this.btnExit.Dock = System.Windows.Forms.DockStyle.Left;
            this.btnExit.Location = new System.Drawing.Point(163, 16);
            this.btnExit.Name = "btnExit";
            this.btnExit.Size = new System.Drawing.Size(40, 30);
            this.btnExit.TabIndex = 3;
            this.btnExit.Text = "Exit";
            this.btnExit.UseVisualStyleBackColor = true;
            this.btnExit.Click += new System.EventHandler(this.btnExit_Click);
            // 
            // btnEdit
            // 
            this.btnEdit.Dock = System.Windows.Forms.DockStyle.Left;
            this.btnEdit.Location = new System.Drawing.Point(123, 16);
            this.btnEdit.Name = "btnEdit";
            this.btnEdit.Size = new System.Drawing.Size(40, 30);
            this.btnEdit.TabIndex = 2;
            this.btnEdit.Text = "Edit";
            this.btnEdit.UseVisualStyleBackColor = true;
            this.btnEdit.Click += new System.EventHandler(this.btnEdit_Click);
            // 
            // btnHelp
            // 
            this.btnHelp.Dock = System.Windows.Forms.DockStyle.Left;
            this.btnHelp.Location = new System.Drawing.Point(83, 16);
            this.btnHelp.Name = "btnHelp";
            this.btnHelp.Size = new System.Drawing.Size(40, 30);
            this.btnHelp.TabIndex = 1;
            this.btnHelp.Text = "Help";
            this.btnHelp.UseVisualStyleBackColor = true;
            this.btnHelp.Click += new System.EventHandler(this.btnHelp_Click);
            // 
            // btnNew
            // 
            this.btnNew.Dock = System.Windows.Forms.DockStyle.Left;
            this.btnNew.Location = new System.Drawing.Point(43, 16);
            this.btnNew.Name = "btnNew";
            this.btnNew.Size = new System.Drawing.Size(40, 30);
            this.btnNew.TabIndex = 0;
            this.btnNew.Text = "New";
            this.btnNew.UseVisualStyleBackColor = true;
            this.btnNew.Click += new System.EventHandler(this.btnNew_Click);
            // 
            // btnSave
            // 
            this.btnSave.Dock = System.Windows.Forms.DockStyle.Left;
            this.btnSave.Location = new System.Drawing.Point(3, 16);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(40, 30);
            this.btnSave.TabIndex = 0;
            this.btnSave.Text = "Save";
            this.btnSave.UseVisualStyleBackColor = true;
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // t1
            // 
            this.t1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.t1.Location = new System.Drawing.Point(6, 16);
            this.t1.Name = "t1";
            this.t1.Size = new System.Drawing.Size(47, 20);
            this.t1.TabIndex = 3;
            // 
            // t3
            // 
            this.t3.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.t3.Location = new System.Drawing.Point(141, 16);
            this.t3.Name = "t3";
            this.t3.Size = new System.Drawing.Size(154, 20);
            this.t3.TabIndex = 4;
            // 
            // gBoxCurrCell
            // 
            this.gBoxCurrCell.Controls.Add(this.btnUpdate);
            this.gBoxCurrCell.Controls.Add(this.t1);
            this.gBoxCurrCell.Controls.Add(this.t2);
            this.gBoxCurrCell.Controls.Add(this.t3);
            this.gBoxCurrCell.Location = new System.Drawing.Point(12, 27);
            this.gBoxCurrCell.Name = "gBoxCurrCell";
            this.gBoxCurrCell.Size = new System.Drawing.Size(358, 49);
            this.gBoxCurrCell.TabIndex = 5;
            this.gBoxCurrCell.TabStop = false;
            this.gBoxCurrCell.Text = "   [Cell]           [Value]                           [Formula/Text]";
            // 
            // btnUpdate
            // 
            this.btnUpdate.Location = new System.Drawing.Point(301, 10);
            this.btnUpdate.Name = "btnUpdate";
            this.btnUpdate.Size = new System.Drawing.Size(51, 30);
            this.btnUpdate.TabIndex = 5;
            this.btnUpdate.Text = "Update";
            this.btnUpdate.UseVisualStyleBackColor = true;
            this.btnUpdate.Click += new System.EventHandler(this.btnUpdate_Click);
            // 
            // t2
            // 
            this.t2.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.t2.Location = new System.Drawing.Point(59, 16);
            this.t2.Name = "t2";
            this.t2.Size = new System.Drawing.Size(76, 20);
            this.t2.TabIndex = 4;
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.fileMenu,
            this.editMenu,
            this.helpMenu});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(602, 24);
            this.menuStrip1.TabIndex = 6;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // fileMenu
            // 
            this.fileMenu.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.newFileToolStripMenuItem,
            this.openFileToolStripMenuItem,
            this.saveFileToolStripMenuItem,
            this.saveFileAsToolStripMenuItem,
            this.exitToolStripMenuItem});
            this.fileMenu.Name = "fileMenu";
            this.fileMenu.Size = new System.Drawing.Size(37, 20);
            this.fileMenu.Text = "&File";
            // 
            // newFileToolStripMenuItem
            // 
            this.newFileToolStripMenuItem.Name = "newFileToolStripMenuItem";
            this.newFileToolStripMenuItem.Size = new System.Drawing.Size(135, 22);
            this.newFileToolStripMenuItem.Text = "&New File";
            this.newFileToolStripMenuItem.Click += new System.EventHandler(this.menuFile_Click);
            // 
            // openFileToolStripMenuItem
            // 
            this.openFileToolStripMenuItem.Name = "openFileToolStripMenuItem";
            this.openFileToolStripMenuItem.Size = new System.Drawing.Size(135, 22);
            this.openFileToolStripMenuItem.Text = "&Open File";
            this.openFileToolStripMenuItem.Click += new System.EventHandler(this.menuFile_Click);
            // 
            // saveFileToolStripMenuItem
            // 
            this.saveFileToolStripMenuItem.Name = "saveFileToolStripMenuItem";
            this.saveFileToolStripMenuItem.Size = new System.Drawing.Size(135, 22);
            this.saveFileToolStripMenuItem.Text = "&Save File";
            this.saveFileToolStripMenuItem.Click += new System.EventHandler(this.menuFile_Click);
            // 
            // saveFileAsToolStripMenuItem
            // 
            this.saveFileAsToolStripMenuItem.Name = "saveFileAsToolStripMenuItem";
            this.saveFileAsToolStripMenuItem.Size = new System.Drawing.Size(135, 22);
            this.saveFileAsToolStripMenuItem.Text = "Save File &As";
            this.saveFileAsToolStripMenuItem.Click += new System.EventHandler(this.menuFile_Click);
            // 
            // exitToolStripMenuItem
            // 
            this.exitToolStripMenuItem.Name = "exitToolStripMenuItem";
            this.exitToolStripMenuItem.Size = new System.Drawing.Size(135, 22);
            this.exitToolStripMenuItem.Text = "E&xit";
            this.exitToolStripMenuItem.Click += new System.EventHandler(this.exitToolStripMenuItem_Click);
            // 
            // editMenu
            // 
            this.editMenu.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripMenuItem1,
            this.toolStripMenuItem2,
            this.toolStripSeparator3,
            this.pasteToolStripMenuItem,
            this.toolStripMenuItem3,
            this.copyRangeToolStripMenuItem,
            this.toolStripMenuItem5,
            this.toolStripMenuItem4,
            this.toolStripSeparator2,
            this.insertRowToolStripMenuItem,
            this.insertColumnToolStripMenuItem,
            this.deleteRowToolStripMenuItem,
            this.deleteColumnToolStripMenuItem,
            this.excludeRowsToolStripMenuItem,
            this.excludeColumnsToolStripMenuItem});
            this.editMenu.Name = "editMenu";
            this.editMenu.Size = new System.Drawing.Size(39, 20);
            this.editMenu.Text = "Edit";
            // 
            // toolStripMenuItem1
            // 
            this.toolStripMenuItem1.Name = "toolStripMenuItem1";
            this.toolStripMenuItem1.Size = new System.Drawing.Size(201, 22);
            this.toolStripMenuItem1.Text = "&Edit Cell";
            this.toolStripMenuItem1.Click += new System.EventHandler(this.menuEdit_Click);
            // 
            // toolStripMenuItem2
            // 
            this.toolStripMenuItem2.Name = "toolStripMenuItem2";
            this.toolStripMenuItem2.Size = new System.Drawing.Size(201, 22);
            this.toolStripMenuItem2.Text = "&View Cell";
            this.toolStripMenuItem2.Click += new System.EventHandler(this.menuEdit_Click);
            // 
            // toolStripSeparator3
            // 
            this.toolStripSeparator3.Name = "toolStripSeparator3";
            this.toolStripSeparator3.Size = new System.Drawing.Size(198, 6);
            // 
            // pasteToolStripMenuItem
            // 
            this.pasteToolStripMenuItem.Name = "pasteToolStripMenuItem";
            this.pasteToolStripMenuItem.Size = new System.Drawing.Size(201, 22);
            this.pasteToolStripMenuItem.Text = "&Format Range";
            this.pasteToolStripMenuItem.Click += new System.EventHandler(this.menuEdit_Click);
            // 
            // toolStripMenuItem3
            // 
            this.toolStripMenuItem3.Name = "toolStripMenuItem3";
            this.toolStripMenuItem3.Size = new System.Drawing.Size(201, 22);
            this.toolStripMenuItem3.Text = "&Delete Range";
            this.toolStripMenuItem3.Click += new System.EventHandler(this.menuEdit_Click);
            // 
            // copyRangeToolStripMenuItem
            // 
            this.copyRangeToolStripMenuItem.Name = "copyRangeToolStripMenuItem";
            this.copyRangeToolStripMenuItem.Size = new System.Drawing.Size(201, 22);
            this.copyRangeToolStripMenuItem.Text = "&Copy Range";
            this.copyRangeToolStripMenuItem.Click += new System.EventHandler(this.menuEdit_Click);
            // 
            // toolStripMenuItem5
            // 
            this.toolStripMenuItem5.Name = "toolStripMenuItem5";
            this.toolStripMenuItem5.Size = new System.Drawing.Size(201, 22);
            this.toolStripMenuItem5.Text = "&Resize Columns";
            this.toolStripMenuItem5.Click += new System.EventHandler(this.menuEdit_Click);
            // 
            // toolStripMenuItem4
            // 
            this.toolStripMenuItem4.Name = "toolStripMenuItem4";
            this.toolStripMenuItem4.Size = new System.Drawing.Size(201, 22);
            this.toolStripMenuItem4.Text = "&Hide/Unhide Rows/Cols";
            this.toolStripMenuItem4.Click += new System.EventHandler(this.menuEdit_Click);
            // 
            // toolStripSeparator2
            // 
            this.toolStripSeparator2.Name = "toolStripSeparator2";
            this.toolStripSeparator2.Size = new System.Drawing.Size(198, 6);
            // 
            // insertRowToolStripMenuItem
            // 
            this.insertRowToolStripMenuItem.Name = "insertRowToolStripMenuItem";
            this.insertRowToolStripMenuItem.Size = new System.Drawing.Size(201, 22);
            this.insertRowToolStripMenuItem.Text = "&Insert Row";
            this.insertRowToolStripMenuItem.Click += new System.EventHandler(this.menuEdit_Click);
            // 
            // insertColumnToolStripMenuItem
            // 
            this.insertColumnToolStripMenuItem.Name = "insertColumnToolStripMenuItem";
            this.insertColumnToolStripMenuItem.Size = new System.Drawing.Size(201, 22);
            this.insertColumnToolStripMenuItem.Text = "I&nsert Column";
            this.insertColumnToolStripMenuItem.Click += new System.EventHandler(this.menuEdit_Click);
            // 
            // deleteRowToolStripMenuItem
            // 
            this.deleteRowToolStripMenuItem.Name = "deleteRowToolStripMenuItem";
            this.deleteRowToolStripMenuItem.Size = new System.Drawing.Size(201, 22);
            this.deleteRowToolStripMenuItem.Text = "De&lete Row(s)";
            this.deleteRowToolStripMenuItem.Click += new System.EventHandler(this.menuEdit_Click);
            // 
            // deleteColumnToolStripMenuItem
            // 
            this.deleteColumnToolStripMenuItem.Name = "deleteColumnToolStripMenuItem";
            this.deleteColumnToolStripMenuItem.Size = new System.Drawing.Size(201, 22);
            this.deleteColumnToolStripMenuItem.Text = "Dele&te Column(s)";
            this.deleteColumnToolStripMenuItem.Click += new System.EventHandler(this.menuEdit_Click);
            // 
            // excludeRowsToolStripMenuItem
            // 
            this.excludeRowsToolStripMenuItem.Name = "excludeRowsToolStripMenuItem";
            this.excludeRowsToolStripMenuItem.Size = new System.Drawing.Size(201, 22);
            this.excludeRowsToolStripMenuItem.Text = "E&xclude Row(s)";
            this.excludeRowsToolStripMenuItem.Click += new System.EventHandler(this.menuEdit_Click);
            // 
            // excludeColumnsToolStripMenuItem
            // 
            this.excludeColumnsToolStripMenuItem.Name = "excludeColumnsToolStripMenuItem";
            this.excludeColumnsToolStripMenuItem.Size = new System.Drawing.Size(201, 22);
            this.excludeColumnsToolStripMenuItem.Text = "Excl&ude Column(s)";
            this.excludeColumnsToolStripMenuItem.Click += new System.EventHandler(this.menuEdit_Click);
            // 
            // helpMenu
            // 
            this.helpMenu.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.showHelpToolStripMenuItem,
            this.functionsToolStripMenuItem,
            this.aboutToolStripMenuItem});
            this.helpMenu.Name = "helpMenu";
            this.helpMenu.Size = new System.Drawing.Size(44, 20);
            this.helpMenu.Text = "Help";
            // 
            // showHelpToolStripMenuItem
            // 
            this.showHelpToolStripMenuItem.Name = "showHelpToolStripMenuItem";
            this.showHelpToolStripMenuItem.Size = new System.Drawing.Size(131, 22);
            this.showHelpToolStripMenuItem.Text = "Show &Help";
            this.showHelpToolStripMenuItem.Click += new System.EventHandler(this.menuHelp_Click);
            // 
            // functionsToolStripMenuItem
            // 
            this.functionsToolStripMenuItem.Name = "functionsToolStripMenuItem";
            this.functionsToolStripMenuItem.Size = new System.Drawing.Size(131, 22);
            this.functionsToolStripMenuItem.Text = "&Functions";
            this.functionsToolStripMenuItem.Click += new System.EventHandler(this.menuHelp_Click);
            // 
            // aboutToolStripMenuItem
            // 
            this.aboutToolStripMenuItem.Name = "aboutToolStripMenuItem";
            this.aboutToolStripMenuItem.Size = new System.Drawing.Size(131, 22);
            this.aboutToolStripMenuItem.Text = "&About";
            this.aboutToolStripMenuItem.Click += new System.EventHandler(this.menuHelp_Click);
            // 
            // Form4
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(602, 538);
            this.Controls.Add(this.gBoxCurrCell);
            this.Controls.Add(this.gBoxShortCuts);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.menuStrip1);
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "Form4";
            this.Text = "Basic Visicalc Spreadsheet";
            this.Load += new System.EventHandler(this.Form4_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dbGridTable)).EndInit();
            this.groupBox1.ResumeLayout(false);
            this.gBoxShortCuts.ResumeLayout(false);
            this.gBoxCurrCell.ResumeLayout(false);
            this.gBoxCurrCell.PerformLayout();
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DataGridView dbGridTable;
        private System.Windows.Forms.DataGridViewTextBoxColumn A;
        private System.Windows.Forms.DataGridViewTextBoxColumn B;
        private System.Windows.Forms.DataGridViewTextBoxColumn C;
        private System.Windows.Forms.DataGridViewTextBoxColumn D;
        private System.Windows.Forms.DataGridViewTextBoxColumn E;
        private System.Windows.Forms.DataGridViewTextBoxColumn F;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.GroupBox gBoxShortCuts;
        private System.Windows.Forms.TextBox t1;
        private System.Windows.Forms.TextBox t3;
        private System.Windows.Forms.Button btnExit;
        private System.Windows.Forms.Button btnEdit;
        private System.Windows.Forms.Button btnHelp;
        private System.Windows.Forms.Button btnNew;
        private System.Windows.Forms.Button btnSave;
        private System.Windows.Forms.GroupBox gBoxCurrCell;
        private System.Windows.Forms.TextBox t2;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem fileMenu;
        private System.Windows.Forms.ToolStripMenuItem newFileToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem openFileToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem saveFileToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem saveFileAsToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem exitToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem editMenu;
        private System.Windows.Forms.ToolStripMenuItem insertRowToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem insertColumnToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem copyRangeToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem pasteToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem deleteRowToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem deleteColumnToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem excludeRowsToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem excludeColumnsToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem helpMenu;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem1;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem2;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator3;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem3;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem5;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem4;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator2;
        private System.Windows.Forms.ToolStripMenuItem showHelpToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem functionsToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem aboutToolStripMenuItem;
        private System.Windows.Forms.Button btnUpdate;
    }
}
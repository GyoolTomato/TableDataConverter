namespace TableDataConverter
{
    partial class Form1
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            button1 = new Button();
            button2 = new Button();
            listBox1 = new CheckedListBox();
            label1 = new Label();
            labelBytesPath = new Label();
            textBoxBytesPath = new TextBox();
            buttonBrowseBytesPath = new Button();
            labelScriptPath = new Label();
            textBoxScriptPath = new TextBox();
            buttonBrowseScriptPath = new Button();
            buttonOpenBytesPath = new Button();
            buttonOpenScriptPath = new Button();
            labelDatabasePath = new Label();
            textBoxDatabasePath = new TextBox();
            buttonBrowseDatabase = new Button();
            buttonOpenDatabasePath = new Button();
            buttonSelectAllTables = new Button();
            buttonClearTables = new Button();
            buttonImportDatabase = new Button();
            labelTables = new Label();
            SuspendLayout();
            // 
            // button1
            // 
            button1.Location = new Point(13, 816);
            button1.Margin = new Padding(4);
            button1.Name = "button1";
            button1.Size = new Size(180, 72);
            button1.TabIndex = 0;
            button1.Text = "목록 새로고침";
            button1.UseVisualStyleBackColor = true;
            button1.Click += OnBtn_Refresh;
            // 
            // button2
            // 
            button2.Location = new Point(201, 816);
            button2.Margin = new Padding(4);
            button2.Name = "button2";
            button2.Size = new Size(532, 72);
            button2.TabIndex = 1;
            button2.Text = "전체 Data 변환";
            button2.UseVisualStyleBackColor = true;
            button2.Click += OnBtn_Confirm;
            // 
            // listBox1
            // 
            listBox1.CheckOnClick = true;
            listBox1.DrawMode = DrawMode.OwnerDrawFixed;
            listBox1.FormattingEnabled = true;
            listBox1.ItemHeight = 30;
            listBox1.Location = new Point(13, 42);
            listBox1.Name = "listBox1";
            listBox1.Size = new Size(720, 364);
            listBox1.TabIndex = 2;
            listBox1.DrawItem += OnTableDrawItem;
            listBox1.ItemCheck += OnTableItemCheck;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Font = new Font("맑은 고딕", 11F);
            label1.Location = new Point(13, 772);
            label1.Name = "label1";
            label1.Size = new Size(0, 36);
            label1.TabIndex = 3;
            // 
            // labelBytesPath
            // 
            labelBytesPath.AutoSize = true;
            labelBytesPath.Location = new Point(13, 424);
            labelBytesPath.Name = "labelBytesPath";
            labelBytesPath.Size = new Size(154, 30);
            labelBytesPath.TabIndex = 4;
            labelBytesPath.Text = ".bytes 저장 경로";
            // 
            // textBoxBytesPath
            // 
            textBoxBytesPath.Location = new Point(13, 457);
            textBoxBytesPath.Name = "textBoxBytesPath";
            textBoxBytesPath.ReadOnly = true;
            textBoxBytesPath.Size = new Size(490, 35);
            textBoxBytesPath.TabIndex = 5;
            // 
            // buttonBrowseBytesPath
            // 
            buttonBrowseBytesPath.Location = new Point(511, 455);
            buttonBrowseBytesPath.Name = "buttonBrowseBytesPath";
            buttonBrowseBytesPath.Size = new Size(102, 40);
            buttonBrowseBytesPath.TabIndex = 6;
            buttonBrowseBytesPath.Text = "찾아보기";
            buttonBrowseBytesPath.UseVisualStyleBackColor = true;
            buttonBrowseBytesPath.Click += OnBtn_BrowseBytesPath;
            // 
            // labelScriptPath
            // 
            labelScriptPath.AutoSize = true;
            labelScriptPath.Location = new Point(13, 507);
            labelScriptPath.Name = "labelScriptPath";
            labelScriptPath.Size = new Size(126, 30);
            labelScriptPath.TabIndex = 7;
            labelScriptPath.Text = ".cs 저장 경로";
            // 
            // textBoxScriptPath
            // 
            textBoxScriptPath.Location = new Point(13, 540);
            textBoxScriptPath.Name = "textBoxScriptPath";
            textBoxScriptPath.ReadOnly = true;
            textBoxScriptPath.Size = new Size(490, 35);
            textBoxScriptPath.TabIndex = 8;
            // 
            // buttonBrowseScriptPath
            // 
            buttonBrowseScriptPath.Location = new Point(511, 538);
            buttonBrowseScriptPath.Name = "buttonBrowseScriptPath";
            buttonBrowseScriptPath.Size = new Size(102, 40);
            buttonBrowseScriptPath.TabIndex = 9;
            buttonBrowseScriptPath.Text = "찾아보기";
            buttonBrowseScriptPath.UseVisualStyleBackColor = true;
            buttonBrowseScriptPath.Click += OnBtn_BrowseScriptPath;
            // 
            // buttonOpenBytesPath
            // 
            buttonOpenBytesPath.Location = new Point(621, 455);
            buttonOpenBytesPath.Name = "buttonOpenBytesPath";
            buttonOpenBytesPath.Size = new Size(112, 40);
            buttonOpenBytesPath.TabIndex = 17;
            buttonOpenBytesPath.Text = "폴더 열기";
            buttonOpenBytesPath.UseVisualStyleBackColor = true;
            buttonOpenBytesPath.Click += OnBtn_OpenBytesPath;
            // 
            // buttonOpenScriptPath
            // 
            buttonOpenScriptPath.Location = new Point(621, 538);
            buttonOpenScriptPath.Name = "buttonOpenScriptPath";
            buttonOpenScriptPath.Size = new Size(112, 40);
            buttonOpenScriptPath.TabIndex = 18;
            buttonOpenScriptPath.Text = "폴더 열기";
            buttonOpenScriptPath.UseVisualStyleBackColor = true;
            buttonOpenScriptPath.Click += OnBtn_OpenScriptPath;
            // 
            // labelDatabasePath
            // 
            labelDatabasePath.AutoSize = true;
            labelDatabasePath.Location = new Point(13, 590);
            labelDatabasePath.Name = "labelDatabasePath";
            labelDatabasePath.Size = new Size(125, 30);
            labelDatabasePath.TabIndex = 10;
            labelDatabasePath.Text = "SQLite DB 파일";
            // 
            // textBoxDatabasePath
            // 
            textBoxDatabasePath.Location = new Point(13, 623);
            textBoxDatabasePath.Name = "textBoxDatabasePath";
            textBoxDatabasePath.ReadOnly = true;
            textBoxDatabasePath.Size = new Size(490, 35);
            textBoxDatabasePath.TabIndex = 11;
            // 
            // buttonBrowseDatabase
            // 
            buttonBrowseDatabase.Location = new Point(511, 621);
            buttonBrowseDatabase.Name = "buttonBrowseDatabase";
            buttonBrowseDatabase.Size = new Size(102, 40);
            buttonBrowseDatabase.TabIndex = 12;
            buttonBrowseDatabase.Text = "찾아보기";
            buttonBrowseDatabase.UseVisualStyleBackColor = true;
            buttonBrowseDatabase.Click += OnBtn_BrowseDatabase;
            // 
            // buttonOpenDatabasePath
            // 
            buttonOpenDatabasePath.Location = new Point(621, 621);
            buttonOpenDatabasePath.Name = "buttonOpenDatabasePath";
            buttonOpenDatabasePath.Size = new Size(112, 40);
            buttonOpenDatabasePath.TabIndex = 19;
            buttonOpenDatabasePath.Text = "폴더 열기";
            buttonOpenDatabasePath.UseVisualStyleBackColor = true;
            buttonOpenDatabasePath.Click += OnBtn_OpenDatabasePath;
            // 
            // buttonSelectAllTables
            // 
            buttonSelectAllTables.Location = new Point(13, 676);
            buttonSelectAllTables.Name = "buttonSelectAllTables";
            buttonSelectAllTables.Size = new Size(120, 48);
            buttonSelectAllTables.TabIndex = 13;
            buttonSelectAllTables.Text = "전체 선택";
            buttonSelectAllTables.UseVisualStyleBackColor = true;
            buttonSelectAllTables.Click += OnBtn_SelectAllTables;
            // 
            // buttonClearTables
            // 
            buttonClearTables.Location = new Point(141, 676);
            buttonClearTables.Name = "buttonClearTables";
            buttonClearTables.Size = new Size(120, 48);
            buttonClearTables.TabIndex = 14;
            buttonClearTables.Text = "전체 해제";
            buttonClearTables.UseVisualStyleBackColor = true;
            buttonClearTables.Click += OnBtn_ClearTables;
            // 
            // buttonImportDatabase
            // 
            buttonImportDatabase.Location = new Point(269, 676);
            buttonImportDatabase.Name = "buttonImportDatabase";
            buttonImportDatabase.Size = new Size(464, 48);
            buttonImportDatabase.TabIndex = 15;
            buttonImportDatabase.Text = "선택한 테이블 DB 반영";
            buttonImportDatabase.UseVisualStyleBackColor = true;
            buttonImportDatabase.Click += OnBtn_ImportDatabase;
            // 
            // labelTables
            // 
            labelTables.AutoSize = true;
            labelTables.Location = new Point(13, 9);
            labelTables.Name = "labelTables";
            labelTables.Size = new Size(299, 30);
            labelTables.TabIndex = 16;
            labelTables.Text = "DB 반영 테이블 선택 (0xx, 9xx 제외)";
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(12F, 30F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(746, 901);
            Controls.Add(buttonOpenDatabasePath);
            Controls.Add(buttonOpenScriptPath);
            Controls.Add(buttonOpenBytesPath);
            Controls.Add(labelTables);
            Controls.Add(buttonImportDatabase);
            Controls.Add(buttonClearTables);
            Controls.Add(buttonSelectAllTables);
            Controls.Add(buttonBrowseDatabase);
            Controls.Add(textBoxDatabasePath);
            Controls.Add(labelDatabasePath);
            Controls.Add(buttonBrowseScriptPath);
            Controls.Add(textBoxScriptPath);
            Controls.Add(labelScriptPath);
            Controls.Add(buttonBrowseBytesPath);
            Controls.Add(textBoxBytesPath);
            Controls.Add(labelBytesPath);
            Controls.Add(label1);
            Controls.Add(listBox1);
            Controls.Add(button2);
            Controls.Add(button1);
            FormBorderStyle = FormBorderStyle.FixedSingle;
            Margin = new Padding(4);
            MaximizeBox = false;
            Name = "Form1";
            Text = "TDC";
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Button button1;
        private Button button2;
        private CheckedListBox listBox1;
        private Label label1;
        private Label labelBytesPath;
        private TextBox textBoxBytesPath;
        private Button buttonBrowseBytesPath;
        private Label labelScriptPath;
        private TextBox textBoxScriptPath;
        private Button buttonBrowseScriptPath;
        private Button buttonOpenBytesPath;
        private Button buttonOpenScriptPath;
        private Label labelDatabasePath;
        private TextBox textBoxDatabasePath;
        private Button buttonBrowseDatabase;
        private Button buttonOpenDatabasePath;
        private Button buttonSelectAllTables;
        private Button buttonClearTables;
        private Button buttonImportDatabase;
        private Label labelTables;
    }
}

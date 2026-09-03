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
            listBox1 = new ListBox();
            label1 = new Label();
            labelBytesPath = new Label();
            textBoxBytesPath = new TextBox();
            buttonBrowseBytesPath = new Button();
            labelScriptPath = new Label();
            textBoxScriptPath = new TextBox();
            buttonBrowseScriptPath = new Button();
            SuspendLayout();
            // 
            // button1
            // 
            button1.Location = new Point(13, 652);
            button1.Margin = new Padding(4);
            button1.Name = "button1";
            button1.Size = new Size(180, 72);
            button1.TabIndex = 0;
            button1.Text = "Refresh";
            button1.UseVisualStyleBackColor = true;
            button1.Click += OnBtn_Refresh;
            // 
            // button2
            // 
            button2.Location = new Point(201, 652);
            button2.Margin = new Padding(4);
            button2.Name = "button2";
            button2.Size = new Size(532, 72);
            button2.TabIndex = 1;
            button2.Text = "Confirm";
            button2.UseVisualStyleBackColor = true;
            button2.Click += OnBtn_Confirm;
            // 
            // listBox1
            // 
            listBox1.FormattingEnabled = true;
            listBox1.ItemHeight = 30;
            listBox1.Location = new Point(13, 12);
            listBox1.Name = "listBox1";
            listBox1.Size = new Size(720, 394);
            listBox1.TabIndex = 2;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Font = new Font("맑은 고딕", 11F);
            label1.Location = new Point(13, 608);
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
            textBoxBytesPath.Size = new Size(610, 35);
            textBoxBytesPath.TabIndex = 5;
            // 
            // buttonBrowseBytesPath
            // 
            buttonBrowseBytesPath.Location = new Point(631, 455);
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
            textBoxScriptPath.Size = new Size(610, 35);
            textBoxScriptPath.TabIndex = 8;
            // 
            // buttonBrowseScriptPath
            // 
            buttonBrowseScriptPath.Location = new Point(631, 538);
            buttonBrowseScriptPath.Name = "buttonBrowseScriptPath";
            buttonBrowseScriptPath.Size = new Size(102, 40);
            buttonBrowseScriptPath.TabIndex = 9;
            buttonBrowseScriptPath.Text = "찾아보기";
            buttonBrowseScriptPath.UseVisualStyleBackColor = true;
            buttonBrowseScriptPath.Click += OnBtn_BrowseScriptPath;
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(12F, 30F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(746, 737);
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
        private ListBox listBox1;
        private Label label1;
        private Label labelBytesPath;
        private TextBox textBoxBytesPath;
        private Button buttonBrowseBytesPath;
        private Label labelScriptPath;
        private TextBox textBoxScriptPath;
        private Button buttonBrowseScriptPath;
    }
}

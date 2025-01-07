namespace Hakbal
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            ScannerNameLable = new Label();
            ScannerMakeLable = new Label();
            ScannerSerialNumberLabel = new Label();
            BarcodeSampleNameLabel = new Label();
            LogFileLabel = new Label();
            BrowseFilesButton = new Button();
            ScannerNameTextBox = new TextBox();
            ScannerMakeTextBox = new TextBox();
            ScannerSerialNumberTextBox = new TextBox();
            BarcodeSampleNameTextBox = new TextBox();
            LogFilePathTextBox = new TextBox();
            CompileResultsButton = new Button();
            ErrorMessageLabel = new Label();
            ErrorMessageTextBox = new TextBox();
            ClearErrorMessageButton = new Button();
            ClearAllButton = new Button();
            AddLogToDataSetButton = new Button();
            label1 = new Label();
            LogsInDataSetLabel = new Label();
            label3 = new Label();
            DeleteLastAddedButton = new Button();
            DataSetListBox = new ListBox();
            AuthorLabel = new Label();
            ScorecardGroupNumberLabel = new Label();
            ScorecardGroupNumberComboBox = new ComboBox();
            HowToCompareResultsLabel = new Label();
            HowToCompareResultsComboBox = new ComboBox();
            SuspendLayout();
            // 
            // ScannerNameLable
            // 
            ScannerNameLable.AutoSize = true;
            ScannerNameLable.Location = new Point(12, 191);
            ScannerNameLable.Name = "ScannerNameLable";
            ScannerNameLable.Size = new Size(87, 15);
            ScannerNameLable.TabIndex = 0;
            ScannerNameLable.Text = "Scanner Name:";
            // 
            // ScannerMakeLable
            // 
            ScannerMakeLable.AutoSize = true;
            ScannerMakeLable.Location = new Point(12, 235);
            ScannerMakeLable.Name = "ScannerMakeLable";
            ScannerMakeLable.Size = new Size(84, 15);
            ScannerMakeLable.TabIndex = 1;
            ScannerMakeLable.Text = "Scanner Make:";
            // 
            // ScannerSerialNumberLabel
            // 
            ScannerSerialNumberLabel.AutoSize = true;
            ScannerSerialNumberLabel.Location = new Point(12, 280);
            ScannerSerialNumberLabel.Name = "ScannerSerialNumberLabel";
            ScannerSerialNumberLabel.Size = new Size(130, 15);
            ScannerSerialNumberLabel.TabIndex = 2;
            ScannerSerialNumberLabel.Text = "Scanner Serial Number:";
            // 
            // BarcodeSampleNameLabel
            // 
            BarcodeSampleNameLabel.AutoSize = true;
            BarcodeSampleNameLabel.Location = new Point(12, 323);
            BarcodeSampleNameLabel.Name = "BarcodeSampleNameLabel";
            BarcodeSampleNameLabel.Size = new Size(130, 15);
            BarcodeSampleNameLabel.TabIndex = 3;
            BarcodeSampleNameLabel.Text = "Barcode Sample Name:";
            // 
            // LogFileLabel
            // 
            LogFileLabel.AutoSize = true;
            LogFileLabel.Location = new Point(10, 147);
            LogFileLabel.Name = "LogFileLabel";
            LogFileLabel.Size = new Size(51, 15);
            LogFileLabel.TabIndex = 4;
            LogFileLabel.Text = "Log File:";
            // 
            // BrowseFilesButton
            // 
            BrowseFilesButton.Location = new Point(451, 165);
            BrowseFilesButton.Name = "BrowseFilesButton";
            BrowseFilesButton.Size = new Size(137, 23);
            BrowseFilesButton.TabIndex = 9;
            BrowseFilesButton.Text = "Browse Files";
            BrowseFilesButton.UseVisualStyleBackColor = true;
            BrowseFilesButton.Click += BrowseFilesButton_Click;
            // 
            // ScannerNameTextBox
            // 
            ScannerNameTextBox.Location = new Point(12, 209);
            ScannerNameTextBox.Name = "ScannerNameTextBox";
            ScannerNameTextBox.Size = new Size(576, 23);
            ScannerNameTextBox.TabIndex = 11;
            // 
            // ScannerMakeTextBox
            // 
            ScannerMakeTextBox.Location = new Point(12, 253);
            ScannerMakeTextBox.Name = "ScannerMakeTextBox";
            ScannerMakeTextBox.Size = new Size(576, 23);
            ScannerMakeTextBox.TabIndex = 12;
            // 
            // ScannerSerialNumberTextBox
            // 
            ScannerSerialNumberTextBox.Location = new Point(12, 298);
            ScannerSerialNumberTextBox.Name = "ScannerSerialNumberTextBox";
            ScannerSerialNumberTextBox.Size = new Size(576, 23);
            ScannerSerialNumberTextBox.TabIndex = 13;
            // 
            // BarcodeSampleNameTextBox
            // 
            BarcodeSampleNameTextBox.Location = new Point(12, 341);
            BarcodeSampleNameTextBox.Name = "BarcodeSampleNameTextBox";
            BarcodeSampleNameTextBox.Size = new Size(576, 23);
            BarcodeSampleNameTextBox.TabIndex = 14;
            // 
            // LogFilePathTextBox
            // 
            LogFilePathTextBox.Location = new Point(12, 165);
            LogFilePathTextBox.Name = "LogFilePathTextBox";
            LogFilePathTextBox.Size = new Size(434, 23);
            LogFilePathTextBox.TabIndex = 15;
            // 
            // CompileResultsButton
            // 
            CompileResultsButton.Location = new Point(12, 484);
            CompileResultsButton.Name = "CompileResultsButton";
            CompileResultsButton.Size = new Size(1051, 96);
            CompileResultsButton.TabIndex = 16;
            CompileResultsButton.Text = "Compile Summary Graph Data";
            CompileResultsButton.UseVisualStyleBackColor = true;
            CompileResultsButton.Click += CompileResultsButton_Click;
            // 
            // ErrorMessageLabel
            // 
            ErrorMessageLabel.AutoSize = true;
            ErrorMessageLabel.Location = new Point(12, 9);
            ErrorMessageLabel.Name = "ErrorMessageLabel";
            ErrorMessageLabel.Size = new Size(84, 15);
            ErrorMessageLabel.TabIndex = 17;
            ErrorMessageLabel.Text = "Error Message:";
            // 
            // ErrorMessageTextBox
            // 
            ErrorMessageTextBox.Cursor = Cursors.IBeam;
            ErrorMessageTextBox.Font = new Font("Courier New", 9F, FontStyle.Bold, GraphicsUnit.Point);
            ErrorMessageTextBox.ForeColor = Color.Red;
            ErrorMessageTextBox.Location = new Point(12, 27);
            ErrorMessageTextBox.Multiline = true;
            ErrorMessageTextBox.Name = "ErrorMessageTextBox";
            ErrorMessageTextBox.ReadOnly = true;
            ErrorMessageTextBox.Size = new Size(462, 112);
            ErrorMessageTextBox.TabIndex = 18;
            // 
            // ClearErrorMessageButton
            // 
            ClearErrorMessageButton.Location = new Point(480, 27);
            ClearErrorMessageButton.Name = "ClearErrorMessageButton";
            ClearErrorMessageButton.Size = new Size(104, 111);
            ClearErrorMessageButton.TabIndex = 19;
            ClearErrorMessageButton.Text = "Clear Error Message";
            ClearErrorMessageButton.UseVisualStyleBackColor = true;
            ClearErrorMessageButton.Click += ClearErrorMessageButton_Click;
            // 
            // ClearAllButton
            // 
            ClearAllButton.Location = new Point(429, 413);
            ClearAllButton.Name = "ClearAllButton";
            ClearAllButton.Size = new Size(156, 47);
            ClearAllButton.TabIndex = 20;
            ClearAllButton.Text = "Clear All";
            ClearAllButton.UseVisualStyleBackColor = true;
            ClearAllButton.Click += ClearAllButton_Click;
            // 
            // AddLogToDataSetButton
            // 
            AddLogToDataSetButton.Location = new Point(12, 413);
            AddLogToDataSetButton.Margin = new Padding(3, 2, 3, 2);
            AddLogToDataSetButton.Name = "AddLogToDataSetButton";
            AddLogToDataSetButton.Size = new Size(411, 47);
            AddLogToDataSetButton.TabIndex = 21;
            AddLogToDataSetButton.Text = "Add Log To Data Set";
            AddLogToDataSetButton.UseVisualStyleBackColor = true;
            AddLogToDataSetButton.Click += AddLogToDataSetButton_Click;
            // 
            // label1
            // 
            label1.BorderStyle = BorderStyle.Fixed3D;
            label1.Location = new Point(593, 9);
            label1.Name = "label1";
            label1.Size = new Size(2, 463);
            label1.TabIndex = 22;
            // 
            // LogsInDataSetLabel
            // 
            LogsInDataSetLabel.AutoSize = true;
            LogsInDataSetLabel.Location = new Point(606, 9);
            LogsInDataSetLabel.Name = "LogsInDataSetLabel";
            LogsInDataSetLabel.Size = new Size(94, 15);
            LogsInDataSetLabel.TabIndex = 23;
            LogsInDataSetLabel.Text = "Logs in Data Set:";
            // 
            // label3
            // 
            label3.BorderStyle = BorderStyle.Fixed3D;
            label3.Location = new Point(14, 472);
            label3.Name = "label3";
            label3.Size = new Size(1050, 2);
            label3.TabIndex = 24;
            // 
            // DeleteLastAddedButton
            // 
            DeleteLastAddedButton.Location = new Point(600, 413);
            DeleteLastAddedButton.Margin = new Padding(3, 2, 3, 2);
            DeleteLastAddedButton.Name = "DeleteLastAddedButton";
            DeleteLastAddedButton.Size = new Size(463, 47);
            DeleteLastAddedButton.TabIndex = 26;
            DeleteLastAddedButton.Text = "Delete Last Added";
            DeleteLastAddedButton.UseVisualStyleBackColor = true;
            DeleteLastAddedButton.Click += DeleteLastAddedButton_Click;
            // 
            // DataSetListBox
            // 
            DataSetListBox.FormattingEnabled = true;
            DataSetListBox.HorizontalScrollbar = true;
            DataSetListBox.ItemHeight = 15;
            DataSetListBox.Location = new Point(606, 27);
            DataSetListBox.Margin = new Padding(3, 2, 3, 2);
            DataSetListBox.Name = "DataSetListBox";
            DataSetListBox.ScrollAlwaysVisible = true;
            DataSetListBox.Size = new Size(458, 379);
            DataSetListBox.TabIndex = 27;
            // 
            // AuthorLabel
            // 
            AuthorLabel.AutoSize = true;
            AuthorLabel.Location = new Point(12, 588);
            AuthorLabel.Name = "AuthorLabel";
            AuthorLabel.Size = new Size(438, 15);
            AuthorLabel.TabIndex = 28;
            AuthorLabel.Text = "Created by Daniel Calabrese (DC1923). Please reach out with any questions/issues.";
            // 
            // ScorecardGroupNumberLabel
            // 
            ScorecardGroupNumberLabel.AutoSize = true;
            ScorecardGroupNumberLabel.Location = new Point(12, 367);
            ScorecardGroupNumberLabel.Name = "ScorecardGroupNumberLabel";
            ScorecardGroupNumberLabel.Size = new Size(145, 15);
            ScorecardGroupNumberLabel.TabIndex = 30;
            ScorecardGroupNumberLabel.Text = "Scorecard Group Number:";
            // 
            // ScorecardGroupNumberComboBox
            // 
            ScorecardGroupNumberComboBox.DropDownStyle = ComboBoxStyle.DropDownList;
            ScorecardGroupNumberComboBox.FormattingEnabled = true;
            ScorecardGroupNumberComboBox.Items.AddRange(new object[] { "1", "2", "3", "4", "5" });
            ScorecardGroupNumberComboBox.Location = new Point(14, 385);
            ScorecardGroupNumberComboBox.Name = "ScorecardGroupNumberComboBox";
            ScorecardGroupNumberComboBox.Size = new Size(130, 23);
            ScorecardGroupNumberComboBox.TabIndex = 31;
            // 
            // HowToCompareResultsLabel
            // 
            HowToCompareResultsLabel.AutoSize = true;
            HowToCompareResultsLabel.Location = new Point(226, 367);
            HowToCompareResultsLabel.Name = "HowToCompareResultsLabel";
            HowToCompareResultsLabel.Size = new Size(203, 15);
            HowToCompareResultsLabel.TabIndex = 32;
            HowToCompareResultsLabel.Text = "How To Compare (Highlight) Results:";
            // 
            // HowToCompareResultsComboBox
            // 
            HowToCompareResultsComboBox.DropDownStyle = ComboBoxStyle.DropDownList;
            HowToCompareResultsComboBox.FormattingEnabled = true;
            HowToCompareResultsComboBox.Items.AddRange(new object[] { "Do Not Compare (No Highlighting)", "All By Column (Top to Bottom)", "All By Row (Left to Right)", "Decode Time Average Only by Column (Top to Bottom)", "Decode Time Average Only by Row (Left to Right)" });
            HowToCompareResultsComboBox.Location = new Point(226, 385);
            HowToCompareResultsComboBox.Margin = new Padding(3, 2, 3, 2);
            HowToCompareResultsComboBox.Name = "HowToCompareResultsComboBox";
            HowToCompareResultsComboBox.Size = new Size(358, 23);
            HowToCompareResultsComboBox.TabIndex = 33;
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(1075, 614);
            Controls.Add(HowToCompareResultsComboBox);
            Controls.Add(HowToCompareResultsLabel);
            Controls.Add(ScorecardGroupNumberComboBox);
            Controls.Add(ScorecardGroupNumberLabel);
            Controls.Add(AuthorLabel);
            Controls.Add(DataSetListBox);
            Controls.Add(DeleteLastAddedButton);
            Controls.Add(label3);
            Controls.Add(LogsInDataSetLabel);
            Controls.Add(label1);
            Controls.Add(AddLogToDataSetButton);
            Controls.Add(ClearAllButton);
            Controls.Add(ClearErrorMessageButton);
            Controls.Add(ErrorMessageTextBox);
            Controls.Add(ErrorMessageLabel);
            Controls.Add(ScannerNameTextBox);
            Controls.Add(CompileResultsButton);
            Controls.Add(LogFilePathTextBox);
            Controls.Add(BarcodeSampleNameTextBox);
            Controls.Add(ScannerSerialNumberTextBox);
            Controls.Add(ScannerMakeTextBox);
            Controls.Add(BrowseFilesButton);
            Controls.Add(LogFileLabel);
            Controls.Add(BarcodeSampleNameLabel);
            Controls.Add(ScannerSerialNumberLabel);
            Controls.Add(ScannerMakeLable);
            Controls.Add(ScannerNameLable);
            Icon = (Icon)resources.GetObject("$this.Icon");
            MaximizeBox = false;
            Name = "Form1";
            Text = "Hakbal 3.0";
            Load += Form1_Load;
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Label ScannerNameLable;
        private Label ScannerMakeLable;
        private Label ScannerSerialNumberLabel;
        private Label BarcodeSampleNameLabel;
        private Label LogFileLabel;
        private Button BrowseFilesButton;
        private TextBox ScannerNameTextBox;
        private TextBox ScannerMakeTextBox;
        private TextBox ScannerSerialNumberTextBox;
        private TextBox BarcodeSampleNameTextBox;
        private TextBox LogFilePathTextBox;
        private Button CompileResultsButton;
        private Label ErrorMessageLabel;
        private TextBox ErrorMessageTextBox;
        private Button ClearErrorMessageButton;
        private Button ClearAllButton;
        private Button AddLogToDataSetButton;
        private Label label1;
        private Label LogsInDataSetLabel;
        private Label label3;
        private Button DeleteLastAddedButton;
        private ListBox DataSetListBox;
        private Label AuthorLabel;
        private Label ScorecardGroupNumberLabel;
        public ComboBox ScorecardGroupNumberComboBox;
        private Label HowToCompareResultsLabel;
        public ComboBox HowToCompareResultsComboBox;
        private ComboBox comboBox1;
    }
}
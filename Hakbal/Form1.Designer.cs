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
            SuspendLayout();
            // 
            // ScannerNameLable
            // 
            ScannerNameLable.AutoSize = true;
            ScannerNameLable.Location = new Point(14, 188);
            ScannerNameLable.Name = "ScannerNameLable";
            ScannerNameLable.Size = new Size(108, 20);
            ScannerNameLable.TabIndex = 0;
            ScannerNameLable.Text = "Scanner Name:";
            // 
            // ScannerMakeLable
            // 
            ScannerMakeLable.AutoSize = true;
            ScannerMakeLable.Location = new Point(14, 247);
            ScannerMakeLable.Name = "ScannerMakeLable";
            ScannerMakeLable.Size = new Size(104, 20);
            ScannerMakeLable.TabIndex = 1;
            ScannerMakeLable.Text = "Scanner Make:";
            // 
            // ScannerSerialNumberLabel
            // 
            ScannerSerialNumberLabel.AutoSize = true;
            ScannerSerialNumberLabel.Location = new Point(14, 307);
            ScannerSerialNumberLabel.Name = "ScannerSerialNumberLabel";
            ScannerSerialNumberLabel.Size = new Size(163, 20);
            ScannerSerialNumberLabel.TabIndex = 2;
            ScannerSerialNumberLabel.Text = "Scanner Serial Number:";
            // 
            // BarcodeSampleNameLabel
            // 
            BarcodeSampleNameLabel.AutoSize = true;
            BarcodeSampleNameLabel.Location = new Point(14, 364);
            BarcodeSampleNameLabel.Name = "BarcodeSampleNameLabel";
            BarcodeSampleNameLabel.Size = new Size(165, 20);
            BarcodeSampleNameLabel.TabIndex = 3;
            BarcodeSampleNameLabel.Text = "Barcode Sample Name:";
            // 
            // LogFileLabel
            // 
            LogFileLabel.AutoSize = true;
            LogFileLabel.Location = new Point(18, 423);
            LogFileLabel.Name = "LogFileLabel";
            LogFileLabel.Size = new Size(64, 20);
            LogFileLabel.TabIndex = 4;
            LogFileLabel.Text = "Log File:";
            // 
            // BrowseFilesButton
            // 
            BrowseFilesButton.Location = new Point(515, 443);
            BrowseFilesButton.Margin = new Padding(3, 4, 3, 4);
            BrowseFilesButton.Name = "BrowseFilesButton";
            BrowseFilesButton.Size = new Size(157, 31);
            BrowseFilesButton.TabIndex = 9;
            BrowseFilesButton.Text = "Browse Files";
            BrowseFilesButton.UseVisualStyleBackColor = true;
            BrowseFilesButton.Click += BrowseFilesButton_Click;
            // 
            // ScannerNameTextBox
            // 
            ScannerNameTextBox.Location = new Point(14, 212);
            ScannerNameTextBox.Margin = new Padding(3, 4, 3, 4);
            ScannerNameTextBox.Name = "ScannerNameTextBox";
            ScannerNameTextBox.Size = new Size(658, 27);
            ScannerNameTextBox.TabIndex = 11;
            // 
            // ScannerMakeTextBox
            // 
            ScannerMakeTextBox.Location = new Point(14, 271);
            ScannerMakeTextBox.Margin = new Padding(3, 4, 3, 4);
            ScannerMakeTextBox.Name = "ScannerMakeTextBox";
            ScannerMakeTextBox.Size = new Size(658, 27);
            ScannerMakeTextBox.TabIndex = 12;
            // 
            // ScannerSerialNumberTextBox
            // 
            ScannerSerialNumberTextBox.Location = new Point(14, 331);
            ScannerSerialNumberTextBox.Margin = new Padding(3, 4, 3, 4);
            ScannerSerialNumberTextBox.Name = "ScannerSerialNumberTextBox";
            ScannerSerialNumberTextBox.Size = new Size(658, 27);
            ScannerSerialNumberTextBox.TabIndex = 13;
            // 
            // BarcodeSampleNameTextBox
            // 
            BarcodeSampleNameTextBox.Location = new Point(14, 388);
            BarcodeSampleNameTextBox.Margin = new Padding(3, 4, 3, 4);
            BarcodeSampleNameTextBox.Name = "BarcodeSampleNameTextBox";
            BarcodeSampleNameTextBox.Size = new Size(658, 27);
            BarcodeSampleNameTextBox.TabIndex = 14;
            // 
            // LogFilePathTextBox
            // 
            LogFilePathTextBox.Location = new Point(14, 447);
            LogFilePathTextBox.Margin = new Padding(3, 4, 3, 4);
            LogFilePathTextBox.Name = "LogFilePathTextBox";
            LogFilePathTextBox.Size = new Size(495, 27);
            LogFilePathTextBox.TabIndex = 15;
            // 
            // CompileResultsButton
            // 
            CompileResultsButton.Location = new Point(14, 645);
            CompileResultsButton.Margin = new Padding(3, 4, 3, 4);
            CompileResultsButton.Name = "CompileResultsButton";
            CompileResultsButton.Size = new Size(1047, 128);
            CompileResultsButton.TabIndex = 16;
            CompileResultsButton.Text = "Compile Summary Graph Data";
            CompileResultsButton.UseVisualStyleBackColor = true;
            CompileResultsButton.Click += CompileResultsButton_Click;
            // 
            // ErrorMessageLabel
            // 
            ErrorMessageLabel.AutoSize = true;
            ErrorMessageLabel.Location = new Point(14, 12);
            ErrorMessageLabel.Name = "ErrorMessageLabel";
            ErrorMessageLabel.Size = new Size(106, 20);
            ErrorMessageLabel.TabIndex = 17;
            ErrorMessageLabel.Text = "Error Message:";
            // 
            // ErrorMessageTextBox
            // 
            ErrorMessageTextBox.Cursor = Cursors.IBeam;
            ErrorMessageTextBox.Font = new Font("Courier New", 9F, FontStyle.Bold, GraphicsUnit.Point);
            ErrorMessageTextBox.ForeColor = Color.Red;
            ErrorMessageTextBox.Location = new Point(14, 36);
            ErrorMessageTextBox.Margin = new Padding(3, 4, 3, 4);
            ErrorMessageTextBox.Multiline = true;
            ErrorMessageTextBox.Name = "ErrorMessageTextBox";
            ErrorMessageTextBox.ReadOnly = true;
            ErrorMessageTextBox.Size = new Size(527, 148);
            ErrorMessageTextBox.TabIndex = 18;
            // 
            // ClearErrorMessageButton
            // 
            ClearErrorMessageButton.Location = new Point(549, 36);
            ClearErrorMessageButton.Margin = new Padding(3, 4, 3, 4);
            ClearErrorMessageButton.Name = "ClearErrorMessageButton";
            ClearErrorMessageButton.Size = new Size(119, 148);
            ClearErrorMessageButton.TabIndex = 19;
            ClearErrorMessageButton.Text = "Clear Error Message";
            ClearErrorMessageButton.UseVisualStyleBackColor = true;
            ClearErrorMessageButton.Click += ClearErrorMessageButton_Click;
            // 
            // ClearAllButton
            // 
            ClearAllButton.Location = new Point(550, 485);
            ClearAllButton.Margin = new Padding(3, 4, 3, 4);
            ClearAllButton.Name = "ClearAllButton";
            ClearAllButton.Size = new Size(119, 128);
            ClearAllButton.TabIndex = 20;
            ClearAllButton.Text = "Clear All";
            ClearAllButton.UseVisualStyleBackColor = true;
            ClearAllButton.Click += ClearAllButton_Click;
            // 
            // AddLogToDataSetButton
            // 
            AddLogToDataSetButton.Location = new Point(14, 485);
            AddLogToDataSetButton.Name = "AddLogToDataSetButton";
            AddLogToDataSetButton.Size = new Size(530, 128);
            AddLogToDataSetButton.TabIndex = 21;
            AddLogToDataSetButton.Text = "Add Log To Data Set";
            AddLogToDataSetButton.UseVisualStyleBackColor = true;
            AddLogToDataSetButton.Click += AddLogToDataSetButton_Click;
            // 
            // label1
            // 
            label1.BorderStyle = BorderStyle.Fixed3D;
            label1.Location = new Point(678, 12);
            label1.Name = "label1";
            label1.Size = new Size(2, 617);
            label1.TabIndex = 22;
            // 
            // LogsInDataSetLabel
            // 
            LogsInDataSetLabel.AutoSize = true;
            LogsInDataSetLabel.Location = new Point(693, 12);
            LogsInDataSetLabel.Name = "LogsInDataSetLabel";
            LogsInDataSetLabel.Size = new Size(120, 20);
            LogsInDataSetLabel.TabIndex = 23;
            LogsInDataSetLabel.Text = "Logs in Data Set:";
            // 
            // label3
            // 
            label3.BorderStyle = BorderStyle.Fixed3D;
            label3.Location = new Point(16, 629);
            label3.Name = "label3";
            label3.Size = new Size(1045, 3);
            label3.TabIndex = 24;
            // 
            // DeleteLastAddedButton
            // 
            DeleteLastAddedButton.Location = new Point(686, 485);
            DeleteLastAddedButton.Name = "DeleteLastAddedButton";
            DeleteLastAddedButton.Size = new Size(375, 127);
            DeleteLastAddedButton.TabIndex = 26;
            DeleteLastAddedButton.Text = "Delete Last Added";
            DeleteLastAddedButton.UseVisualStyleBackColor = true;
            DeleteLastAddedButton.Click += DeleteLastAddedButton_Click;
            // 
            // DataSetListBox
            // 
            DataSetListBox.FormattingEnabled = true;
            DataSetListBox.HorizontalScrollbar = true;
            DataSetListBox.ItemHeight = 20;
            DataSetListBox.Location = new Point(693, 36);
            DataSetListBox.Name = "DataSetListBox";
            DataSetListBox.ScrollAlwaysVisible = true;
            DataSetListBox.Size = new Size(368, 444);
            DataSetListBox.TabIndex = 27;
            // 
            // AuthorLabel
            // 
            AuthorLabel.AutoSize = true;
            AuthorLabel.Location = new Point(14, 784);
            AuthorLabel.Name = "AuthorLabel";
            AuthorLabel.Size = new Size(551, 20);
            AuthorLabel.TabIndex = 28;
            AuthorLabel.Text = "Created by Daniel Calabrese (DC1923). Please reach out with any questions/issues.";
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(8F, 20F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(1074, 813);
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
            Margin = new Padding(3, 4, 3, 4);
            MaximizeBox = false;
            Name = "Form1";
            Text = "Hakbal 2.0";
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
    }
}
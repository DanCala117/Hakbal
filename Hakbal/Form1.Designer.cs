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
            ScannerNameLable = new Label();
            ScannerMakeLable = new Label();
            ScannerSerialNumberLabel = new Label();
            BarcodeSampleNameLabel = new Label();
            LogFileLabel = new Label();
            SetScannerNameButton = new Button();
            SetScannerMakeButton = new Button();
            SetScannerSerialNumberButton = new Button();
            SetBarcodeSampleNameButton = new Button();
            BrowseFilesButton = new Button();
            SetLogPathButton = new Button();
            ScannerNameTextBox = new TextBox();
            ScannerMakeTextBox = new TextBox();
            ScannerSerialNumberTextBox = new TextBox();
            BarcodeSampleNameTextBox = new TextBox();
            LogFilePathTextBox = new TextBox();
            CompileResultsButton = new Button();
            ErrorMessageLabel = new Label();
            ErrorMessageTextBox = new TextBox();
            SuspendLayout();
            // 
            // ScannerNameLable
            // 
            ScannerNameLable.AutoSize = true;
            ScannerNameLable.Location = new Point(12, 115);
            ScannerNameLable.Name = "ScannerNameLable";
            ScannerNameLable.Size = new Size(84, 15);
            ScannerNameLable.TabIndex = 0;
            ScannerNameLable.Text = "Scanner Name";
            // 
            // ScannerMakeLable
            // 
            ScannerMakeLable.AutoSize = true;
            ScannerMakeLable.Location = new Point(12, 159);
            ScannerMakeLable.Name = "ScannerMakeLable";
            ScannerMakeLable.Size = new Size(81, 15);
            ScannerMakeLable.TabIndex = 1;
            ScannerMakeLable.Text = "Scanner Make";
            // 
            // ScannerSerialNumberLabel
            // 
            ScannerSerialNumberLabel.AutoSize = true;
            ScannerSerialNumberLabel.Location = new Point(12, 203);
            ScannerSerialNumberLabel.Name = "ScannerSerialNumberLabel";
            ScannerSerialNumberLabel.Size = new Size(127, 15);
            ScannerSerialNumberLabel.TabIndex = 2;
            ScannerSerialNumberLabel.Text = "Scanner Serial Number";
            // 
            // BarcodeSampleNameLabel
            // 
            BarcodeSampleNameLabel.AutoSize = true;
            BarcodeSampleNameLabel.Location = new Point(12, 247);
            BarcodeSampleNameLabel.Name = "BarcodeSampleNameLabel";
            BarcodeSampleNameLabel.Size = new Size(127, 15);
            BarcodeSampleNameLabel.TabIndex = 3;
            BarcodeSampleNameLabel.Text = "Barcode Sample Name";
            // 
            // LogFileLabel
            // 
            LogFileLabel.AutoSize = true;
            LogFileLabel.Location = new Point(12, 295);
            LogFileLabel.Name = "LogFileLabel";
            LogFileLabel.Size = new Size(48, 15);
            LogFileLabel.TabIndex = 4;
            LogFileLabel.Text = "Log File";
            // 
            // SetScannerNameButton
            // 
            SetScannerNameButton.Location = new Point(364, 133);
            SetScannerNameButton.Name = "SetScannerNameButton";
            SetScannerNameButton.Size = new Size(220, 23);
            SetScannerNameButton.TabIndex = 5;
            SetScannerNameButton.Text = "Set Scanner Name";
            SetScannerNameButton.UseVisualStyleBackColor = true;
            SetScannerNameButton.Click += SetScannerNameButton_Click;
            // 
            // SetScannerMakeButton
            // 
            SetScannerMakeButton.Location = new Point(364, 177);
            SetScannerMakeButton.Name = "SetScannerMakeButton";
            SetScannerMakeButton.Size = new Size(220, 23);
            SetScannerMakeButton.TabIndex = 6;
            SetScannerMakeButton.Text = "Set Scanner Make";
            SetScannerMakeButton.UseVisualStyleBackColor = true;
            SetScannerMakeButton.Click += SetScannerMakeButton_Click;
            // 
            // SetScannerSerialNumberButton
            // 
            SetScannerSerialNumberButton.Location = new Point(364, 221);
            SetScannerSerialNumberButton.Name = "SetScannerSerialNumberButton";
            SetScannerSerialNumberButton.Size = new Size(220, 23);
            SetScannerSerialNumberButton.TabIndex = 7;
            SetScannerSerialNumberButton.Text = "Set Scanner Serial Number";
            SetScannerSerialNumberButton.UseVisualStyleBackColor = true;
            SetScannerSerialNumberButton.Click += SetScannerSerialNumberButton_Click;
            // 
            // SetBarcodeSampleNameButton
            // 
            SetBarcodeSampleNameButton.Location = new Point(364, 265);
            SetBarcodeSampleNameButton.Name = "SetBarcodeSampleNameButton";
            SetBarcodeSampleNameButton.Size = new Size(220, 23);
            SetBarcodeSampleNameButton.TabIndex = 8;
            SetBarcodeSampleNameButton.Text = "Set Barcode Sample Name";
            SetBarcodeSampleNameButton.UseVisualStyleBackColor = true;
            SetBarcodeSampleNameButton.Click += SetBarcodeSampleNameButton_Click;
            // 
            // BrowseFilesButton
            // 
            BrowseFilesButton.Location = new Point(301, 313);
            BrowseFilesButton.Name = "BrowseFilesButton";
            BrowseFilesButton.Size = new Size(137, 23);
            BrowseFilesButton.TabIndex = 9;
            BrowseFilesButton.Text = "Browse Files";
            BrowseFilesButton.UseVisualStyleBackColor = true;
            BrowseFilesButton.Click += BrowseFilesButton_Click;
            // 
            // SetLogPathButton
            // 
            SetLogPathButton.Location = new Point(444, 313);
            SetLogPathButton.Name = "SetLogPathButton";
            SetLogPathButton.Size = new Size(137, 23);
            SetLogPathButton.TabIndex = 10;
            SetLogPathButton.Text = "Set Log File Path";
            SetLogPathButton.UseVisualStyleBackColor = true;
            SetLogPathButton.Click += SetLogPathButton_Click;
            // 
            // ScannerNameTextBox
            // 
            ScannerNameTextBox.Location = new Point(12, 133);
            ScannerNameTextBox.Name = "ScannerNameTextBox";
            ScannerNameTextBox.Size = new Size(346, 23);
            ScannerNameTextBox.TabIndex = 11;
            // 
            // ScannerMakeTextBox
            // 
            ScannerMakeTextBox.Location = new Point(12, 177);
            ScannerMakeTextBox.Name = "ScannerMakeTextBox";
            ScannerMakeTextBox.Size = new Size(346, 23);
            ScannerMakeTextBox.TabIndex = 12;
            // 
            // ScannerSerialNumberTextBox
            // 
            ScannerSerialNumberTextBox.Location = new Point(12, 221);
            ScannerSerialNumberTextBox.Name = "ScannerSerialNumberTextBox";
            ScannerSerialNumberTextBox.Size = new Size(346, 23);
            ScannerSerialNumberTextBox.TabIndex = 13;
            // 
            // BarcodeSampleNameTextBox
            // 
            BarcodeSampleNameTextBox.Location = new Point(12, 265);
            BarcodeSampleNameTextBox.Name = "BarcodeSampleNameTextBox";
            BarcodeSampleNameTextBox.Size = new Size(346, 23);
            BarcodeSampleNameTextBox.TabIndex = 14;
            // 
            // LogFilePathTextBox
            // 
            LogFilePathTextBox.Location = new Point(12, 313);
            LogFilePathTextBox.Name = "LogFilePathTextBox";
            LogFilePathTextBox.Size = new Size(283, 23);
            LogFilePathTextBox.TabIndex = 15;
            // 
            // CompileResultsButton
            // 
            CompileResultsButton.Location = new Point(12, 342);
            CompileResultsButton.Name = "CompileResultsButton";
            CompileResultsButton.Size = new Size(572, 96);
            CompileResultsButton.TabIndex = 16;
            CompileResultsButton.Text = "Compile SummaryGraphData";
            CompileResultsButton.UseVisualStyleBackColor = true;
            CompileResultsButton.Click += CompileResultsButton_Click;
            // 
            // ErrorMessageLabel
            // 
            ErrorMessageLabel.AutoSize = true;
            ErrorMessageLabel.Location = new Point(12, 9);
            ErrorMessageLabel.Name = "ErrorMessageLabel";
            ErrorMessageLabel.Size = new Size(81, 15);
            ErrorMessageLabel.TabIndex = 17;
            ErrorMessageLabel.Text = "Error Message";
            // 
            // ErrorMessageTextBox
            // 
            ErrorMessageTextBox.Cursor = Cursors.IBeam;
            ErrorMessageTextBox.Font = new Font("Courier New", 9F, FontStyle.Bold, GraphicsUnit.Point);
            ErrorMessageTextBox.Location = new Point(12, 27);
            ErrorMessageTextBox.Multiline = true;
            ErrorMessageTextBox.Name = "ErrorMessageTextBox";
            ErrorMessageTextBox.ReadOnly = true;
            ErrorMessageTextBox.Size = new Size(572, 65);
            ErrorMessageTextBox.TabIndex = 18;
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(593, 450);
            Controls.Add(ErrorMessageTextBox);
            Controls.Add(ErrorMessageLabel);
            Controls.Add(ScannerNameTextBox);
            Controls.Add(CompileResultsButton);
            Controls.Add(LogFilePathTextBox);
            Controls.Add(BarcodeSampleNameTextBox);
            Controls.Add(ScannerSerialNumberTextBox);
            Controls.Add(ScannerMakeTextBox);
            Controls.Add(SetLogPathButton);
            Controls.Add(BrowseFilesButton);
            Controls.Add(SetBarcodeSampleNameButton);
            Controls.Add(SetScannerSerialNumberButton);
            Controls.Add(SetScannerMakeButton);
            Controls.Add(SetScannerNameButton);
            Controls.Add(LogFileLabel);
            Controls.Add(BarcodeSampleNameLabel);
            Controls.Add(ScannerSerialNumberLabel);
            Controls.Add(ScannerMakeLable);
            Controls.Add(ScannerNameLable);
            Name = "Form1";
            Text = "Hakbal";
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Label ScannerNameLable;
        private Label ScannerMakeLable;
        private Label ScannerSerialNumberLabel;
        private Label BarcodeSampleNameLabel;
        private Label LogFileLabel;
        private Button SetScannerNameButton;
        private Button SetScannerMakeButton;
        private Button SetScannerSerialNumberButton;
        private Button SetBarcodeSampleNameButton;
        private Button BrowseFilesButton;
        private Button SetLogPathButton;
        private TextBox ScannerNameTextBox;
        private TextBox ScannerMakeTextBox;
        private TextBox ScannerSerialNumberTextBox;
        private TextBox BarcodeSampleNameTextBox;
        private TextBox LogFilePathTextBox;
        private Button CompileResultsButton;
        private Label ErrorMessageLabel;
        private TextBox ErrorMessageTextBox;
    }
}
using Microsoft.VisualBasic.Logging;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace Hakbal
{
    public partial class Form1 : Form
    {
        //Create an instance of the TargetLog class
        //TO BE DELETED
        //TargetLog TLog = new TargetLog();

        //list to contain each instance of the target log class
        //needed to select multiple files
        List<TargetLog> DataSet = new List<TargetLog>();

        /// <summary>
        /// loads GUI
        /// </summary>
        public Form1()
        {
            InitializeComponent();

            //Window will no longer be resizable
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
        }

        #region GUI_Elements
        /// <summary>
        /// What happens when the form initiallylaods
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Form1_Load(object sender, EventArgs e)
        {
            //GUI ToolTips
            System.Windows.Forms.ToolTip MyToolTip = new System.Windows.Forms.ToolTip();
            MyToolTip.AutoPopDelay = 5000;
            MyToolTip.InitialDelay = 500;
            MyToolTip.ReshowDelay = 500;
            MyToolTip.ShowAlways = true;
            MyToolTip.SetToolTip(this.ScannerNameTextBox, "Example: \"DS8178_001-R00_Defaults\".");
            MyToolTip.SetToolTip(this.ScannerMakeTextBox, "Example: \"Zebra\".");
            MyToolTip.SetToolTip(this.ScannerSerialNumberTextBox, "Example: \"01234567890\".");
            MyToolTip.SetToolTip(this.BarcodeSampleNameTextBox, "Example: \"Paper_QR10m\".");
            MyToolTip.SetToolTip(this.LogFilePathTextBox, "Please select the correct file path.");
            MyToolTip.SetToolTip(this.AddLogToDataSetButton, "Adds the data above to the data set to compile. Please add one or more logs.");
            MyToolTip.SetToolTip(this.ClearAllButton, "Clears all of the data above.");
            MyToolTip.SetToolTip(this.DataSetListBox, "List of all logs to be compiled.");
            MyToolTip.SetToolTip(this.DeleteLastAddedButton, "Deletes only the last added log from the data set.");
            MyToolTip.SetToolTip(this.CompileResultsButton, "This will create a folder on your desktop named \"Hakbal_Files\" and save the compiled results there.");
        }

        /// <summary>
        /// Opens up the file explorer for the user to select their log
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BrowseFilesButton_Click(object sender, EventArgs e)
        {
            //dialog box parameters
            OpenFileDialog LogSelectDialog = new OpenFileDialog();
            LogSelectDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            LogSelectDialog.Filter = "log files (*.log)|*.log|All files (*.*)|*.*";
            LogSelectDialog.Multiselect = false;
            LogSelectDialog.Title = "Log Selection";

            DialogResult SelectedFiles = LogSelectDialog.ShowDialog();

            //gets chosen log
            try
            {
                if (SelectedFiles == DialogResult.OK)
                {
                    LogFilePathTextBox.Text = string.Join(Environment.NewLine, LogSelectDialog.FileNames);
                }
                else
                {
                    //only if the dialog box is opened and no log is selected
                    ErrorMessageTextBox.Text = "You Closed The Dialog Box Without Selecting Any Files.";
                }
            }
            catch (Exception)
            {
                ErrorMessageTextBox.Text = "Error Finding Log File.";
            }
        }

        /// <summary>
        /// Adds most recent log to the list of logs to compile
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void AddLogToDataSetButton_Click(object sender, EventArgs e)
        {
            try
            {
                //gets user givin data and adds it to the data set to compile
                TargetLog TLog = new TargetLog();

                TLog.ScannerName = ScannerNameTextBox.Text;
                TLog.ScannerMake = ScannerMakeTextBox.Text;
                TLog.ScannerSerialNumber = ScannerSerialNumberTextBox.Text;
                TLog.BarcodeSampleName = BarcodeSampleNameTextBox.Text;
                TLog.LogFilePath = LogFilePathTextBox.Text;

                DataSet.Add(TLog);

                //prints data set to text box
                if (TLog.ToString() != "Log: , , ")
                {
                    DataSetListBox.Items.Add(TLog.ToString());
                }
                else
                {
                    ErrorMessageTextBox.Text = "No data has been entered.";
                }
            }
            catch (Exception)
            {
                ErrorMessageTextBox.Text = "There was an issue adding the data to the data set.";
            }
        }

        /// <summary>
        /// Clears all text boxes when clicked
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ClearAllButton_Click(object sender, EventArgs e)
        {
            try
            {
                //emptys the text boxes
                ScannerNameTextBox.Text = String.Empty;
                ScannerMakeTextBox.Text = String.Empty;
                ScannerSerialNumberTextBox.Text = String.Empty;
                BarcodeSampleNameTextBox.Text = String.Empty;
                LogFilePathTextBox.Text = String.Empty;
            }
            catch (Exception)
            {
                ErrorMessageTextBox.Text = "There was an issue clearing all text boxes.";
            }
        }

        /// <summary>
        /// Deletes the last entered log from the data set on button click
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void DeleteLastAddedButton_Click(object sender, EventArgs e)
        {
            try
            {
                //gets rid of the last entry in the data set
                DataSet.RemoveAt(DataSet.Count - 1);

                //prints updated data set
                DataSetListBox.Items.Clear();
                foreach (object log in DataSet)
                {
                    DataSetListBox.Items.Add(log.ToString());
                }
            }
            catch (Exception)
            {

                ErrorMessageTextBox.Text = "There are no more entrys to delete.";
            }
        }

        /// <summary>
        /// Clears the error message text box when clicked
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ClearErrorMessageButton_Click(object sender, EventArgs e)
        {
            //emptys text box
            ErrorMessageTextBox.Text = String.Empty;
        }

        /// <summary>
        /// Compiles the selected log
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CompileResultsButton_Click(object sender, EventArgs e)
        {
            // If you use EPPlus in a noncommercial context
            // according to the Polyform Noncommercial license:
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var excelPackage = new ExcelPackage())
            {
                //creates the Raw Results tab for each log
                try
                {
                    foreach (TargetLog log in DataSet)
                    {
                        //creates the tab for each log in the data set
                        var RawResultsSheet = excelPackage.Workbook.Worksheets.Add("RawResults_" + log.ScannerName + "_" + log.BarcodeSampleName);

                        //Tab header with scanner + barcode info 
                        RawResultsSheet.Cells["A1"].Value = "Scanner Name: " + log.ScannerName;
                        RawResultsSheet.Cells["A1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        RawResultsSheet.Cells["A1"].Style.Fill.BackgroundColor.SetColor(Color.Orange);

                        RawResultsSheet.Cells["A2"].Value = "Scanner Make: " + log.ScannerMake;
                        RawResultsSheet.Cells["A2"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        RawResultsSheet.Cells["A2"].Style.Fill.BackgroundColor.SetColor(Color.Orange);

                        RawResultsSheet.Cells["A3"].Value = "Scanner Serial Number: " + log.ScannerSerialNumber;
                        RawResultsSheet.Cells["A3"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        RawResultsSheet.Cells["A3"].Style.Fill.BackgroundColor.SetColor(Color.Orange);

                        RawResultsSheet.Cells["A4"].Value = "Barcode Sample Name: " + log.BarcodeSampleName;
                        RawResultsSheet.Cells["A4"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        RawResultsSheet.Cells["A4"].Style.Fill.BackgroundColor.SetColor(Color.Orange);

                        RawResultsSheet.Cells["A1:A4"].AutoFitColumns();

                        //raw data table header
                        var LogFileNameGraphHeader = RawResultsSheet.Cells["B5:M5"];
                        LogFileNameGraphHeader.Merge = true;
                        LogFileNameGraphHeader.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        LogFileNameGraphHeader.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        RawResultsSheet.Cells["B5"].Value = GetFileName(LogFilePathTextBox.Text.ToString()); //TODO FIX THIS LINE!!!!!!!!!!!!!!!!!!!!!!!!!

                        RawResultsSheet.Cells["B6"].Value = "Distance (In)";
                        RawResultsSheet.Cells["B6"].Style.Font.Bold = true;
                        RawResultsSheet.Cells["B6"].AutoFitColumns();

                        RawResultsSheet.Cells["C6"].Value = "Decode Times (MS)";
                        RawResultsSheet.Cells["C6"].Style.Font.Bold = true;

                        RawResultsSheet.Cells["M6"].Value = "Average (MS)";
                        RawResultsSheet.Cells["M6"].Style.Font.Bold = true;
                        RawResultsSheet.Cells["M6"].AutoFitColumns();

                        RawResultsSheet.Cells["B6:M6"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        RawResultsSheet.Cells["B6:M6"].Style.Fill.BackgroundColor.SetColor(Color.Gray);

                        log.SummaryGraphData = GetDecodeTimesAndAverages(log.LogFilePath);
                        log.DecodeTimes = Flatten2DArray(GetDecodeTimesOnly(log.SummaryGraphData));
                        log.SnappyData = GetSnappyTimes(GetDecodeTimesAndAverages(log.LogFilePath));
                        log.SnappyDecodeTimes = Flatten2DArray(GetDecodeTimesOnly(log.SnappyData));//AllLogPaths

                        int Rows = log.SummaryGraphData.GetLength(0);
                        int Cols = log.SummaryGraphData.GetLength(1);

                        for (int i = 0; i < Rows; i++)
                        {
                            for (int j = 0; j < Cols; j++)
                            {
                                //RawResultsSheet.Cells[i + 7, j + 2].Value = String.Format("{0:N2}", TLog.SummaryGraphData[i, j]);
                                RawResultsSheet.Cells[i + 7, j + 2].Value = log.SummaryGraphData[i, j];
                                RawResultsSheet.Cells[i + 7, j + 2].Style.Numberformat.Format = "0";
                            }
                        }

                        //Formatting for the graph data
                        RawResultsSheet.Cells["B:B"].Style.Numberformat.Format = "0.00";
                        RawResultsSheet.Cells["M:M"].Style.Numberformat.Format = "0.00";

                        //
                        var LineGraph = RawResultsSheet.Drawings.AddChart("Results", OfficeOpenXml.Drawing.Chart.eChartType.LineMarkersStacked);

                        //From Row, From Col - To Row, To Col
                        var YRange = RawResultsSheet.Cells[7, 2, GetDistancesOnly(log.SummaryGraphData).Length + 6, 2];
                        var XRange = RawResultsSheet.Cells[7, 13, GetDistanceAveragesOnly(log.SummaryGraphData).Length + 6, 13];
                        var Series = LineGraph.Series.Add(XRange, YRange);

                        LineGraph.Title.Text = log.ScannerName + "_" + log.BarcodeSampleName + "_Average_Results";
                        LineGraph.XAxis.Title.Text = "Distance (In)";
                        LineGraph.YAxis.Title.Text = "Decode Time (MS)";
                        LineGraph.XAxis.MajorUnit = .1;
                        LineGraph.SetPosition(4, 10, 13, 10);
                        LineGraph.SetSize(1000, 600);
                    }
                }
                catch (Exception)
                {

                    ErrorMessageTextBox.Text = "There was an error creating the raw data tab. Did you select a propper log file to compile? Did you select the same log twice?";
                }

                //creates the scorecard summary tab for each log
                try
                {
                    foreach (TargetLog log in DataSet)
                    {
                        var ScorecardSheet = excelPackage.Workbook.Worksheets.Add("SC_" + log.ScannerName + "_" + log.BarcodeSampleName);

                        //scorecard header styling
                        ScorecardSheet.Cells["B1"].Value = "Decode Range (In):";
                        ScorecardSheet.Cells["B1"].Style.Font.Bold = true;
                        ScorecardSheet.Cells["C1"].Value = "The largest range of distances that the scanner decoded at with no gap in decode. If the scanner decoded outside of this range, it is NOT included.";
                        ScorecardSheet.Cells["C1"].Style.Font.Bold = true;

                        ScorecardSheet.Cells["B2"].Value = "Minimum Distance (In):";
                        ScorecardSheet.Cells["B2"].Style.Font.Bold = true;
                        ScorecardSheet.Cells["C2"].Value = "The minimum distance of the Decode Range.";
                        ScorecardSheet.Cells["C2"].Style.Font.Bold = true;

                        ScorecardSheet.Cells["B3"].Value = "Maximum Distance (In):";
                        ScorecardSheet.Cells["B3"].Style.Font.Bold = true;
                        ScorecardSheet.Cells["C3"].Value = "The maximum distance of the Decode Range.";
                        ScorecardSheet.Cells["C3"].Style.Font.Bold = true;

                        ScorecardSheet.Cells["B4"].Value = "Total Average (MS):";
                        ScorecardSheet.Cells["B4"].Style.Font.Bold = true;
                        ScorecardSheet.Cells["C4"].Value = "The average decode time of the decode range.";
                        ScorecardSheet.Cells["C4"].Style.Font.Bold = true;

                        /*Commenting out becasue this feature is no longer needed
                        ScorecardSheet.Cells[""].Value = "Total Average (90%) (MS):";
                        ScorecardSheet.Cells[""].Style.Font.Bold = true;
                        ScorecardSheet.Cells[""].Value = "The average decode time of the first 90% of the decode range.";
                        ScorecardSheet.Cells[""].Style.Font.Bold = true;*/

                        ScorecardSheet.Cells["B5"].Value = "Standard Deviation (MS):";
                        ScorecardSheet.Cells["B5"].Style.Font.Bold = true;
                        ScorecardSheet.Cells["C5"].Value = "The standard deviation of the decode range.";
                        ScorecardSheet.Cells["C5"].Style.Font.Bold = true;

                        ScorecardSheet.Cells["B6"].Value = "Highest Decode Time (MS)";
                        ScorecardSheet.Cells["B6"].Style.Font.Bold = true;
                        ScorecardSheet.Cells["C6"].Value = "The highest decode time from the Decode Range.";
                        ScorecardSheet.Cells["C6"].Style.Font.Bold = true;

                        ScorecardSheet.Cells["B7"].Value = "Lowest Decode Time (MS):";
                        ScorecardSheet.Cells["B7"].Style.Font.Bold = true;
                        ScorecardSheet.Cells["C7"].Value = "The lowest decode time from the Deocde Range.";
                        ScorecardSheet.Cells["C7"].Style.Font.Bold = true;

                        ScorecardSheet.Cells["B9"].Value = "Total Range:";
                        ScorecardSheet.Cells["B9"].Style.Font.Bold = true;
                        ScorecardSheet.Cells["C9"].Value = "Statistics for the entire range of decodes consisting of any scans as long as there was a decode of any length.";
                        ScorecardSheet.Cells["C9"].Style.Font.Bold = true;

                        ScorecardSheet.Cells["B10"].Value = "Snappy Range:";
                        ScorecardSheet.Cells["B10"].Style.Font.Bold = true;
                        ScorecardSheet.Cells["C10"].Value = "Statistics for only the \"snappiest\" range where decode times are 500ms and under.";
                        ScorecardSheet.Cells["C10"].Style.Font.Bold = true;

                        //prints sample name
                        ScorecardSheet.Cells["C13"].Value = log.BarcodeSampleName;
                        ScorecardSheet.Cells["C13"].Style.Font.Bold = true;
                        ScorecardSheet.Cells["C13"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        ScorecardSheet.Cells["C13"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        ScorecardSheet.Cells["C13"].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#FFF2CC"));

                        //prints scanner name
                        ScorecardSheet.Cells["C14"].Value = log.ScannerMake + "_" + log.ScannerName;
                        ScorecardSheet.Cells["C14"].Style.Font.Bold = true;
                        ScorecardSheet.Cells["C14"].Style.Border.Top.Style = ExcelBorderStyle.Thick;
                        ScorecardSheet.Cells["C14"].Style.Border.Right.Style = ExcelBorderStyle.Thick;
                        ScorecardSheet.Cells["C14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        ScorecardSheet.Cells["C14"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        ScorecardSheet.Cells["C14"].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#DDEBF7"));

                        //Notes field for the user to enter in any relevant notes about the scanner
                        ScorecardSheet.Cells["C15"].Value = "Notes:";
                        ScorecardSheet.Cells["C15"].Style.Font.Bold = true;
                        ScorecardSheet.Cells["C15"].Style.Border.Right.Style = ExcelBorderStyle.Thick;
                        ScorecardSheet.Cells["C15"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        ScorecardSheet.Cells["C15"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        ScorecardSheet.Cells["C15"].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#C5D9F1"));
                        ScorecardSheet.Row(16).Height = 75;
                        ScorecardSheet.Cells["C16"].Style.Border.Right.Style = ExcelBorderStyle.Thick;
                        ScorecardSheet.Cells["C16"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        ScorecardSheet.Cells["C16"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        ScorecardSheet.Cells["C16"].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#C5D9F1"));

                        //worksheet styling
                        var TotalRangeTag = ScorecardSheet.Cells["A17:A23"];
                        TotalRangeTag.Merge = true;
                        TotalRangeTag.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        TotalRangeTag.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        ScorecardSheet.Cells["A17"].Value = "Total Range";
                        ScorecardSheet.Cells["A17"].Style.TextRotation = 45;

                        ScorecardSheet.Column(1).Width = 14;
                        ScorecardSheet.Column(2).Width = 30;
                        ScorecardSheet.Column(3).Width = 35;

                        ScorecardSheet.Cells["C17:C30"].Style.Numberformat.Format = "0.00";

                        var SnappyRangeTag = ScorecardSheet.Cells["A24:A30"];
                        SnappyRangeTag.Merge = true;
                        SnappyRangeTag.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        SnappyRangeTag.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        ScorecardSheet.Cells["A24"].Value = "Snappy Range";
                        ScorecardSheet.Cells["A24"].Style.TextRotation = 45;

                        //total range results
                        ScorecardSheet.Cells["B17"].Value = "Minimum Distance (In):";
                        ScorecardSheet.Cells["B17"].Style.Font.Bold = true;
                        ScorecardSheet.Cells["C17"].Value = CalculateClosestRange(log.SummaryGraphData);
                        ScorecardSheet.Cells["C17"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        ScorecardSheet.Cells["C17"].Style.Border.Right.Style = ExcelBorderStyle.Thick;

                        ScorecardSheet.Cells["B18"].Value = "Maximum Distance (In):";
                        ScorecardSheet.Cells["B18"].Style.Font.Bold = true;
                        ScorecardSheet.Cells["C18"].Value = CalculateFarthestRange(log.SummaryGraphData);
                        ScorecardSheet.Cells["C18"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        ScorecardSheet.Cells["C18"].Style.Border.Right.Style = ExcelBorderStyle.Thick;

                        ScorecardSheet.Cells["B19"].Value = "Decode Range (In):";
                        ScorecardSheet.Cells["B19"].Style.Font.Bold = true;
                        ScorecardSheet.Cells["C19"].Value = CalculateDecodeRange(log.SummaryGraphData);
                        ScorecardSheet.Cells["C19"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        ScorecardSheet.Cells["C19"].Style.Border.Right.Style = ExcelBorderStyle.Thick;

                        ScorecardSheet.Cells["B20"].Value = "Total Average (MS):";
                        ScorecardSheet.Cells["B20"].Style.Font.Bold = true;
                        ScorecardSheet.Cells["C20"].Value = CalculateAverageDecodeTime(log.SummaryGraphData);
                        ScorecardSheet.Cells["C20"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        ScorecardSheet.Cells["C20"].Style.Border.Right.Style = ExcelBorderStyle.Thick;

                        /*Commenting out becasue this feature is no longer needed
                        ScorecardSheet.Cells[""].Value = "Total Average (90%) (MS):";
                        ScorecardSheet.Cells[""].Style.Font.Bold = true;
                        ScorecardSheet.Cells[""].Value = CalculateNinetyPercentAverageDecodeTime(log.SummaryGraphData);
                        ScorecardSheet.Cells[""].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        ScorecardSheet.Cells[""].Style.Border.Right.Style = ExcelBorderStyle.Thick;*/

                        ScorecardSheet.Cells["B21"].Value = "Standard Deviation (MS):";
                        ScorecardSheet.Cells["B21"].Style.Font.Bold = true;
                        ScorecardSheet.Cells["C21"].Value = CalculateStandardDeviationDecodeTime(log.DecodeTimes);
                        ScorecardSheet.Cells["C21"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        ScorecardSheet.Cells["C21"].Style.Border.Right.Style = ExcelBorderStyle.Thick;

                        ScorecardSheet.Cells["B22"].Value = "Highest Decode Time (MS)";
                        ScorecardSheet.Cells["B22"].Style.Font.Bold = true;
                        ScorecardSheet.Cells["C22"].Value = CalculateHighestDecodeTime(log.DecodeTimes);
                        ScorecardSheet.Cells["C22"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        ScorecardSheet.Cells["C22"].Style.Border.Right.Style = ExcelBorderStyle.Thick;

                        ScorecardSheet.Cells["B23"].Value = "Lowest Decode Time (MS):";
                        ScorecardSheet.Cells["B23"].Style.Font.Bold = true;
                        ScorecardSheet.Cells["C23"].Value = CalculateLowestDecodeTime(log.DecodeTimes);
                        ScorecardSheet.Cells["C23"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        ScorecardSheet.Cells["C23"].Style.Border.Right.Style = ExcelBorderStyle.Thick;

                        //styling for total range results
                        var TotalRangeResults = ScorecardSheet.Cells["A17:C23"];
                        TotalRangeResults.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        TotalRangeResults.Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#F2F2F2"));

                        //snappy range results
                        ScorecardSheet.Cells["B24"].Value = "Snappy Minimum Distance (In):";
                        ScorecardSheet.Cells["B24"].Style.Font.Bold = true;
                        ScorecardSheet.Cells["C24"].Value = CalculateClosestRange(log.SnappyData);
                        ScorecardSheet.Cells["C24"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        ScorecardSheet.Cells["C24"].Style.Border.Right.Style = ExcelBorderStyle.Thick;

                        ScorecardSheet.Cells["B25"].Value = "Snappy Maximum Distance (In):";
                        ScorecardSheet.Cells["B25"].Style.Font.Bold = true;
                        ScorecardSheet.Cells["C25"].Value = CalculateFarthestRange(log.SnappyData);
                        ScorecardSheet.Cells["C25"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        ScorecardSheet.Cells["C25"].Style.Border.Right.Style = ExcelBorderStyle.Thick;

                        ScorecardSheet.Cells["B26"].Value = "Snappy Decode Range (In):";
                        ScorecardSheet.Cells["B26"].Style.Font.Bold = true;
                        ScorecardSheet.Cells["C26"].Value = CalculateDecodeRange(log.SnappyData);
                        ScorecardSheet.Cells["C26"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        ScorecardSheet.Cells["C26"].Style.Border.Right.Style = ExcelBorderStyle.Thick;

                        ScorecardSheet.Cells["B27"].Value = "Snappy Total Average (MS):";
                        ScorecardSheet.Cells["B27"].Style.Font.Bold = true;
                        ScorecardSheet.Cells["C27"].Value = CalculateAverageDecodeTime(log.SnappyData);
                        ScorecardSheet.Cells["C27"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        ScorecardSheet.Cells["C27"].Style.Border.Right.Style = ExcelBorderStyle.Thick;

                        /*Commenting out becasue this feature is no longer needed
                        ScorecardSheet.Cells[""].Value = "Snappy Total Average (90%) (MS):";
                        ScorecardSheet.Cells[""].Style.Font.Bold = true;
                        ScorecardSheet.Cells[""].Value = CalculateNinetyPercentAverageDecodeTime(og.SummaryGraphData);
                        ScorecardSheet.Cells[""].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        ScorecardSheet.Cells[""].Style.Border.Right.Style = ExcelBorderStyle.Thick;*/

                        ScorecardSheet.Cells["B28"].Value = "Snappy Standard Deviation (MS):";
                        ScorecardSheet.Cells["B28"].Style.Font.Bold = true;
                        ScorecardSheet.Cells["C28"].Value = CalculateStandardDeviationDecodeTime(log.DecodeTimes);
                        ScorecardSheet.Cells["C28"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        ScorecardSheet.Cells["C28"].Style.Border.Right.Style = ExcelBorderStyle.Thick;

                        ScorecardSheet.Cells["B29"].Value = "Snappy Highest Decode Time (MS)";
                        ScorecardSheet.Cells["B29"].Style.Font.Bold = true;
                        ScorecardSheet.Cells["C29"].Value = CalculateHighestDecodeTime(log.DecodeTimes);
                        ScorecardSheet.Cells["C29"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        ScorecardSheet.Cells["C29"].Style.Border.Right.Style = ExcelBorderStyle.Thick;

                        ScorecardSheet.Cells["B30"].Value = "Lowest Decode Time (MS):";
                        ScorecardSheet.Cells["B30"].Style.Font.Bold = true;
                        ScorecardSheet.Cells["C30"].Value = CalculateLowestDecodeTime(log.DecodeTimes);
                        ScorecardSheet.Cells["C30"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        ScorecardSheet.Cells["C30"].Style.Border.Right.Style = ExcelBorderStyle.Thick;

                        //styling for snappy range results
                        var SnappyRangeResults = ScorecardSheet.Cells["A24:C30"];
                        SnappyRangeResults.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        SnappyRangeResults.Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D9D9D9"));
                    }
                }
                catch (Exception)
                {

                    ErrorMessageTextBox.Text = "There was an error creating the scorecard tab. Did you select a propper log file to compile? Did you select the same log twice?";
                }

                //create the spreadsheet and save it to a new folder on the users desktop
                try
                {
                    if (!DataSet.Any())
                    {
                        //checks if the dataset is empty before compiling
                        ErrorMessageTextBox.Text = "You have not selected any valid log files to compile.";
                    }
                    else
                    {
                        // Create a new folder on the desktop
                        string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                        string SaveFolder = Path.Combine(desktopPath, "Hakbal_Files");
                        if (!Directory.Exists(SaveFolder))
                        {
                            Directory.CreateDirectory(SaveFolder);
                        }

                        //saves the file to the desktop
                        string SavePath = Path.Combine(SaveFolder, "Compiled Result_" + DateTime.Now.ToString("yyyyMMddHHmmssffff") + ".xlsx");
                        var FileInfo = new System.IO.FileInfo(SavePath);
                        excelPackage.SaveAs(FileInfo);
                    }
                }
                catch (Exception)
                {

                    ErrorMessageTextBox.Text = "There was an error compiling these logs together. Did you select valid log files?";
                }
            }
        }
        #endregion

        #region Excel_Helper_Functions
        private void PopulateRawResultsSheet(TargetLog log, ExcelPackage sheet)
        {
            //wip for 3.0
        }
        #endregion

        #region Math_Helper_Functions
        /// <summary>
        /// Seperates the single string contaning all of the log files into an array 
        /// </summary>
        /// <param name="LogPathsString"></param>
        /// <returns></returns>
        private string[] CreateLogFilePathsArray(string LogPathsString)
        {
            string[] LogPathsArray = LogPathsString.Split(new[] { '\n' }, StringSplitOptions.RemoveEmptyEntries);

            return LogPathsArray;
        }

        /// <summary>
        /// Reads through the selected log file to get decode data and adds average decode tims for each distance
        /// </summary>
        /// <param name="LogPath"></param>
        /// <returns></returns>
        private float[,] GetDecodeTimesAndAverages(string LogPath)
        {
            //All log content goes in this StringArray
            string[] LogContent = File.ReadAllLines(LogPath);

            //Filters out lines without relevant data
            var ModifiedLogContent = LogContent.Select(line => line.Split(new[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries)).Where(words => words.Contains("DECODE_TIME")).ToArray();

            //New StringArray to hold only decode data
            string[,] FinalStringLogContent = new string[ModifiedLogContent.Length, ModifiedLogContent.Max(words => words.Length)];

            //Fills new StringArray with only decode data + ranges
            for (int i = 0; i < ModifiedLogContent.Length; i++)
            {
                for (int j = 0; j < ModifiedLogContent[i].Length; j++)
                {
                    FinalStringLogContent[i, j] = ModifiedLogContent[i][j];
                }
            }

            //Gets rid of any words in the string array
            FinalStringLogContent = RemoveLettersAndWords(FinalStringLogContent);

            //Converts String array to Float array
            float[,] FinalFloatLogContent = ConvertStringArrayToFloatArray(FinalStringLogContent);

            //Adds the average for each distance to the end of each distances decode times
            FinalFloatLogContent = CalculateDistanceAverageDecodeTimes(FinalFloatLogContent);

            return FinalFloatLogContent;
        }

        /// <summary>
        /// Gets rid of unwanted letters and words from a 2D StringArray
        /// </summary>
        /// <param name="StringArray"></param>
        /// <returns></returns>
        private string[,] RemoveLettersAndWords(string[,] StringArray)
        {
            //Gets rid of any words from StringArray
            for (int i = 0; i < StringArray.GetLength(0); i++)
            {
                string[] words = StringArray[i, 0].Split(new[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);
                words = words.Where(word => !word.Any(char.IsLetter)).ToArray();
                StringArray[i, 0] = string.Join("", words);
            }

            //gets rid of blank entries
            int Rows = StringArray.GetLength(0);
            int Cols = StringArray.GetLength(1) - 1;

            String[,] ModifiedStringArray = new string[Rows, Cols];

            //gets rid of first entry which will be empty
            for (int i = 0; i < Rows; i++)
            {
                for (int j = 1; j < Cols + 1; j++)
                {
                    ModifiedStringArray[i, j - 1] = StringArray[i, j];
                }
            }

            return ModifiedStringArray;
        }

        /// <summary>
        /// Converts any 2D string StringArray into a 2D Float StringArray
        /// </summary>
        /// <param name="StringArray"></param>
        /// <returns></returns>
        private float[,] ConvertStringArrayToFloatArray(string[,] StringArray)
        {
            int Rows = StringArray.GetLength(0);
            int Cols = StringArray.GetLength(1);

            float[,] FloatArray = new float[Rows, Cols];

            for (int i = 0; i < Rows; i++)
            {
                for (int j = 0; j < Cols; j++)
                {
                    if (float.TryParse(StringArray[i, j], out float Result))
                    {
                        FloatArray[i, j] = Result;
                    }
                    else
                    {
                        //do nothing
                    }
                }
            }

            return FloatArray;
        }

        /// <summary>
        /// Calculates the average decode time at each distance, and adds that average to the end of array for each distance
        /// </summary>
        /// <param name="FloatArray"></param>
        /// <returns></returns>
        private float[,] CalculateDistanceAverageDecodeTimes(float[,] FloatArray)
        {
            //Adds an additional entry in each row of the array to store the average for each distance
            int Rows = FloatArray.GetLength(0);
            int Cols = FloatArray.GetLength(1);

            float[,] BiggerFloatArray = new float[Rows, Cols + 1];

            for (int i = 0; i < Rows; i++)
            {
                for (int j = 0; j < Cols; j++)
                {
                    BiggerFloatArray[i, j] = FloatArray[i, j];
                }

                BiggerFloatArray[i, Cols] = 0;
            }

            //Adds the Average to the end of each row
            int Rows2 = BiggerFloatArray.GetLength(0);
            int Cols2 = BiggerFloatArray.GetLength(1) - 1;

            for (int i = 0; i < Rows2; i++)
            {
                float DistanceSum = 0;

                for (int j = 1; j < Cols2 + 1; j++)
                {
                    DistanceSum += BiggerFloatArray[i, j];
                }

                float DistanceAverage = DistanceSum / 10;

                BiggerFloatArray[i, Cols2] = DistanceAverage;
            }

            return BiggerFloatArray;
        }

        /// <summary>
        /// Finds the range with the "snappiest" decode times (500ms and under) 
        /// </summary>
        /// <param name="FloatArray"></param>
        /// <returns></returns>
        private float[,] GetSnappyTimes(float[,] FloatArray)
        {
            int Rows = FloatArray.GetLength(0);
            int Cols = FloatArray.GetLength(1);

            int ResultCount = 0;

            //check if initial array is empty
            if (FloatArray.GetLength(0) == 0 && FloatArray.GetLength(1) == 0)
            {
                //array is empty
                return FloatArray;
            }

            //iterate through each row starting from the first row to get the new size
            for (int i = 0; i < Rows; i++)
            {
                //check if snappy time or not
                if (FloatArray[i, Cols - 1] <= 500 && FloatArray[i, Cols - 1] != 0)
                {
                    ResultCount++;
                }
            }

            //new 2d float array to hold the results of the snappy range only
            float[,] SnappyLogConetent = new float[ResultCount, Cols];

            int CurrentIndex = 0;

            //gets only the ranges in the snappy range and adds them to a new 2d float array
            for (int i = 0; i < Rows; i++)
            {
                if (FloatArray[i, Cols - 1] <= 500 && FloatArray[i, Cols - 1] != 0)
                {
                    for (int j = 0; j < Cols; j++)
                    {
                        SnappyLogConetent[CurrentIndex, j] = FloatArray[i, j];
                    }
                    CurrentIndex++;
                }
            }

            //checks if the new array did not find a snappy area
            if (SnappyLogConetent.GetLength(0) == 0 && SnappyLogConetent.GetLength(1) == 0)
            {
                return FloatArray;
            }

            return SnappyLogConetent;
        }

        /// <summary>
        /// Calculates the closest decocde range
        /// </summary>
        /// <returns></returns>
        private float CalculateClosestRange(float[,] FloatArray)
        {
            float ClosestRange = 0;

            int Rows = FloatArray.GetLength(0);

            if (FloatArray.Length == 0)
            {
                ClosestRange = 0;
            }
            else
            {
                for (int i = 0; i < Rows; i++)
                {
                    if (FloatArray[i, 1] == 0)
                    {
                        //do nothing
                    }
                    else
                    {
                        ClosestRange = FloatArray[i, 0];
                        return ClosestRange;
                    }
                }
            }

            return ClosestRange;
        }

        /// <summary>
        /// Calculates the fartherst range
        /// </summary>
        /// <returns></returns>
        private float CalculateFarthestRange(float[,] FloatArray)
        {
            float FarthestRange;

            if (FloatArray.Length == 0)
            {
                FarthestRange = 0;
            }
            else
            {
                int Rows = FloatArray.GetLength(0);

                FarthestRange = FloatArray[Rows - 2, 0];
            }

            return FarthestRange;
        }

        /// <summary>
        /// Calculates the decode range using the closest and farthest range
        /// </summary>
        /// <param name="FloatArray"></param>
        /// <returns></returns>
        private float CalculateDecodeRange(float[,] FloatArray)
        {
            float DecodeRange = 0;

            DecodeRange = CalculateFarthestRange(FloatArray) - CalculateClosestRange(FloatArray);

            return DecodeRange;
        }

        /// <summary>
        /// Calculates the average decode time from all decode times in the log
        /// </summary>
        /// <param name="FloatArray"></param>
        /// <returns></returns>
        private float CalculateAverageDecodeTime(float[,] FloatArray)
        {
            float TotalSumDecodeTime = 0;
            float TotalAverageDecodeTime;

            float[] AllDistanceAverages = new float[FloatArray.GetLength(0)];

            for (int i = 0; i < FloatArray.GetLength(0); i++)
            {
                AllDistanceAverages[i] = FloatArray[i, FloatArray.GetLength(1) - 1];
            }

            foreach (float element in AllDistanceAverages)
            {
                TotalSumDecodeTime += element;
            }

            TotalAverageDecodeTime = TotalSumDecodeTime / AllDistanceAverages.Length;

            return TotalAverageDecodeTime;
        }

        /*Commenting out becasue this feature is no longer needed
        /// <summary>
        /// Calculates the average decode time from the first 90% of all decode times in a log
        /// </summary>
        /// <returns></returns>
        private float CalculateNinetyPercentAverageDecodeTime(float[,] FloatArray)
        {
            float NinetyPercentAverageDecodeTime = 0;

            NinetyPercentAverageDecodeTime = CalculateAverageDecodeTime(FloatArray) * 0.90f;

            return NinetyPercentAverageDecodeTime;
        }*/

        /// <summary>
        /// Calculates the standard deviation decode time from all decodes in a log
        /// </summary>
        /// <param name="FloatArray"></param>
        /// <returns></returns>
        private float CalculateStandardDeviationDecodeTime(float[] FloatArray)
        {
            float StandardDeviationDecodeTime = 0;
            float Sum = 0;
            float Mean = 0;
            float SumSquaredDifferences = 0;
            float Variance = 0;

            foreach (float Value in FloatArray)
            {
                Sum += Value;
            }

            Mean = Sum / FloatArray.Length;

            foreach (float Value in FloatArray)
            {
                SumSquaredDifferences += (Value - Mean) * (Value - Mean);
            }

            Variance = SumSquaredDifferences / FloatArray.Length;

            StandardDeviationDecodeTime = (float)Math.Sqrt(Variance);

            return StandardDeviationDecodeTime;
        }

        /// <summary>
        /// Determines the highest decode time from all decodes in a log
        /// </summary>
        /// <param name="FloatArray"></param>
        /// <returns></returns>
        private float CalculateHighestDecodeTime(float[] FloatArray)
        {
            float HighestDecodeTime = 0;

            Array.Sort(FloatArray);

            HighestDecodeTime = FloatArray[FloatArray.Length - 1];

            return HighestDecodeTime;
        }

        /// <summary>
        /// Determines the lowest decode time from all decodes in a log
        /// </summary>
        /// <param name="FloatArray"></param>
        /// <returns></returns>
        private float CalculateLowestDecodeTime(float[] FloatArray)
        {
            float LowestDecodeTime = 0;

            List<float> NoZeroesList = new List<float>();

            foreach (float Value in FloatArray)
            {
                if (Value != 0.0f)
                {
                    NoZeroesList.Add(Value);
                }
            }

            float[] NoZeroesArray = NoZeroesList.ToArray();

            Array.Sort(NoZeroesArray);

            LowestDecodeTime = NoZeroesArray[0];

            return LowestDecodeTime;
        }

        /// <summary>
        /// Gets only the raw decode times from the array
        /// </summary>
        /// <param name="FloatArray"></param>
        /// <returns></returns>
        private float[,] GetDecodeTimesOnly(float[,] FloatArray)
        {
            int Rows = FloatArray.GetLength(0);
            int Cols = FloatArray.GetLength(1);

            float[,] DecodeTimesOnly = new float[Rows, Cols - 2];

            for (int i = 0; i < Rows; i++)
            {
                for (int j = 1; j < Cols - 1; j++)
                {
                    DecodeTimesOnly[i, j - 1] = FloatArray[i, j];
                }
            }

            return DecodeTimesOnly;
        }

        /// <summary>
        /// Gets Distances in order from the array for the Excel line graph
        /// </summary>
        /// <param name="FloatArray"></param>
        /// <returns></returns>
        private float[] GetDistancesOnly(float[,] FloatArray)
        {
            int Rows = FloatArray.GetLength(0);
            float[] Distances = new float[Rows];

            for (int i = 0; i < Rows; i++)
            {
                Distances[i] = FloatArray[i, 0];
            }

            return Distances;
        }

        /// <summary>
        /// Gets Distance Averages in order from the array for the Excel line graph
        /// </summary>
        /// <param name="FloatArray"></param>
        /// <returns></returns>
        private float[] GetDistanceAveragesOnly(float[,] FloatArray)
        {
            int Rows = FloatArray.GetLength(0);
            float[] DistanceAverages = new float[Rows];

            for (int i = 0; i < Rows; i++)
            {
                int Cols = FloatArray.GetLength(1);
                DistanceAverages[i] = FloatArray[i, Cols - 1];
            }

            return DistanceAverages;
        }

        /// <summary>
        /// Flattens a 2D float array into a 1D float array
        /// </summary>
        /// <param name="FloatArray"></param>
        /// <returns></returns>
        private float[] Flatten2DArray(float[,] FloatArray)
        {
            float[] FlattenedArray = new float[FloatArray.Length];

            int index = 0;

            foreach (float Value in FloatArray)
            {
                FlattenedArray[index++] = Value;
            }

            return FlattenedArray;
        }

        /// <summary>
        /// Gets the file name for the graph without the full path
        /// </summary>
        /// <returns></returns>
        private string GetFileName(string FullPath)
        {
            // Find the last occurrence of "\"
            int lastIndex = FullPath.LastIndexOf('\\');

            // If "\" is found
            if (lastIndex >= 0)
            {
                // Extract the substring after the last "\"
                return FullPath.Substring(lastIndex + 1);
            }
            else
            {
                // If "\" is not found, return the original input string
                return FullPath;
            }
        }
        #endregion
    }
}
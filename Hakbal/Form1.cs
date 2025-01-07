using Microsoft.VisualBasic.Logging;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Data.Common;
using System.IO.Packaging;
using System.Linq.Expressions;
using System.Windows.Forms;
using System.Linq;
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
            HowToCompareResultsComboBox.SelectedIndex = 0;
            ScorecardGroupNumberComboBox.SelectedIndex = 0;
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
                TLog.ScorecardGroupNumber = ScorecardGroupNumberComboBox.Text;
                TLog.LogFilePath = LogFilePathTextBox.Text.Replace("\"", ""); //gets rid of quotes incase they are there from copying the path

                if (String.IsNullOrEmpty(ScannerNameTextBox.Text) || String.IsNullOrEmpty(ScannerMakeTextBox.Text) || String.IsNullOrEmpty(ScannerSerialNumberTextBox.Text) || String.IsNullOrEmpty(BarcodeSampleNameTextBox.Text) || ScorecardGroupNumberComboBox.SelectedItem == null || String.IsNullOrEmpty(LogFilePathTextBox.Text))
                {
                    throw new Exception();
                }

                DataSet.Add(TLog);

                //prints data set to text box
                if (TLog.ToString() != "Log: , , ")
                {
                    DataSetListBox.Items.Add(TLog.ToString());
                }
                else
                {
                    throw new Exception();
                }
            }
            catch (Exception)
            {
                ErrorMessageTextBox.Text = "you missed a field! Please enter all information about your log.";
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

                //DOESNT WORK FIX THIS
                //DataSetListBox.Items.Remove(DataSetListBox.Items[DataSetListBox.SelectedIndex]);

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
                    int LogCount = 1;

                    foreach (TargetLog log in DataSet)
                    {
                        //creates the tab for each log in the data set
                        var RawResultsSheet = excelPackage.Workbook.Worksheets.Add(LogCount.ToString() + "_" + log.ScannerName + "_" + log.BarcodeSampleName);
                        LogCount++;

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
                        //RawResultsSheet.Cells["B5"].Value = GetFileName(LogFilePathTextBox.Text.ToString()); //TODO FIX THIS LINE!!!!!!!!!!!!!!!!!!!!!!!!!
                        RawResultsSheet.Cells["B5"].Value = GetFileName(log.LogFilePath.ToString());

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

                        //From CurrentRow, From Col - To CurrentRow, To Col
                        var YRange = RawResultsSheet.Cells[7, 2, GetDistancesOnly(log.SummaryGraphData).Length + 6, 2];
                        var XRange = RawResultsSheet.Cells[7, 13, GetDistanceAveragesOnlyForGraph(log.SummaryGraphData).Length + 6, 13];
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
                    //Scorecard tab creation
                    var ScorecardSheet = excelPackage.Workbook.Worksheets.Add("Scorecard");

                    //freezes the first columns for easy visibility
                    ScorecardSheet.View.FreezePanes(1, 3);

                    //Scorecard tab column sizing
                    ScorecardSheet.Column(1).Width = 14;
                    ScorecardSheet.Column(2).Width = 32;
                    for (char column = 'C'; column <= 'Z'; column++)
                    {
                        ScorecardSheet.Column(column - 'A' + 1).Width = 35;
                    }

                    //scorecard tab header styling
                    ScorecardSheet.Cells["B1"].Value = "Decode Range (in):";
                    ScorecardSheet.Cells["B1"].Style.Font.Bold = true;
                    ScorecardSheet.Cells["C1"].Value = "The largest range of distances that the scanner decoded at with no gap in decode. If the scanner decoded outside of this range, it is NOT included.";
                    ScorecardSheet.Cells["C1"].Style.Font.Bold = true;

                    ScorecardSheet.Cells["B2"].Value = "Minimum Distance (in):";
                    ScorecardSheet.Cells["B2"].Style.Font.Bold = true;
                    ScorecardSheet.Cells["C2"].Value = "The minimum distance of the Decode Range.";
                    ScorecardSheet.Cells["C2"].Style.Font.Bold = true;

                    ScorecardSheet.Cells["B3"].Value = "Maximum Distance (in):";
                    ScorecardSheet.Cells["B3"].Style.Font.Bold = true;
                    ScorecardSheet.Cells["C3"].Value = "The maximum distance of the Decode Range.";
                    ScorecardSheet.Cells["C3"].Style.Font.Bold = true;

                    ScorecardSheet.Cells["B4"].Value = "Decode Time Average (ms):";
                    ScorecardSheet.Cells["B4"].Style.Font.Bold = true;
                    ScorecardSheet.Cells["C4"].Value = "The average decode time of the first 90% of the decode range.";
                    ScorecardSheet.Cells["C4"].Style.Font.Bold = true;

                    ScorecardSheet.Cells["B5"].Value = "Std Dev Decode Time (ms):";
                    ScorecardSheet.Cells["B5"].Style.Font.Bold = true;
                    ScorecardSheet.Cells["C5"].Value = "The standard deviation decode time of the first 90% of the decode range.";
                    ScorecardSheet.Cells["C5"].Style.Font.Bold = true;

                    ScorecardSheet.Cells["B6"].Value = "Highest Decode Time (ms)";
                    ScorecardSheet.Cells["B6"].Style.Font.Bold = true;
                    ScorecardSheet.Cells["C6"].Value = "The highest decode time from the first 90% Decode Range.";
                    ScorecardSheet.Cells["C6"].Style.Font.Bold = true;

                    ScorecardSheet.Cells["B7"].Value = "Lowest Decode Time (ms):";
                    ScorecardSheet.Cells["B7"].Style.Font.Bold = true;
                    ScorecardSheet.Cells["C7"].Value = "The lowest decode time from the first 90% Deocde Range.";
                    ScorecardSheet.Cells["C7"].Style.Font.Bold = true;

                    ScorecardSheet.Cells["B9"].Value = "Total Range:";
                    ScorecardSheet.Cells["B9"].Style.Font.Bold = true;
                    ScorecardSheet.Cells["C9"].Value = "Statistics for the entire range of decodes consisting of any scans as long as there was a decode of any length.";
                    ScorecardSheet.Cells["C9"].Style.Font.Bold = true;

                    ScorecardSheet.Cells["B10"].Value = "Snappy Range:";
                    ScorecardSheet.Cells["B10"].Style.Font.Bold = true;
                    ScorecardSheet.Cells["C10"].Value = "Statistics for only the \"snappiest\" range where decode times are 100ms and under.";
                    ScorecardSheet.Cells["C10"].Style.Font.Bold = true;

                    //styling
                    ScorecardSheet.Column(1).Width = 14;
                    ScorecardSheet.Column(2).Width = 30;
                    ScorecardSheet.Column(3).Width = 35;

                    //prints out summaries for each log
                    foreach (TargetLog log in DataSet)
                    {
                        //WriteLogDataToScorecard(ScorecardSheet, log, log.ScorecardGroupNumber);

                        //seperates the groups
                        if (log.ScorecardGroupNumber == "1")
                        {
                            if (ScorecardSheet.Cells["C13"].Value == null)
                            {
                                //total range results
                                var TotalRangeTag = ScorecardSheet.Cells["A17:A23"];
                                TotalRangeTag.Merge = true;
                                TotalRangeTag.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                TotalRangeTag.Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                                ScorecardSheet.Cells["A17"].Value = "Total Range";
                                ScorecardSheet.Cells["A17"].Style.TextRotation = 45;

                                ScorecardSheet.Cells["B17"].Value = "Minimum Distance (in):";
                                ScorecardSheet.Cells["B17"].Style.Font.Bold = true;

                                ScorecardSheet.Cells["B18"].Value = "Maximum Distance (in):";
                                ScorecardSheet.Cells["B18"].Style.Font.Bold = true;

                                ScorecardSheet.Cells["B19"].Value = "Decode Range (in):";
                                ScorecardSheet.Cells["B19"].Style.Font.Bold = true;

                                ScorecardSheet.Cells["B20"].Value = "Decode Time Average (ms):";
                                ScorecardSheet.Cells["B20"].Style.Font.Bold = true;

                                ScorecardSheet.Cells["B21"].Value = "Std Dev (ms):";
                                ScorecardSheet.Cells["B21"].Style.Font.Bold = true;

                                ScorecardSheet.Cells["B22"].Value = "Highest Decode Time (ms)";
                                ScorecardSheet.Cells["B22"].Style.Font.Bold = true;

                                ScorecardSheet.Cells["B23"].Value = "Lowest Decode Time (ms):";
                                ScorecardSheet.Cells["B23"].Style.Font.Bold = true;

                                //styling for total range results
                                var TotalRangeResults = ScorecardSheet.Cells["A17:C23"];
                                TotalRangeResults.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                TotalRangeResults.Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#F2F2F2"));

                                //snappy range results
                                var SnappyRangeTag = ScorecardSheet.Cells["A24:A30"];
                                SnappyRangeTag.Merge = true;
                                SnappyRangeTag.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                SnappyRangeTag.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                ScorecardSheet.Cells["A24"].Value = "Snappy Range";
                                ScorecardSheet.Cells["A24"].Style.TextRotation = 45;

                                ScorecardSheet.Cells["B24"].Value = "Snappy Minimum Distance (in):";
                                ScorecardSheet.Cells["B24"].Style.Font.Bold = true;

                                ScorecardSheet.Cells["B25"].Value = "Snappy Maximum Distance (in):";
                                ScorecardSheet.Cells["B25"].Style.Font.Bold = true;

                                ScorecardSheet.Cells["B26"].Value = "Snappy Decode Range (in):";
                                ScorecardSheet.Cells["B26"].Style.Font.Bold = true;

                                ScorecardSheet.Cells["B27"].Value = "Snappy Decode Average (ms):";
                                ScorecardSheet.Cells["B27"].Style.Font.Bold = true;

                                ScorecardSheet.Cells["B28"].Value = "Snappy Std Dev (ms):";
                                ScorecardSheet.Cells["B28"].Style.Font.Bold = true;

                                ScorecardSheet.Cells["B29"].Value = "Snappy Highest Decode Time (ms)";
                                ScorecardSheet.Cells["B29"].Style.Font.Bold = true;

                                ScorecardSheet.Cells["B30"].Value = "SnappyLowest Decode Time (ms):";
                                ScorecardSheet.Cells["B30"].Style.Font.Bold = true;

                                //styling for snappy range results
                                var SnappyRangeResults = ScorecardSheet.Cells["A24:C30"];
                                SnappyRangeResults.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                SnappyRangeResults.Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D9D9D9"));
                            }

                            WriteLogDataToScorecard(ScorecardSheet, log, log.ScorecardGroupNumber);
                        }
                        else if (log.ScorecardGroupNumber == "2")
                        {
                            if (ScorecardSheet.Cells["C33"].Value == null)
                            {
                                //total range results
                                var TotalRangeTag = ScorecardSheet.Cells["A37:A43"];
                                TotalRangeTag.Merge = true;
                                TotalRangeTag.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                TotalRangeTag.Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                                ScorecardSheet.Cells["A37"].Value = "Total Range";
                                ScorecardSheet.Cells["A37"].Style.TextRotation = 45;

                                ScorecardSheet.Cells["B37"].Value = "Minimum Distance (in):";
                                ScorecardSheet.Cells["B37"].Style.Font.Bold = true;

                                ScorecardSheet.Cells["B38"].Value = "Maximum Distance (in):";
                                ScorecardSheet.Cells["B38"].Style.Font.Bold = true;

                                ScorecardSheet.Cells["B39"].Value = "Decode Range (in):";
                                ScorecardSheet.Cells["B39"].Style.Font.Bold = true;

                                ScorecardSheet.Cells["B40"].Value = "Decode Time Average (ms):";
                                ScorecardSheet.Cells["B40"].Style.Font.Bold = true;

                                ScorecardSheet.Cells["B41"].Value = "Std Dev (ms):";
                                ScorecardSheet.Cells["B41"].Style.Font.Bold = true;

                                ScorecardSheet.Cells["B42"].Value = "Highest Decode Time (ms)";
                                ScorecardSheet.Cells["B42"].Style.Font.Bold = true;

                                ScorecardSheet.Cells["B43"].Value = "Lowest Decode Time (ms):";
                                ScorecardSheet.Cells["B43"].Style.Font.Bold = true;

                                //styling for total range results
                                var TotalRangeResults = ScorecardSheet.Cells["A37:C43"];
                                TotalRangeResults.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                TotalRangeResults.Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#F2F2F2"));

                                //snappy range results
                                var SnappyRangeTag = ScorecardSheet.Cells["A44:A50"];
                                SnappyRangeTag.Merge = true;
                                SnappyRangeTag.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                SnappyRangeTag.Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                                ScorecardSheet.Cells["A44"].Value = "Snappy Range";
                                ScorecardSheet.Cells["A44"].Style.TextRotation = 45;

                                ScorecardSheet.Cells["B44"].Value = "Snappy Minimum Distance (in):";
                                ScorecardSheet.Cells["B44"].Style.Font.Bold = true;

                                ScorecardSheet.Cells["B45"].Value = "Snappy Maximum Distance (in):";
                                ScorecardSheet.Cells["B45"].Style.Font.Bold = true;

                                ScorecardSheet.Cells["B46"].Value = "Snappy Decode Range (in):";
                                ScorecardSheet.Cells["B46"].Style.Font.Bold = true;

                                ScorecardSheet.Cells["B47"].Value = "Snappy Decode Time Average (ms):";
                                ScorecardSheet.Cells["B47"].Style.Font.Bold = true;

                                ScorecardSheet.Cells["B48"].Value = "Snappy Std Dev (ms):";
                                ScorecardSheet.Cells["B48"].Style.Font.Bold = true;

                                ScorecardSheet.Cells["B49"].Value = "Snappy Highest Decode Time (ms)";
                                ScorecardSheet.Cells["B49"].Style.Font.Bold = true;

                                ScorecardSheet.Cells["B50"].Value = "Snappy Lowest Decode Time (ms):";
                                ScorecardSheet.Cells["B50"].Style.Font.Bold = true;

                                //styling for snappy range results
                                var SnappyRangeResults = ScorecardSheet.Cells["A44:C50"];
                                SnappyRangeResults.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                SnappyRangeResults.Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D9D9D9"));
                            }

                            WriteLogDataToScorecard(ScorecardSheet, log, log.ScorecardGroupNumber);
                        }
                        else if (log.ScorecardGroupNumber == "3")
                        {
                            if (ScorecardSheet.Cells["C53"].Value == null)
                            {
                                //total range results
                                var TotalRangeTag = ScorecardSheet.Cells["A57:A63"];
                                TotalRangeTag.Merge = true;
                                TotalRangeTag.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                TotalRangeTag.Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                                ScorecardSheet.Cells["A57"].Value = "Total Range";
                                ScorecardSheet.Cells["A57"].Style.TextRotation = 45;

                                ScorecardSheet.Cells["B57"].Value = "Minimum Distance (in):";
                                ScorecardSheet.Cells["B57"].Style.Font.Bold = true;

                                ScorecardSheet.Cells["B58"].Value = "Maximum Distance (in):";
                                ScorecardSheet.Cells["B58"].Style.Font.Bold = true;

                                ScorecardSheet.Cells["B59"].Value = "Decode Range (in):";
                                ScorecardSheet.Cells["B59"].Style.Font.Bold = true;

                                ScorecardSheet.Cells["B60"].Value = "Decode Time Average (ms):";
                                ScorecardSheet.Cells["B60"].Style.Font.Bold = true;

                                ScorecardSheet.Cells["B61"].Value = "Std Dev (ms):";
                                ScorecardSheet.Cells["B61"].Style.Font.Bold = true;

                                ScorecardSheet.Cells["B62"].Value = "Highest Decode Time (ms)";
                                ScorecardSheet.Cells["B62"].Style.Font.Bold = true;

                                ScorecardSheet.Cells["B63"].Value = "Lowest Decode Time (ms):";
                                ScorecardSheet.Cells["B63"].Style.Font.Bold = true;

                                //styling for total range results
                                var TotalRangeResults = ScorecardSheet.Cells["A57:C63"];
                                TotalRangeResults.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                TotalRangeResults.Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#F2F2F2"));

                                //snappy range results
                                var SnappyRangeTag = ScorecardSheet.Cells["A64:A70"];
                                SnappyRangeTag.Merge = true;
                                SnappyRangeTag.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                SnappyRangeTag.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                ScorecardSheet.Cells["A64"].Value = "Snappy Range";
                                ScorecardSheet.Cells["A64"].Style.TextRotation = 45;

                                ScorecardSheet.Cells["B64"].Value = "Snappy Minimum Distance (in):";
                                ScorecardSheet.Cells["B64"].Style.Font.Bold = true;

                                ScorecardSheet.Cells["B65"].Value = "Snappy Maximum Distance (in):";
                                ScorecardSheet.Cells["B65"].Style.Font.Bold = true;

                                ScorecardSheet.Cells["B66"].Value = "Snappy Decode Range (in):";
                                ScorecardSheet.Cells["B66"].Style.Font.Bold = true;

                                ScorecardSheet.Cells["B67"].Value = "Snappy Decode Time Average (ms):";
                                ScorecardSheet.Cells["B67"].Style.Font.Bold = true;

                                ScorecardSheet.Cells["B68"].Value = "Snappy Std Dev (ms):";
                                ScorecardSheet.Cells["B68"].Style.Font.Bold = true;

                                ScorecardSheet.Cells["B69"].Value = "Snappy Highest Decode Time (ms)";
                                ScorecardSheet.Cells["B69"].Style.Font.Bold = true;

                                ScorecardSheet.Cells["B70"].Value = "Snappy Lowest Decode Time (ms):";
                                ScorecardSheet.Cells["B70"].Style.Font.Bold = true;

                                //styling for snappy range results
                                var SnappyRangeResults = ScorecardSheet.Cells["A64:C70"];
                                SnappyRangeResults.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                SnappyRangeResults.Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D9D9D9"));
                            }

                            WriteLogDataToScorecard(ScorecardSheet, log, log.ScorecardGroupNumber);
                        }
                        else if (log.ScorecardGroupNumber == "4")
                        {
                            if (ScorecardSheet.Cells["C73"].Value == null)
                            {
                                //total range results
                                var TotalRangeTag = ScorecardSheet.Cells["A77:A83"];
                                TotalRangeTag.Merge = true;
                                TotalRangeTag.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                TotalRangeTag.Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                                ScorecardSheet.Cells["A77"].Value = "Total Range";
                                ScorecardSheet.Cells["A77"].Style.TextRotation = 45;

                                ScorecardSheet.Cells["B77"].Value = "Minimum Distance (in):";
                                ScorecardSheet.Cells["B77"].Style.Font.Bold = true;

                                ScorecardSheet.Cells["B78"].Value = "Maximum Distance (in):";
                                ScorecardSheet.Cells["B78"].Style.Font.Bold = true;

                                ScorecardSheet.Cells["B79"].Value = "Decode Range (in):";
                                ScorecardSheet.Cells["B79"].Style.Font.Bold = true;

                                ScorecardSheet.Cells["B80"].Value = "Decode Time Average (ms):";
                                ScorecardSheet.Cells["B80"].Style.Font.Bold = true;

                                ScorecardSheet.Cells["B81"].Value = "Std Dev (ms):";
                                ScorecardSheet.Cells["B81"].Style.Font.Bold = true;

                                ScorecardSheet.Cells["B82"].Value = "Highest Decode Time (ms)";
                                ScorecardSheet.Cells["B82"].Style.Font.Bold = true;

                                ScorecardSheet.Cells["B83"].Value = "Lowest Decode Time (ms):";
                                ScorecardSheet.Cells["B83"].Style.Font.Bold = true;

                                //styling for total range results
                                var TotalRangeResults = ScorecardSheet.Cells["A77:C83"];
                                TotalRangeResults.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                TotalRangeResults.Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#F2F2F2"));

                                //snappy range results
                                var SnappyRangeTag = ScorecardSheet.Cells["A84:A90"];
                                SnappyRangeTag.Merge = true;
                                SnappyRangeTag.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                SnappyRangeTag.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                ScorecardSheet.Cells["A84"].Value = "Snappy Range";
                                ScorecardSheet.Cells["A84"].Style.TextRotation = 45;

                                ScorecardSheet.Cells["B84"].Value = "Snappy Minimum Distance (in):";
                                ScorecardSheet.Cells["B84"].Style.Font.Bold = true;

                                ScorecardSheet.Cells["B85"].Value = "Snappy Maximum Distance (in):";
                                ScorecardSheet.Cells["B85"].Style.Font.Bold = true;

                                ScorecardSheet.Cells["B86"].Value = "Snappy Decode Range (in):";
                                ScorecardSheet.Cells["B86"].Style.Font.Bold = true;

                                ScorecardSheet.Cells["B87"].Value = "Snappy Decode Time Average (ms):";
                                ScorecardSheet.Cells["B87"].Style.Font.Bold = true;

                                ScorecardSheet.Cells["B88"].Value = "Snappy Std Dev (ms):";
                                ScorecardSheet.Cells["B88"].Style.Font.Bold = true;

                                ScorecardSheet.Cells["B89"].Value = "Snappy Highest Decode Time (ms)";
                                ScorecardSheet.Cells["B89"].Style.Font.Bold = true;

                                ScorecardSheet.Cells["B90"].Value = "Lowest Decode Time (ms):";
                                ScorecardSheet.Cells["B90"].Style.Font.Bold = true;

                                //styling for snappy range results
                                var SnappyRangeResults = ScorecardSheet.Cells["A84:C90"];
                                SnappyRangeResults.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                SnappyRangeResults.Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D9D9D9"));
                            }

                            WriteLogDataToScorecard(ScorecardSheet, log, log.ScorecardGroupNumber);
                        }
                        else if (log.ScorecardGroupNumber == "5")
                        {
                            if (ScorecardSheet.Cells["C93"].Value == null)
                            {
                                //total range results
                                var TotalRangeTag = ScorecardSheet.Cells["A97:A103"];
                                TotalRangeTag.Merge = true;
                                TotalRangeTag.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                TotalRangeTag.Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                                ScorecardSheet.Cells["A97"].Value = "Total Range";
                                ScorecardSheet.Cells["A97"].Style.TextRotation = 45;

                                ScorecardSheet.Cells["B97"].Value = "Minimum Distance (in):";
                                ScorecardSheet.Cells["B97"].Style.Font.Bold = true;

                                ScorecardSheet.Cells["B98"].Value = "Maximum Distance (in):";
                                ScorecardSheet.Cells["B98"].Style.Font.Bold = true;

                                ScorecardSheet.Cells["B99"].Value = "Decode Range (in):";
                                ScorecardSheet.Cells["B99"].Style.Font.Bold = true;

                                ScorecardSheet.Cells["B100"].Value = "Decode Time Average (ms):";
                                ScorecardSheet.Cells["B100"].Style.Font.Bold = true;

                                ScorecardSheet.Cells["B101"].Value = "Std Dev (ms):";
                                ScorecardSheet.Cells["B101"].Style.Font.Bold = true;

                                ScorecardSheet.Cells["B102"].Value = "Highest Decode Time (ms)";
                                ScorecardSheet.Cells["B102"].Style.Font.Bold = true;

                                ScorecardSheet.Cells["B103"].Value = "Lowest Decode Time (ms):";
                                ScorecardSheet.Cells["B103"].Style.Font.Bold = true;

                                //styling for total range results
                                var TotalRangeResults = ScorecardSheet.Cells["A97:C103"];
                                TotalRangeResults.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                TotalRangeResults.Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#F2F2F2"));

                                //snappy range results
                                var SnappyRangeTag = ScorecardSheet.Cells["A104:A110"];
                                SnappyRangeTag.Merge = true;
                                SnappyRangeTag.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                SnappyRangeTag.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                ScorecardSheet.Cells["A104"].Value = "Snappy Range";
                                ScorecardSheet.Cells["A104"].Style.TextRotation = 45;

                                ScorecardSheet.Cells["B104"].Value = "Snappy Minimum Distance (in):";
                                ScorecardSheet.Cells["B104"].Style.Font.Bold = true;

                                ScorecardSheet.Cells["B105"].Value = "Snappy Maximum Distance (in):";
                                ScorecardSheet.Cells["B105"].Style.Font.Bold = true;

                                ScorecardSheet.Cells["B106"].Value = "Snappy Decode Range (in):";
                                ScorecardSheet.Cells["B106"].Style.Font.Bold = true;

                                ScorecardSheet.Cells["B107"].Value = "Snappy Decode Time Average (ms):";
                                ScorecardSheet.Cells["B107"].Style.Font.Bold = true;

                                ScorecardSheet.Cells["B108"].Value = "Snappy Std Dev (ms):";
                                ScorecardSheet.Cells["B108"].Style.Font.Bold = true;

                                ScorecardSheet.Cells["B109"].Value = "Snappy Highest Decode Time (ms)";
                                ScorecardSheet.Cells["B109"].Style.Font.Bold = true;

                                ScorecardSheet.Cells["B110"].Value = "Lowest Decode Time (ms):";
                                ScorecardSheet.Cells["B110"].Style.Font.Bold = true;

                                //styling for snappy range results
                                var SnappyRangeResults = ScorecardSheet.Cells["A104:C110"];
                                SnappyRangeResults.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                SnappyRangeResults.Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D9D9D9"));
                            }

                            WriteLogDataToScorecard(ScorecardSheet, log, log.ScorecardGroupNumber);
                        }
                    }

                    //handles how the user wants to compare the data
                    //HighlightBestResultsHorizontal(ScorecardSheet);
                    string HowToCompare = HowToCompareResultsComboBox.SelectedItem.ToString();

                    switch (HowToCompare)
                    {
                        case "All By Row (Left to Right)":
                            HighlightBestResultsHorizontal(ScorecardSheet);
                            break;

                        case "All By Column (Top to Bottom)":
                            HighlightBestResultsVertical(ScorecardSheet);
                            break;

                        case "Decode Time Average Only by Column (Top to Bottom)":
                            HighlightBestAverageDecodeTimeResultsVertical(ScorecardSheet);
                            break;

                        case "Decode Time Average Only by Row (Left to Right)":
                            HighlightBestAverageDecodeTimeResultsHorizontal(ScorecardSheet);
                            break;

                        case "Do Not Compare (No Highlighting)":
                            break;
                    }

                    //move scorecard worksheet to the beginning
                    excelPackage.Workbook.Worksheets.MoveToStart(ScorecardSheet.Name);
                }
                catch (Exception)
                {
                    ErrorMessageTextBox.Text = "There was an error creating the scorecard tab. Did you select a propper log file to compile? Did you select the same log twice?";
                }

                //create the spreadsheet file and save it to a new folder on the users desktop
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
        /// <summary>
        /// Writes a summary for each log file to the scorecard tab
        /// </summary>
        /// <param name="ScorecardSheet"></param>
        /// <param name="Log"></param>
        /// <param name="GroupNumber"></param>
        private void WriteLogDataToScorecard(ExcelWorksheet ScorecardSheet, TargetLog Log, string GroupNumber)
        {
            //Prints data to CurrentRow 1 of the scorecard tab
            if (GroupNumber == "1")
            {
                int Row1 = 13;
                int Column1 = 3;

                while (true)
                {
                    var CurrentCell = ScorecardSheet.Cells[Row1, Column1];
                    string CurrentCellValue = CurrentCell.Text;

                    if (string.IsNullOrEmpty(CurrentCellValue))
                    {
                        //Prints sample name
                        CurrentCell.Value = Log.BarcodeSampleName;
                        CurrentCell.Style.Font.Bold = true;
                        CurrentCell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        CurrentCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        CurrentCell.Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#FFF2CC"));

                        //prints scanner name
                        CurrentCell.Offset(1, 0).Value = Log.ScannerMake + "_" + Log.ScannerName;
                        CurrentCell.Offset(1, 0).Style.Font.Bold = true;
                        CurrentCell.Offset(1, 0).Style.Border.Top.Style = ExcelBorderStyle.Thick;
                        CurrentCell.Offset(1, 0).Style.Border.Right.Style = ExcelBorderStyle.Thick;
                        CurrentCell.Offset(1, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        CurrentCell.Offset(1, 0).Style.Fill.PatternType = ExcelFillStyle.Solid;
                        CurrentCell.Offset(1, 0).Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#DDEBF7"));

                        //Notes field for the user to enter in any relevant notes about the scanner
                        CurrentCell.Offset(2, 0).Value = "Notes:";
                        CurrentCell.Offset(2, 0).Style.Font.Bold = true;
                        CurrentCell.Offset(2, 0).Style.Border.Right.Style = ExcelBorderStyle.Thick;
                        CurrentCell.Offset(2, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        CurrentCell.Offset(2, 0).Style.Fill.PatternType = ExcelFillStyle.Solid;
                        CurrentCell.Offset(2, 0).Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#C5D9F1"));
                        ScorecardSheet.Row(16).Height = 75;
                        CurrentCell.Offset(3, 0).Style.Border.Right.Style = ExcelBorderStyle.Thick;
                        CurrentCell.Offset(3, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        CurrentCell.Offset(3, 0).Style.Fill.PatternType = ExcelFillStyle.Solid;
                        CurrentCell.Offset(3, 0).Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#C5D9F1"));

                        //total range results
                        CurrentCell.Offset(4, 0).Value = CalculateClosestRange(Log.SummaryGraphData);
                        CurrentCell.Offset(4, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        CurrentCell.Offset(4, 0).Style.Border.Right.Style = ExcelBorderStyle.Thick;
                        CurrentCell.Offset(4, 0).Style.Numberformat.Format = "0.00";
                        CurrentCell.Offset(4, 0).Style.Fill.PatternType = ExcelFillStyle.Solid;
                        CurrentCell.Offset(4, 0).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#F2F2F2"));

                        CurrentCell.Offset(5, 0).Value = CalculateFarthestRange(Log.SummaryGraphData);
                        CurrentCell.Offset(5, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        CurrentCell.Offset(5, 0).Style.Border.Right.Style = ExcelBorderStyle.Thick;
                        CurrentCell.Offset(5, 0).Style.Numberformat.Format = "0.00";
                        CurrentCell.Offset(5, 0).Style.Fill.PatternType = ExcelFillStyle.Solid;
                        CurrentCell.Offset(5, 0).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#F2F2F2"));

                        CurrentCell.Offset(6, 0).Value = CalculateDecodeRange(Log.SummaryGraphData);
                        CurrentCell.Offset(6, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        CurrentCell.Offset(6, 0).Style.Border.Right.Style = ExcelBorderStyle.Thick;
                        CurrentCell.Offset(6, 0).Style.Numberformat.Format = "0.00";
                        CurrentCell.Offset(6, 0).Style.Fill.PatternType = ExcelFillStyle.Solid;
                        CurrentCell.Offset(6, 0).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#F2F2F2"));

                        CurrentCell.Offset(7, 0).Value = CalculateAverageDecodeTime(Log.SummaryGraphData);
                        CurrentCell.Offset(7, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        CurrentCell.Offset(7, 0).Style.Border.Right.Style = ExcelBorderStyle.Thick;
                        CurrentCell.Offset(7, 0).Style.Numberformat.Format = "0.00";
                        CurrentCell.Offset(7, 0).Style.Fill.PatternType = ExcelFillStyle.Solid;
                        CurrentCell.Offset(7, 0).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#F2F2F2"));

                        //CurrentCell.Offset(8, 0).Value = CalculateStandardDeviationDecodeTime(Log.DecodeTimes);
                        CurrentCell.Offset(8, 0).Value = CalculateStandardDeviationDecodeTimeNEW(Log.SummaryGraphData);
                        CurrentCell.Offset(8, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        CurrentCell.Offset(8, 0).Style.Border.Right.Style = ExcelBorderStyle.Thick;
                        CurrentCell.Offset(8, 0).Style.Numberformat.Format = "0.00";
                        CurrentCell.Offset(8, 0).Style.Fill.PatternType = ExcelFillStyle.Solid;
                        CurrentCell.Offset(8, 0).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#F2F2F2"));

                        //CurrentCell.Offset(9, 0).Value = CalculateHighestDecodeTime(Log.DecodeTimes);
                        CurrentCell.Offset(9, 0).Value = CalculateHighestDecodeTimeNEW(Log.SummaryGraphData);
                        CurrentCell.Offset(9, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        CurrentCell.Offset(9, 0).Style.Border.Right.Style = ExcelBorderStyle.Thick;
                        CurrentCell.Offset(9, 0).Style.Numberformat.Format = "0.00";
                        CurrentCell.Offset(9, 0).Style.Fill.PatternType = ExcelFillStyle.Solid;
                        CurrentCell.Offset(9, 0).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#F2F2F2"));

                        //CurrentCell.Offset(10, 0).Value = CalculateLowestDecodeTime(Log.DecodeTimes);
                        CurrentCell.Offset(10, 0).Value = CalculateLowestDecodeTimeNEW(Log.SummaryGraphData);
                        CurrentCell.Offset(10, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        CurrentCell.Offset(10, 0).Style.Border.Right.Style = ExcelBorderStyle.Thick;
                        CurrentCell.Offset(10, 0).Style.Numberformat.Format = "0.00";
                        CurrentCell.Offset(10, 0).Style.Fill.PatternType = ExcelFillStyle.Solid;
                        CurrentCell.Offset(10, 0).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#F2F2F2"));

                        //snappy range results
                        try
                        {
                            CurrentCell.Offset(11, 0).Value = CalculateClosestRange(Log.SnappyData);
                            CurrentCell.Offset(11, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            CurrentCell.Offset(11, 0).Style.Border.Right.Style = ExcelBorderStyle.Thick;
                            CurrentCell.Offset(11, 0).Style.Numberformat.Format = "0.00";
                            CurrentCell.Offset(11, 0).Style.Fill.PatternType = ExcelFillStyle.Solid;
                            CurrentCell.Offset(11, 0).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D9D9D9"));

                            CurrentCell.Offset(12, 0).Value = CalculateFarthestRange(Log.SnappyData);
                            CurrentCell.Offset(12, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            CurrentCell.Offset(12, 0).Style.Border.Right.Style = ExcelBorderStyle.Thick;
                            CurrentCell.Offset(12, 0).Style.Numberformat.Format = "0.00";
                            CurrentCell.Offset(12, 0).Style.Fill.PatternType = ExcelFillStyle.Solid;
                            CurrentCell.Offset(12, 0).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D9D9D9"));

                            CurrentCell.Offset(13, 0).Value = CalculateDecodeRange(Log.SnappyData);
                            CurrentCell.Offset(13, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            CurrentCell.Offset(13, 0).Style.Border.Right.Style = ExcelBorderStyle.Thick;
                            CurrentCell.Offset(13, 0).Style.Numberformat.Format = "0.00";
                            CurrentCell.Offset(13, 0).Style.Fill.PatternType = ExcelFillStyle.Solid;
                            CurrentCell.Offset(13, 0).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D9D9D9"));

                            CurrentCell.Offset(14, 0).Value = CalculateAverageDecodeTime(Log.SnappyData);
                            CurrentCell.Offset(14, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            CurrentCell.Offset(14, 0).Style.Border.Right.Style = ExcelBorderStyle.Thick;
                            CurrentCell.Offset(14, 0).Style.Numberformat.Format = "0.00";
                            CurrentCell.Offset(14, 0).Style.Fill.PatternType = ExcelFillStyle.Solid;
                            CurrentCell.Offset(14, 0).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D9D9D9"));

                            //CurrentCell.Offset(15, 0).Value = CalculateStandardDeviationDecodeTime(Log.SnappyDecodeTimes);
                            CurrentCell.Offset(15, 0).Value = CalculateStandardDeviationDecodeTimeNEW(Log.SnappyData);
                            CurrentCell.Offset(15, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            CurrentCell.Offset(15, 0).Style.Border.Right.Style = ExcelBorderStyle.Thick;
                            CurrentCell.Offset(15, 0).Style.Numberformat.Format = "0.00";
                            CurrentCell.Offset(15, 0).Style.Fill.PatternType = ExcelFillStyle.Solid;
                            CurrentCell.Offset(15, 0).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D9D9D9"));

                            //CurrentCell.Offset(16, 0).Value = CalculateHighestDecodeTime(Log.SnappyDecodeTimes);
                            CurrentCell.Offset(16, 0).Value = CalculateHighestDecodeTimeNEW(Log.SnappyData);
                            CurrentCell.Offset(16, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            CurrentCell.Offset(16, 0).Style.Border.Right.Style = ExcelBorderStyle.Thick;
                            CurrentCell.Offset(16, 0).Style.Numberformat.Format = "0.00";
                            CurrentCell.Offset(16, 0).Style.Fill.PatternType = ExcelFillStyle.Solid;
                            CurrentCell.Offset(16, 0).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D9D9D9"));

                            //CurrentCell.Offset(17, 0).Value = CalculateLowestDecodeTime(Log.SnappyDecodeTimes);
                            CurrentCell.Offset(17, 0).Value = CalculateLowestDecodeTimeNEW(Log.SnappyData);
                            CurrentCell.Offset(17, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            CurrentCell.Offset(17, 0).Style.Border.Right.Style = ExcelBorderStyle.Thick;
                            CurrentCell.Offset(17, 0).Style.Numberformat.Format = "0.00";
                            CurrentCell.Offset(17, 0).Style.Fill.PatternType = ExcelFillStyle.Solid;
                            CurrentCell.Offset(17, 0).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D9D9D9"));
                        }
                        catch (Exception)
                        {
                            //Issue detected with getting snappy data, use regular times
                            CurrentCell.Offset(11, 0).Value = CalculateClosestRange(Log.SummaryGraphData);
                            CurrentCell.Offset(11, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            CurrentCell.Offset(11, 0).Style.Border.Right.Style = ExcelBorderStyle.Thick;
                            CurrentCell.Offset(11, 0).Style.Numberformat.Format = "0.00";
                            CurrentCell.Offset(11, 0).Style.Fill.PatternType = ExcelFillStyle.Solid;
                            CurrentCell.Offset(11, 0).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D9D9D9"));

                            CurrentCell.Offset(12, 0).Value = CalculateFarthestRange(Log.SummaryGraphData);
                            CurrentCell.Offset(12, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            CurrentCell.Offset(12, 0).Style.Border.Right.Style = ExcelBorderStyle.Thick;
                            CurrentCell.Offset(12, 0).Style.Numberformat.Format = "0.00";
                            CurrentCell.Offset(12, 0).Style.Fill.PatternType = ExcelFillStyle.Solid;
                            CurrentCell.Offset(12, 0).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D9D9D9"));

                            CurrentCell.Offset(13, 0).Value = CalculateDecodeRange(Log.SummaryGraphData);
                            CurrentCell.Offset(13, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            CurrentCell.Offset(13, 0).Style.Border.Right.Style = ExcelBorderStyle.Thick;
                            CurrentCell.Offset(13, 0).Style.Numberformat.Format = "0.00";
                            CurrentCell.Offset(13, 0).Style.Fill.PatternType = ExcelFillStyle.Solid;
                            CurrentCell.Offset(13, 0).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D9D9D9"));

                            CurrentCell.Offset(14, 0).Value = CalculateAverageDecodeTime(Log.SummaryGraphData);
                            CurrentCell.Offset(14, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            CurrentCell.Offset(14, 0).Style.Border.Right.Style = ExcelBorderStyle.Thick;
                            CurrentCell.Offset(14, 0).Style.Numberformat.Format = "0.00";
                            CurrentCell.Offset(14, 0).Style.Fill.PatternType = ExcelFillStyle.Solid;
                            CurrentCell.Offset(14, 0).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D9D9D9"));

                            //CurrentCell.Offset(15, 0).Value = CalculateStandardDeviationDecodeTime(Log.DecodeTimes);
                            CurrentCell.Offset(15, 0).Value = CalculateStandardDeviationDecodeTimeNEW(Log.SummaryGraphData);
                            CurrentCell.Offset(15, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            CurrentCell.Offset(15, 0).Style.Border.Right.Style = ExcelBorderStyle.Thick;
                            CurrentCell.Offset(15, 0).Style.Numberformat.Format = "0.00";
                            CurrentCell.Offset(15, 0).Style.Fill.PatternType = ExcelFillStyle.Solid;
                            CurrentCell.Offset(15, 0).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D9D9D9"));

                            //CurrentCell.Offset(16, 0).Value = CalculateHighestDecodeTime(Log.DecodeTimes);
                            CurrentCell.Offset(16, 0).Value = CalculateHighestDecodeTimeNEW(Log.SummaryGraphData);
                            CurrentCell.Offset(16, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            CurrentCell.Offset(16, 0).Style.Border.Right.Style = ExcelBorderStyle.Thick;
                            CurrentCell.Offset(16, 0).Style.Numberformat.Format = "0.00";
                            CurrentCell.Offset(16, 0).Style.Fill.PatternType = ExcelFillStyle.Solid;
                            CurrentCell.Offset(16, 0).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D9D9D9"));

                            //CurrentCell.Offset(17, 0).Value = CalculateLowestDecodeTime(Log.DecodeTimes);
                            CurrentCell.Offset(17, 0).Value = CalculateLowestDecodeTimeNEW(Log.SummaryGraphData);
                            CurrentCell.Offset(17, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            CurrentCell.Offset(17, 0).Style.Border.Right.Style = ExcelBorderStyle.Thick;
                            CurrentCell.Offset(17, 0).Style.Numberformat.Format = "0.00";
                            CurrentCell.Offset(17, 0).Style.Fill.PatternType = ExcelFillStyle.Solid;
                            CurrentCell.Offset(17, 0).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D9D9D9"));
                        }

                        break;
                    }
                    else
                    {
                        Column1++;
                    }
                }
            }

            //Prints data to CurrentRow 2 of the scorecard tab
            if (GroupNumber == "2")
            {
                int Row1 = 33;
                int Column1 = 3;

                while (true)
                {
                    var CurrentCell = ScorecardSheet.Cells[Row1, Column1];
                    string CurrentCellValue = CurrentCell.Text;

                    if (string.IsNullOrEmpty(CurrentCellValue))
                    {
                        //Prints sample name
                        CurrentCell.Value = Log.BarcodeSampleName;
                        CurrentCell.Style.Font.Bold = true;
                        CurrentCell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        CurrentCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        CurrentCell.Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#FFF2CC"));

                        //prints scanner name
                        CurrentCell.Offset(1, 0).Value = Log.ScannerMake + "_" + Log.ScannerName;
                        CurrentCell.Offset(1, 0).Style.Font.Bold = true;
                        CurrentCell.Offset(1, 0).Style.Border.Top.Style = ExcelBorderStyle.Thick;
                        CurrentCell.Offset(1, 0).Style.Border.Right.Style = ExcelBorderStyle.Thick;
                        CurrentCell.Offset(1, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        CurrentCell.Offset(1, 0).Style.Fill.PatternType = ExcelFillStyle.Solid;
                        CurrentCell.Offset(1, 0).Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#DDEBF7"));

                        //Notes field for the user to enter in any relevant notes about the scanner
                        CurrentCell.Offset(2, 0).Value = "Notes:";
                        CurrentCell.Offset(2, 0).Style.Font.Bold = true;
                        CurrentCell.Offset(2, 0).Style.Border.Right.Style = ExcelBorderStyle.Thick;
                        CurrentCell.Offset(2, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        CurrentCell.Offset(2, 0).Style.Fill.PatternType = ExcelFillStyle.Solid;
                        CurrentCell.Offset(2, 0).Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#C5D9F1"));
                        ScorecardSheet.Row(16).Height = 75;
                        CurrentCell.Offset(3, 0).Style.Border.Right.Style = ExcelBorderStyle.Thick;
                        CurrentCell.Offset(3, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        CurrentCell.Offset(3, 0).Style.Fill.PatternType = ExcelFillStyle.Solid;
                        CurrentCell.Offset(3, 0).Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#C5D9F1"));

                        //total range results
                        CurrentCell.Offset(4, 0).Value = CalculateClosestRange(Log.SummaryGraphData);
                        CurrentCell.Offset(4, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        CurrentCell.Offset(4, 0).Style.Border.Right.Style = ExcelBorderStyle.Thick;
                        CurrentCell.Offset(4, 0).Style.Numberformat.Format = "0.00";
                        CurrentCell.Offset(4, 0).Style.Fill.PatternType = ExcelFillStyle.Solid;
                        CurrentCell.Offset(4, 0).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#F2F2F2"));

                        CurrentCell.Offset(5, 0).Value = CalculateFarthestRange(Log.SummaryGraphData);
                        CurrentCell.Offset(5, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        CurrentCell.Offset(5, 0).Style.Border.Right.Style = ExcelBorderStyle.Thick;
                        CurrentCell.Offset(5, 0).Style.Numberformat.Format = "0.00";
                        CurrentCell.Offset(5, 0).Style.Fill.PatternType = ExcelFillStyle.Solid;
                        CurrentCell.Offset(5, 0).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#F2F2F2"));

                        CurrentCell.Offset(6, 0).Value = CalculateDecodeRange(Log.SummaryGraphData);
                        CurrentCell.Offset(6, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        CurrentCell.Offset(6, 0).Style.Border.Right.Style = ExcelBorderStyle.Thick;
                        CurrentCell.Offset(6, 0).Style.Numberformat.Format = "0.00";
                        CurrentCell.Offset(6, 0).Style.Fill.PatternType = ExcelFillStyle.Solid;
                        CurrentCell.Offset(6, 0).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#F2F2F2"));

                        CurrentCell.Offset(7, 0).Value = CalculateAverageDecodeTime(Log.SummaryGraphData);
                        CurrentCell.Offset(7, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        CurrentCell.Offset(7, 0).Style.Border.Right.Style = ExcelBorderStyle.Thick;
                        CurrentCell.Offset(7, 0).Style.Numberformat.Format = "0.00";
                        CurrentCell.Offset(7, 0).Style.Fill.PatternType = ExcelFillStyle.Solid;
                        CurrentCell.Offset(7, 0).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#F2F2F2"));

                        //CurrentCell.Offset(8, 0).Value = CalculateStandardDeviationDecodeTime(Log.DecodeTimes);
                        CurrentCell.Offset(8, 0).Value = CalculateStandardDeviationDecodeTimeNEW(Log.SummaryGraphData);
                        CurrentCell.Offset(8, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        CurrentCell.Offset(8, 0).Style.Border.Right.Style = ExcelBorderStyle.Thick;
                        CurrentCell.Offset(8, 0).Style.Numberformat.Format = "0.00";
                        CurrentCell.Offset(8, 0).Style.Fill.PatternType = ExcelFillStyle.Solid;
                        CurrentCell.Offset(8, 0).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#F2F2F2"));

                        //CurrentCell.Offset(9, 0).Value = CalculateHighestDecodeTime(Log.DecodeTimes);
                        CurrentCell.Offset(9, 0).Value = CalculateHighestDecodeTimeNEW(Log.SummaryGraphData);
                        CurrentCell.Offset(9, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        CurrentCell.Offset(9, 0).Style.Border.Right.Style = ExcelBorderStyle.Thick;
                        CurrentCell.Offset(9, 0).Style.Numberformat.Format = "0.00";
                        CurrentCell.Offset(9, 0).Style.Fill.PatternType = ExcelFillStyle.Solid;
                        CurrentCell.Offset(9, 0).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#F2F2F2"));

                        //CurrentCell.Offset(10, 0).Value = CalculateLowestDecodeTime(Log.DecodeTimes);
                        CurrentCell.Offset(10, 0).Value = CalculateLowestDecodeTimeNEW(Log.SummaryGraphData);
                        CurrentCell.Offset(10, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        CurrentCell.Offset(10, 0).Style.Border.Right.Style = ExcelBorderStyle.Thick;
                        CurrentCell.Offset(10, 0).Style.Numberformat.Format = "0.00";
                        CurrentCell.Offset(10, 0).Style.Fill.PatternType = ExcelFillStyle.Solid;
                        CurrentCell.Offset(10, 0).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#F2F2F2"));

                        //snappy range results
                        try
                        {
                            CurrentCell.Offset(11, 0).Value = CalculateClosestRange(Log.SnappyData);
                            CurrentCell.Offset(11, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            CurrentCell.Offset(11, 0).Style.Border.Right.Style = ExcelBorderStyle.Thick;
                            CurrentCell.Offset(11, 0).Style.Numberformat.Format = "0.00";
                            CurrentCell.Offset(11, 0).Style.Fill.PatternType = ExcelFillStyle.Solid;
                            CurrentCell.Offset(11, 0).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D9D9D9"));

                            CurrentCell.Offset(12, 0).Value = CalculateFarthestRange(Log.SnappyData);
                            CurrentCell.Offset(12, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            CurrentCell.Offset(12, 0).Style.Border.Right.Style = ExcelBorderStyle.Thick;
                            CurrentCell.Offset(12, 0).Style.Numberformat.Format = "0.00";
                            CurrentCell.Offset(12, 0).Style.Fill.PatternType = ExcelFillStyle.Solid;
                            CurrentCell.Offset(12, 0).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D9D9D9"));

                            CurrentCell.Offset(13, 0).Value = CalculateDecodeRange(Log.SnappyData);
                            CurrentCell.Offset(13, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            CurrentCell.Offset(13, 0).Style.Border.Right.Style = ExcelBorderStyle.Thick;
                            CurrentCell.Offset(13, 0).Style.Numberformat.Format = "0.00";
                            CurrentCell.Offset(13, 0).Style.Fill.PatternType = ExcelFillStyle.Solid;
                            CurrentCell.Offset(13, 0).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D9D9D9"));

                            CurrentCell.Offset(14, 0).Value = CalculateAverageDecodeTime(Log.SnappyData);
                            CurrentCell.Offset(14, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            CurrentCell.Offset(14, 0).Style.Border.Right.Style = ExcelBorderStyle.Thick;
                            CurrentCell.Offset(14, 0).Style.Numberformat.Format = "0.00";
                            CurrentCell.Offset(14, 0).Style.Fill.PatternType = ExcelFillStyle.Solid;
                            CurrentCell.Offset(14, 0).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D9D9D9"));

                            //CurrentCell.Offset(15, 0).Value = CalculateStandardDeviationDecodeTime(Log.DecodeTimes);
                            CurrentCell.Offset(15, 0).Value = CalculateStandardDeviationDecodeTimeNEW(Log.SnappyData);
                            CurrentCell.Offset(15, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            CurrentCell.Offset(15, 0).Style.Border.Right.Style = ExcelBorderStyle.Thick;
                            CurrentCell.Offset(15, 0).Style.Numberformat.Format = "0.00";
                            CurrentCell.Offset(15, 0).Style.Fill.PatternType = ExcelFillStyle.Solid;
                            CurrentCell.Offset(15, 0).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D9D9D9"));

                            //CurrentCell.Offset(16, 0).Value = CalculateHighestDecodeTime(Log.DecodeTimes);
                            CurrentCell.Offset(16, 0).Value = CalculateHighestDecodeTimeNEW(Log.SnappyData);
                            CurrentCell.Offset(16, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            CurrentCell.Offset(16, 0).Style.Border.Right.Style = ExcelBorderStyle.Thick;
                            CurrentCell.Offset(16, 0).Style.Numberformat.Format = "0.00";
                            CurrentCell.Offset(16, 0).Style.Fill.PatternType = ExcelFillStyle.Solid;
                            CurrentCell.Offset(16, 0).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D9D9D9"));

                            //CurrentCell.Offset(17, 0).Value = CalculateLowestDecodeTime(Log.DecodeTimes);
                            CurrentCell.Offset(17, 0).Value = CalculateLowestDecodeTimeNEW(Log.SnappyData);
                            CurrentCell.Offset(17, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            CurrentCell.Offset(17, 0).Style.Border.Right.Style = ExcelBorderStyle.Thick;
                            CurrentCell.Offset(17, 0).Style.Numberformat.Format = "0.00";
                            CurrentCell.Offset(17, 0).Style.Fill.PatternType = ExcelFillStyle.Solid;
                            CurrentCell.Offset(17, 0).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D9D9D9"));
                        }
                        catch (Exception)
                        {
                            //issue detected with getting snappy data, use regular times
                            CurrentCell.Offset(11, 0).Value = CalculateClosestRange(Log.SummaryGraphData);
                            CurrentCell.Offset(11, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            CurrentCell.Offset(11, 0).Style.Border.Right.Style = ExcelBorderStyle.Thick;
                            CurrentCell.Offset(11, 0).Style.Numberformat.Format = "0.00";
                            CurrentCell.Offset(11, 0).Style.Fill.PatternType = ExcelFillStyle.Solid;
                            CurrentCell.Offset(11, 0).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D9D9D9"));

                            CurrentCell.Offset(12, 0).Value = CalculateFarthestRange(Log.SummaryGraphData);
                            CurrentCell.Offset(12, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            CurrentCell.Offset(12, 0).Style.Border.Right.Style = ExcelBorderStyle.Thick;
                            CurrentCell.Offset(12, 0).Style.Numberformat.Format = "0.00";
                            CurrentCell.Offset(12, 0).Style.Fill.PatternType = ExcelFillStyle.Solid;
                            CurrentCell.Offset(12, 0).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D9D9D9"));

                            CurrentCell.Offset(13, 0).Value = CalculateDecodeRange(Log.SummaryGraphData);
                            CurrentCell.Offset(13, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            CurrentCell.Offset(13, 0).Style.Border.Right.Style = ExcelBorderStyle.Thick;
                            CurrentCell.Offset(13, 0).Style.Numberformat.Format = "0.00";
                            CurrentCell.Offset(13, 0).Style.Fill.PatternType = ExcelFillStyle.Solid;
                            CurrentCell.Offset(13, 0).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D9D9D9"));

                            CurrentCell.Offset(14, 0).Value = CalculateAverageDecodeTime(Log.SummaryGraphData);
                            CurrentCell.Offset(14, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            CurrentCell.Offset(14, 0).Style.Border.Right.Style = ExcelBorderStyle.Thick;
                            CurrentCell.Offset(14, 0).Style.Numberformat.Format = "0.00";
                            CurrentCell.Offset(14, 0).Style.Fill.PatternType = ExcelFillStyle.Solid;
                            CurrentCell.Offset(14, 0).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D9D9D9"));

                            //CurrentCell.Offset(15, 0).Value = CalculateStandardDeviationDecodeTime(Log.DecodeTimes);
                            CurrentCell.Offset(15, 0).Value = CalculateStandardDeviationDecodeTimeNEW(Log.SummaryGraphData);
                            CurrentCell.Offset(15, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            CurrentCell.Offset(15, 0).Style.Border.Right.Style = ExcelBorderStyle.Thick;
                            CurrentCell.Offset(15, 0).Style.Numberformat.Format = "0.00";
                            CurrentCell.Offset(15, 0).Style.Fill.PatternType = ExcelFillStyle.Solid;
                            CurrentCell.Offset(15, 0).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D9D9D9"));

                            //CurrentCell.Offset(16, 0).Value = CalculateHighestDecodeTime(Log.DecodeTimes);
                            CurrentCell.Offset(16, 0).Value = CalculateHighestDecodeTimeNEW(Log.SummaryGraphData);
                            CurrentCell.Offset(16, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            CurrentCell.Offset(16, 0).Style.Border.Right.Style = ExcelBorderStyle.Thick;
                            CurrentCell.Offset(16, 0).Style.Numberformat.Format = "0.00";
                            CurrentCell.Offset(16, 0).Style.Fill.PatternType = ExcelFillStyle.Solid;
                            CurrentCell.Offset(16, 0).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D9D9D9"));

                            //CurrentCell.Offset(17, 0).Value = CalculateLowestDecodeTime(Log.DecodeTimes);
                            CurrentCell.Offset(17, 0).Value = CalculateLowestDecodeTimeNEW(Log.SummaryGraphData);
                            CurrentCell.Offset(17, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            CurrentCell.Offset(17, 0).Style.Border.Right.Style = ExcelBorderStyle.Thick;
                            CurrentCell.Offset(17, 0).Style.Numberformat.Format = "0.00";
                            CurrentCell.Offset(17, 0).Style.Fill.PatternType = ExcelFillStyle.Solid;
                            CurrentCell.Offset(17, 0).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D9D9D9"));
                        }

                        break;
                    }
                    else
                    {
                        Column1++;
                    }
                }
            }

            //Prints data to CurrentRow 3 of the scorecard tab
            if (GroupNumber == "3")
            {
                int Row1 = 53;
                int Column1 = 3;

                while (true)
                {
                    var CurrentCell = ScorecardSheet.Cells[Row1, Column1];
                    string CurrentCellValue = CurrentCell.Text;

                    if (string.IsNullOrEmpty(CurrentCellValue))
                    {
                        //Prints sample name
                        CurrentCell.Value = Log.BarcodeSampleName;
                        CurrentCell.Style.Font.Bold = true;
                        CurrentCell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        CurrentCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        CurrentCell.Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#FFF2CC"));

                        //prints scanner name
                        CurrentCell.Offset(1, 0).Value = Log.ScannerMake + "_" + Log.ScannerName;
                        CurrentCell.Offset(1, 0).Style.Font.Bold = true;
                        CurrentCell.Offset(1, 0).Style.Border.Top.Style = ExcelBorderStyle.Thick;
                        CurrentCell.Offset(1, 0).Style.Border.Right.Style = ExcelBorderStyle.Thick;
                        CurrentCell.Offset(1, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        CurrentCell.Offset(1, 0).Style.Fill.PatternType = ExcelFillStyle.Solid;
                        CurrentCell.Offset(1, 0).Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#DDEBF7"));

                        //Notes field for the user to enter in any relevant notes about the scanner
                        CurrentCell.Offset(2, 0).Value = "Notes:";
                        CurrentCell.Offset(2, 0).Style.Font.Bold = true;
                        CurrentCell.Offset(2, 0).Style.Border.Right.Style = ExcelBorderStyle.Thick;
                        CurrentCell.Offset(2, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        CurrentCell.Offset(2, 0).Style.Fill.PatternType = ExcelFillStyle.Solid;
                        CurrentCell.Offset(2, 0).Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#C5D9F1"));
                        ScorecardSheet.Row(16).Height = 75;
                        CurrentCell.Offset(3, 0).Style.Border.Right.Style = ExcelBorderStyle.Thick;
                        CurrentCell.Offset(3, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        CurrentCell.Offset(3, 0).Style.Fill.PatternType = ExcelFillStyle.Solid;
                        CurrentCell.Offset(3, 0).Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#C5D9F1"));

                        //total range results
                        CurrentCell.Offset(4, 0).Value = CalculateClosestRange(Log.SummaryGraphData);
                        CurrentCell.Offset(4, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        CurrentCell.Offset(4, 0).Style.Border.Right.Style = ExcelBorderStyle.Thick;
                        CurrentCell.Offset(4, 0).Style.Numberformat.Format = "0.00";
                        CurrentCell.Offset(4, 0).Style.Fill.PatternType = ExcelFillStyle.Solid;
                        CurrentCell.Offset(4, 0).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#F2F2F2"));

                        CurrentCell.Offset(5, 0).Value = CalculateFarthestRange(Log.SummaryGraphData);
                        CurrentCell.Offset(5, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        CurrentCell.Offset(5, 0).Style.Border.Right.Style = ExcelBorderStyle.Thick;
                        CurrentCell.Offset(5, 0).Style.Numberformat.Format = "0.00";
                        CurrentCell.Offset(5, 0).Style.Fill.PatternType = ExcelFillStyle.Solid;
                        CurrentCell.Offset(5, 0).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#F2F2F2"));

                        CurrentCell.Offset(6, 0).Value = CalculateDecodeRange(Log.SummaryGraphData);
                        CurrentCell.Offset(6, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        CurrentCell.Offset(6, 0).Style.Border.Right.Style = ExcelBorderStyle.Thick;
                        CurrentCell.Offset(6, 0).Style.Numberformat.Format = "0.00";
                        CurrentCell.Offset(6, 0).Style.Fill.PatternType = ExcelFillStyle.Solid;
                        CurrentCell.Offset(6, 0).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#F2F2F2"));

                        CurrentCell.Offset(7, 0).Value = CalculateAverageDecodeTime(Log.SummaryGraphData);
                        CurrentCell.Offset(7, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        CurrentCell.Offset(7, 0).Style.Border.Right.Style = ExcelBorderStyle.Thick;
                        CurrentCell.Offset(7, 0).Style.Numberformat.Format = "0.00";
                        CurrentCell.Offset(7, 0).Style.Fill.PatternType = ExcelFillStyle.Solid;
                        CurrentCell.Offset(7, 0).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#F2F2F2"));

                        //CurrentCell.Offset(8, 0).Value = CalculateStandardDeviationDecodeTime(Log.DecodeTimes);
                        CurrentCell.Offset(8, 0).Value = CalculateStandardDeviationDecodeTimeNEW(Log.SummaryGraphData);
                        CurrentCell.Offset(8, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        CurrentCell.Offset(8, 0).Style.Border.Right.Style = ExcelBorderStyle.Thick;
                        CurrentCell.Offset(8, 0).Style.Numberformat.Format = "0.00";
                        CurrentCell.Offset(8, 0).Style.Fill.PatternType = ExcelFillStyle.Solid;
                        CurrentCell.Offset(8, 0).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#F2F2F2"));

                        //CurrentCell.Offset(9, 0).Value = CalculateHighestDecodeTime(Log.DecodeTimes);
                        CurrentCell.Offset(9, 0).Value = CalculateHighestDecodeTimeNEW(Log.SummaryGraphData);
                        CurrentCell.Offset(9, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        CurrentCell.Offset(9, 0).Style.Border.Right.Style = ExcelBorderStyle.Thick;
                        CurrentCell.Offset(9, 0).Style.Numberformat.Format = "0.00";
                        CurrentCell.Offset(9, 0).Style.Fill.PatternType = ExcelFillStyle.Solid;
                        CurrentCell.Offset(9, 0).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#F2F2F2"));

                        //CurrentCell.Offset(10, 0).Value = CalculateLowestDecodeTime(Log.DecodeTimes);
                        CurrentCell.Offset(10, 0).Value = CalculateLowestDecodeTimeNEW(Log.SummaryGraphData);
                        CurrentCell.Offset(10, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        CurrentCell.Offset(10, 0).Style.Border.Right.Style = ExcelBorderStyle.Thick;
                        CurrentCell.Offset(10, 0).Style.Numberformat.Format = "0.00";
                        CurrentCell.Offset(10, 0).Style.Fill.PatternType = ExcelFillStyle.Solid;
                        CurrentCell.Offset(10, 0).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#F2F2F2"));

                        //snappy range results
                        try
                        {
                            CurrentCell.Offset(11, 0).Value = CalculateClosestRange(Log.SnappyData);
                            CurrentCell.Offset(11, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            CurrentCell.Offset(11, 0).Style.Border.Right.Style = ExcelBorderStyle.Thick;
                            CurrentCell.Offset(11, 0).Style.Numberformat.Format = "0.00";
                            CurrentCell.Offset(11, 0).Style.Fill.PatternType = ExcelFillStyle.Solid;
                            CurrentCell.Offset(11, 0).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D9D9D9"));

                            CurrentCell.Offset(12, 0).Value = CalculateFarthestRange(Log.SnappyData);
                            CurrentCell.Offset(12, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            CurrentCell.Offset(12, 0).Style.Border.Right.Style = ExcelBorderStyle.Thick;
                            CurrentCell.Offset(12, 0).Style.Numberformat.Format = "0.00";
                            CurrentCell.Offset(12, 0).Style.Fill.PatternType = ExcelFillStyle.Solid;
                            CurrentCell.Offset(12, 0).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D9D9D9"));

                            CurrentCell.Offset(13, 0).Value = CalculateDecodeRange(Log.SnappyData);
                            CurrentCell.Offset(13, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            CurrentCell.Offset(13, 0).Style.Border.Right.Style = ExcelBorderStyle.Thick;
                            CurrentCell.Offset(13, 0).Style.Numberformat.Format = "0.00";
                            CurrentCell.Offset(13, 0).Style.Fill.PatternType = ExcelFillStyle.Solid;
                            CurrentCell.Offset(13, 0).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D9D9D9"));

                            CurrentCell.Offset(14, 0).Value = CalculateAverageDecodeTime(Log.SnappyData);
                            CurrentCell.Offset(14, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            CurrentCell.Offset(14, 0).Style.Border.Right.Style = ExcelBorderStyle.Thick;
                            CurrentCell.Offset(14, 0).Style.Numberformat.Format = "0.00";
                            CurrentCell.Offset(14, 0).Style.Fill.PatternType = ExcelFillStyle.Solid;
                            CurrentCell.Offset(14, 0).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D9D9D9"));

                            //CurrentCell.Offset(15, 0).Value = CalculateStandardDeviationDecodeTime(Log.DecodeTimes);
                            CurrentCell.Offset(15, 0).Value = CalculateStandardDeviationDecodeTimeNEW(Log.SnappyData);
                            CurrentCell.Offset(15, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            CurrentCell.Offset(15, 0).Style.Border.Right.Style = ExcelBorderStyle.Thick;
                            CurrentCell.Offset(15, 0).Style.Numberformat.Format = "0.00";
                            CurrentCell.Offset(15, 0).Style.Fill.PatternType = ExcelFillStyle.Solid;
                            CurrentCell.Offset(15, 0).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D9D9D9"));

                            //CurrentCell.Offset(16, 0).Value = CalculateHighestDecodeTime(Log.DecodeTimes);
                            CurrentCell.Offset(16, 0).Value = CalculateHighestDecodeTimeNEW(Log.SnappyData);
                            CurrentCell.Offset(16, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            CurrentCell.Offset(16, 0).Style.Border.Right.Style = ExcelBorderStyle.Thick;
                            CurrentCell.Offset(16, 0).Style.Numberformat.Format = "0.00";
                            CurrentCell.Offset(16, 0).Style.Fill.PatternType = ExcelFillStyle.Solid;
                            CurrentCell.Offset(16, 0).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D9D9D9"));

                            //CurrentCell.Offset(17, 0).Value = CalculateLowestDecodeTime(Log.DecodeTimes);
                            CurrentCell.Offset(17, 0).Value = CalculateLowestDecodeTimeNEW(Log.SnappyData);
                            CurrentCell.Offset(17, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            CurrentCell.Offset(17, 0).Style.Border.Right.Style = ExcelBorderStyle.Thick;
                            CurrentCell.Offset(17, 0).Style.Numberformat.Format = "0.00";
                            CurrentCell.Offset(17, 0).Style.Fill.PatternType = ExcelFillStyle.Solid;
                            CurrentCell.Offset(17, 0).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D9D9D9"));
                        }
                        catch (Exception)
                        {
                            //issue detected with getting snappy data, use regular data
                            CurrentCell.Offset(11, 0).Value = CalculateClosestRange(Log.SummaryGraphData);
                            CurrentCell.Offset(11, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            CurrentCell.Offset(11, 0).Style.Border.Right.Style = ExcelBorderStyle.Thick;
                            CurrentCell.Offset(11, 0).Style.Numberformat.Format = "0.00";
                            CurrentCell.Offset(11, 0).Style.Fill.PatternType = ExcelFillStyle.Solid;
                            CurrentCell.Offset(11, 0).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D9D9D9"));

                            CurrentCell.Offset(12, 0).Value = CalculateFarthestRange(Log.SummaryGraphData);
                            CurrentCell.Offset(12, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            CurrentCell.Offset(12, 0).Style.Border.Right.Style = ExcelBorderStyle.Thick;
                            CurrentCell.Offset(12, 0).Style.Numberformat.Format = "0.00";
                            CurrentCell.Offset(12, 0).Style.Fill.PatternType = ExcelFillStyle.Solid;
                            CurrentCell.Offset(12, 0).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D9D9D9"));

                            CurrentCell.Offset(13, 0).Value = CalculateDecodeRange(Log.SummaryGraphData);
                            CurrentCell.Offset(13, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            CurrentCell.Offset(13, 0).Style.Border.Right.Style = ExcelBorderStyle.Thick;
                            CurrentCell.Offset(13, 0).Style.Numberformat.Format = "0.00";
                            CurrentCell.Offset(13, 0).Style.Fill.PatternType = ExcelFillStyle.Solid;
                            CurrentCell.Offset(13, 0).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D9D9D9"));

                            CurrentCell.Offset(14, 0).Value = CalculateAverageDecodeTime(Log.SummaryGraphData);
                            CurrentCell.Offset(14, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            CurrentCell.Offset(14, 0).Style.Border.Right.Style = ExcelBorderStyle.Thick;
                            CurrentCell.Offset(14, 0).Style.Numberformat.Format = "0.00";
                            CurrentCell.Offset(14, 0).Style.Fill.PatternType = ExcelFillStyle.Solid;
                            CurrentCell.Offset(14, 0).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D9D9D9"));

                            //CurrentCell.Offset(15, 0).Value = CalculateStandardDeviationDecodeTime(Log.DecodeTimes);
                            CurrentCell.Offset(15, 0).Value = CalculateStandardDeviationDecodeTimeNEW(Log.SummaryGraphData);
                            CurrentCell.Offset(15, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            CurrentCell.Offset(15, 0).Style.Border.Right.Style = ExcelBorderStyle.Thick;
                            CurrentCell.Offset(15, 0).Style.Numberformat.Format = "0.00";
                            CurrentCell.Offset(15, 0).Style.Fill.PatternType = ExcelFillStyle.Solid;
                            CurrentCell.Offset(15, 0).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D9D9D9"));

                            //CurrentCell.Offset(16, 0).Value = CalculateHighestDecodeTime(Log.DecodeTimes);
                            CurrentCell.Offset(16, 0).Value = CalculateHighestDecodeTimeNEW(Log.SummaryGraphData);
                            CurrentCell.Offset(16, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            CurrentCell.Offset(16, 0).Style.Border.Right.Style = ExcelBorderStyle.Thick;
                            CurrentCell.Offset(16, 0).Style.Numberformat.Format = "0.00";
                            CurrentCell.Offset(16, 0).Style.Fill.PatternType = ExcelFillStyle.Solid;
                            CurrentCell.Offset(16, 0).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D9D9D9"));

                            //CurrentCell.Offset(17, 0).Value = CalculateLowestDecodeTime(Log.DecodeTimes);
                            CurrentCell.Offset(17, 0).Value = CalculateLowestDecodeTimeNEW(Log.SummaryGraphData);
                            CurrentCell.Offset(17, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            CurrentCell.Offset(17, 0).Style.Border.Right.Style = ExcelBorderStyle.Thick;
                            CurrentCell.Offset(17, 0).Style.Numberformat.Format = "0.00";
                            CurrentCell.Offset(17, 0).Style.Fill.PatternType = ExcelFillStyle.Solid;
                            CurrentCell.Offset(17, 0).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D9D9D9"));
                        }

                        break;
                    }
                    else
                    {
                        Column1++;
                    }
                }
            }

            //Prints data to CurrentRow 4 of the scorecard tab
            if (GroupNumber == "4")
            {
                int Row1 = 73;
                int Column1 = 3;

                while (true)
                {
                    var CurrentCell = ScorecardSheet.Cells[Row1, Column1];
                    string CurrentCellValue = CurrentCell.Text;

                    if (string.IsNullOrEmpty(CurrentCellValue))
                    {
                        //Prints sample name
                        CurrentCell.Value = Log.BarcodeSampleName;
                        CurrentCell.Style.Font.Bold = true;
                        CurrentCell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        CurrentCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        CurrentCell.Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#FFF2CC"));

                        //prints scanner name
                        CurrentCell.Offset(1, 0).Value = Log.ScannerMake + "_" + Log.ScannerName;
                        CurrentCell.Offset(1, 0).Style.Font.Bold = true;
                        CurrentCell.Offset(1, 0).Style.Border.Top.Style = ExcelBorderStyle.Thick;
                        CurrentCell.Offset(1, 0).Style.Border.Right.Style = ExcelBorderStyle.Thick;
                        CurrentCell.Offset(1, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        CurrentCell.Offset(1, 0).Style.Fill.PatternType = ExcelFillStyle.Solid;
                        CurrentCell.Offset(1, 0).Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#DDEBF7"));

                        //Notes field for the user to enter in any relevant notes about the scanner
                        CurrentCell.Offset(2, 0).Value = "Notes:";
                        CurrentCell.Offset(2, 0).Style.Font.Bold = true;
                        CurrentCell.Offset(2, 0).Style.Border.Right.Style = ExcelBorderStyle.Thick;
                        CurrentCell.Offset(2, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        CurrentCell.Offset(2, 0).Style.Fill.PatternType = ExcelFillStyle.Solid;
                        CurrentCell.Offset(2, 0).Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#C5D9F1"));
                        ScorecardSheet.Row(16).Height = 75;
                        CurrentCell.Offset(3, 0).Style.Border.Right.Style = ExcelBorderStyle.Thick;
                        CurrentCell.Offset(3, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        CurrentCell.Offset(3, 0).Style.Fill.PatternType = ExcelFillStyle.Solid;
                        CurrentCell.Offset(3, 0).Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#C5D9F1"));

                        //total range results
                        CurrentCell.Offset(4, 0).Value = CalculateClosestRange(Log.SummaryGraphData);
                        CurrentCell.Offset(4, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        CurrentCell.Offset(4, 0).Style.Border.Right.Style = ExcelBorderStyle.Thick;
                        CurrentCell.Offset(4, 0).Style.Numberformat.Format = "0.00";
                        CurrentCell.Offset(4, 0).Style.Fill.PatternType = ExcelFillStyle.Solid;
                        CurrentCell.Offset(4, 0).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#F2F2F2"));

                        CurrentCell.Offset(5, 0).Value = CalculateFarthestRange(Log.SummaryGraphData);
                        CurrentCell.Offset(5, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        CurrentCell.Offset(5, 0).Style.Border.Right.Style = ExcelBorderStyle.Thick;
                        CurrentCell.Offset(5, 0).Style.Numberformat.Format = "0.00";
                        CurrentCell.Offset(5, 0).Style.Fill.PatternType = ExcelFillStyle.Solid;
                        CurrentCell.Offset(5, 0).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#F2F2F2"));

                        CurrentCell.Offset(6, 0).Value = CalculateDecodeRange(Log.SummaryGraphData);
                        CurrentCell.Offset(6, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        CurrentCell.Offset(6, 0).Style.Border.Right.Style = ExcelBorderStyle.Thick;
                        CurrentCell.Offset(6, 0).Style.Numberformat.Format = "0.00";
                        CurrentCell.Offset(6, 0).Style.Fill.PatternType = ExcelFillStyle.Solid;
                        CurrentCell.Offset(6, 0).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#F2F2F2"));

                        CurrentCell.Offset(7, 0).Value = CalculateAverageDecodeTime(Log.SummaryGraphData);
                        CurrentCell.Offset(7, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        CurrentCell.Offset(7, 0).Style.Border.Right.Style = ExcelBorderStyle.Thick;
                        CurrentCell.Offset(7, 0).Style.Numberformat.Format = "0.00";
                        CurrentCell.Offset(7, 0).Style.Fill.PatternType = ExcelFillStyle.Solid;
                        CurrentCell.Offset(7, 0).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#F2F2F2"));

                        //CurrentCell.Offset(8, 0).Value = CalculateStandardDeviationDecodeTime(Log.DecodeTimes);
                        CurrentCell.Offset(8, 0).Value = CalculateStandardDeviationDecodeTimeNEW(Log.SummaryGraphData);
                        CurrentCell.Offset(8, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        CurrentCell.Offset(8, 0).Style.Border.Right.Style = ExcelBorderStyle.Thick;
                        CurrentCell.Offset(8, 0).Style.Numberformat.Format = "0.00";
                        CurrentCell.Offset(8, 0).Style.Fill.PatternType = ExcelFillStyle.Solid;
                        CurrentCell.Offset(8, 0).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#F2F2F2"));

                        //CurrentCell.Offset(9, 0).Value = CalculateHighestDecodeTime(Log.DecodeTimes);
                        CurrentCell.Offset(9, 0).Value = CalculateHighestDecodeTimeNEW(Log.SummaryGraphData);
                        CurrentCell.Offset(9, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        CurrentCell.Offset(9, 0).Style.Border.Right.Style = ExcelBorderStyle.Thick;
                        CurrentCell.Offset(9, 0).Style.Numberformat.Format = "0.00";
                        CurrentCell.Offset(9, 0).Style.Fill.PatternType = ExcelFillStyle.Solid;
                        CurrentCell.Offset(9, 0).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#F2F2F2"));

                        //CurrentCell.Offset(10, 0).Value = CalculateLowestDecodeTime(Log.DecodeTimes);
                        CurrentCell.Offset(10, 0).Value = CalculateLowestDecodeTimeNEW(Log.SummaryGraphData);
                        CurrentCell.Offset(10, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        CurrentCell.Offset(10, 0).Style.Border.Right.Style = ExcelBorderStyle.Thick;
                        CurrentCell.Offset(10, 0).Style.Numberformat.Format = "0.00";
                        CurrentCell.Offset(10, 0).Style.Fill.PatternType = ExcelFillStyle.Solid;
                        CurrentCell.Offset(10, 0).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#F2F2F2"));

                        //snappy range results
                        try
                        {
                            CurrentCell.Offset(11, 0).Value = CalculateClosestRange(Log.SnappyData);
                            CurrentCell.Offset(11, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            CurrentCell.Offset(11, 0).Style.Border.Right.Style = ExcelBorderStyle.Thick;
                            CurrentCell.Offset(11, 0).Style.Numberformat.Format = "0.00";
                            CurrentCell.Offset(11, 0).Style.Fill.PatternType = ExcelFillStyle.Solid;
                            CurrentCell.Offset(11, 0).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D9D9D9"));

                            CurrentCell.Offset(12, 0).Value = CalculateFarthestRange(Log.SnappyData);
                            CurrentCell.Offset(12, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            CurrentCell.Offset(12, 0).Style.Border.Right.Style = ExcelBorderStyle.Thick;
                            CurrentCell.Offset(12, 0).Style.Numberformat.Format = "0.00";
                            CurrentCell.Offset(12, 0).Style.Fill.PatternType = ExcelFillStyle.Solid;
                            CurrentCell.Offset(12, 0).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D9D9D9"));

                            CurrentCell.Offset(13, 0).Value = CalculateDecodeRange(Log.SnappyData);
                            CurrentCell.Offset(13, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            CurrentCell.Offset(13, 0).Style.Border.Right.Style = ExcelBorderStyle.Thick;
                            CurrentCell.Offset(13, 0).Style.Numberformat.Format = "0.00";
                            CurrentCell.Offset(13, 0).Style.Fill.PatternType = ExcelFillStyle.Solid;
                            CurrentCell.Offset(13, 0).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D9D9D9"));

                            CurrentCell.Offset(14, 0).Value = CalculateAverageDecodeTime(Log.SnappyData);
                            CurrentCell.Offset(14, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            CurrentCell.Offset(14, 0).Style.Border.Right.Style = ExcelBorderStyle.Thick;
                            CurrentCell.Offset(14, 0).Style.Numberformat.Format = "0.00";
                            CurrentCell.Offset(14, 0).Style.Fill.PatternType = ExcelFillStyle.Solid;
                            CurrentCell.Offset(14, 0).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D9D9D9"));

                            //CurrentCell.Offset(15, 0).Value = CalculateStandardDeviationDecodeTime(Log.DecodeTimes);
                            CurrentCell.Offset(15, 0).Value = CalculateStandardDeviationDecodeTimeNEW(Log.SnappyData);
                            CurrentCell.Offset(15, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            CurrentCell.Offset(15, 0).Style.Border.Right.Style = ExcelBorderStyle.Thick;
                            CurrentCell.Offset(15, 0).Style.Numberformat.Format = "0.00";
                            CurrentCell.Offset(15, 0).Style.Fill.PatternType = ExcelFillStyle.Solid;
                            CurrentCell.Offset(15, 0).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D9D9D9"));

                            //CurrentCell.Offset(16, 0).Value = CalculateHighestDecodeTime(Log.DecodeTimes);
                            CurrentCell.Offset(16, 0).Value = CalculateHighestDecodeTimeNEW(Log.SnappyData);
                            CurrentCell.Offset(16, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            CurrentCell.Offset(16, 0).Style.Border.Right.Style = ExcelBorderStyle.Thick;
                            CurrentCell.Offset(16, 0).Style.Numberformat.Format = "0.00";
                            CurrentCell.Offset(16, 0).Style.Fill.PatternType = ExcelFillStyle.Solid;
                            CurrentCell.Offset(16, 0).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D9D9D9"));

                            //CurrentCell.Offset(17, 0).Value = CalculateLowestDecodeTime(Log.DecodeTimes);
                            CurrentCell.Offset(17, 0).Value = CalculateLowestDecodeTimeNEW(Log.SnappyData);
                            CurrentCell.Offset(17, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            CurrentCell.Offset(17, 0).Style.Border.Right.Style = ExcelBorderStyle.Thick;
                            CurrentCell.Offset(17, 0).Style.Numberformat.Format = "0.00";
                            CurrentCell.Offset(17, 0).Style.Fill.PatternType = ExcelFillStyle.Solid;
                            CurrentCell.Offset(17, 0).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D9D9D9"));
                        }
                        catch (Exception)
                        {
                            //issue detected with getting snapppy data, use regular data instead
                            CurrentCell.Offset(11, 0).Value = CalculateClosestRange(Log.SummaryGraphData);
                            CurrentCell.Offset(11, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            CurrentCell.Offset(11, 0).Style.Border.Right.Style = ExcelBorderStyle.Thick;
                            CurrentCell.Offset(11, 0).Style.Numberformat.Format = "0.00";
                            CurrentCell.Offset(11, 0).Style.Fill.PatternType = ExcelFillStyle.Solid;
                            CurrentCell.Offset(11, 0).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D9D9D9"));

                            CurrentCell.Offset(12, 0).Value = CalculateFarthestRange(Log.SummaryGraphData);
                            CurrentCell.Offset(12, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            CurrentCell.Offset(12, 0).Style.Border.Right.Style = ExcelBorderStyle.Thick;
                            CurrentCell.Offset(12, 0).Style.Numberformat.Format = "0.00";
                            CurrentCell.Offset(12, 0).Style.Fill.PatternType = ExcelFillStyle.Solid;
                            CurrentCell.Offset(12, 0).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D9D9D9"));

                            CurrentCell.Offset(13, 0).Value = CalculateDecodeRange(Log.SummaryGraphData);
                            CurrentCell.Offset(13, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            CurrentCell.Offset(13, 0).Style.Border.Right.Style = ExcelBorderStyle.Thick;
                            CurrentCell.Offset(13, 0).Style.Numberformat.Format = "0.00";
                            CurrentCell.Offset(13, 0).Style.Fill.PatternType = ExcelFillStyle.Solid;
                            CurrentCell.Offset(13, 0).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D9D9D9"));

                            CurrentCell.Offset(14, 0).Value = CalculateAverageDecodeTime(Log.SummaryGraphData);
                            CurrentCell.Offset(14, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            CurrentCell.Offset(14, 0).Style.Border.Right.Style = ExcelBorderStyle.Thick;
                            CurrentCell.Offset(14, 0).Style.Numberformat.Format = "0.00";
                            CurrentCell.Offset(14, 0).Style.Fill.PatternType = ExcelFillStyle.Solid;
                            CurrentCell.Offset(14, 0).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D9D9D9"));

                            //CurrentCell.Offset(15, 0).Value = CalculateStandardDeviationDecodeTime(Log.DecodeTimes);
                            CurrentCell.Offset(15, 0).Value = CalculateStandardDeviationDecodeTimeNEW(Log.SummaryGraphData);
                            CurrentCell.Offset(15, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            CurrentCell.Offset(15, 0).Style.Border.Right.Style = ExcelBorderStyle.Thick;
                            CurrentCell.Offset(15, 0).Style.Numberformat.Format = "0.00";
                            CurrentCell.Offset(15, 0).Style.Fill.PatternType = ExcelFillStyle.Solid;
                            CurrentCell.Offset(15, 0).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D9D9D9"));

                            //CurrentCell.Offset(16, 0).Value = CalculateHighestDecodeTime(Log.DecodeTimes);
                            CurrentCell.Offset(16, 0).Value = CalculateHighestDecodeTimeNEW(Log.SummaryGraphData);
                            CurrentCell.Offset(16, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            CurrentCell.Offset(16, 0).Style.Border.Right.Style = ExcelBorderStyle.Thick;
                            CurrentCell.Offset(16, 0).Style.Numberformat.Format = "0.00";
                            CurrentCell.Offset(16, 0).Style.Fill.PatternType = ExcelFillStyle.Solid;
                            CurrentCell.Offset(16, 0).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D9D9D9"));

                            //CurrentCell.Offset(17, 0).Value = CalculateLowestDecodeTime(Log.DecodeTimes);
                            CurrentCell.Offset(17, 0).Value = CalculateLowestDecodeTimeNEW(Log.SummaryGraphData);
                            CurrentCell.Offset(17, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            CurrentCell.Offset(17, 0).Style.Border.Right.Style = ExcelBorderStyle.Thick;
                            CurrentCell.Offset(17, 0).Style.Numberformat.Format = "0.00";
                            CurrentCell.Offset(17, 0).Style.Fill.PatternType = ExcelFillStyle.Solid;
                            CurrentCell.Offset(17, 0).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D9D9D9"));
                        }

                        break;
                    }
                    else
                    {
                        Column1++;
                    }
                }
            }

            //Prints data to CurrentRow 5 of the scorecard tab
            if (GroupNumber == "5")
            {
                int Row1 = 93;
                int Column1 = 3;

                while (true)
                {
                    var CurrentCell = ScorecardSheet.Cells[Row1, Column1];
                    string CurrentCellValue = CurrentCell.Text;

                    if (string.IsNullOrEmpty(CurrentCellValue))
                    {
                        //Prints sample name
                        CurrentCell.Value = Log.BarcodeSampleName;
                        CurrentCell.Style.Font.Bold = true;
                        CurrentCell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        CurrentCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        CurrentCell.Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#FFF2CC"));

                        //prints scanner name
                        CurrentCell.Offset(1, 0).Value = Log.ScannerMake + "_" + Log.ScannerName;
                        CurrentCell.Offset(1, 0).Style.Font.Bold = true;
                        CurrentCell.Offset(1, 0).Style.Border.Top.Style = ExcelBorderStyle.Thick;
                        CurrentCell.Offset(1, 0).Style.Border.Right.Style = ExcelBorderStyle.Thick;
                        CurrentCell.Offset(1, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        CurrentCell.Offset(1, 0).Style.Fill.PatternType = ExcelFillStyle.Solid;
                        CurrentCell.Offset(1, 0).Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#DDEBF7"));

                        //Notes field for the user to enter in any relevant notes about the scanner
                        CurrentCell.Offset(2, 0).Value = "Notes:";
                        CurrentCell.Offset(2, 0).Style.Font.Bold = true;
                        CurrentCell.Offset(2, 0).Style.Border.Right.Style = ExcelBorderStyle.Thick;
                        CurrentCell.Offset(2, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        CurrentCell.Offset(2, 0).Style.Fill.PatternType = ExcelFillStyle.Solid;
                        CurrentCell.Offset(2, 0).Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#C5D9F1"));
                        ScorecardSheet.Row(16).Height = 75;
                        CurrentCell.Offset(3, 0).Style.Border.Right.Style = ExcelBorderStyle.Thick;
                        CurrentCell.Offset(3, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        CurrentCell.Offset(3, 0).Style.Fill.PatternType = ExcelFillStyle.Solid;
                        CurrentCell.Offset(3, 0).Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#C5D9F1"));

                        //total range results
                        CurrentCell.Offset(4, 0).Value = CalculateClosestRange(Log.SummaryGraphData);
                        CurrentCell.Offset(4, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        CurrentCell.Offset(4, 0).Style.Border.Right.Style = ExcelBorderStyle.Thick;
                        CurrentCell.Offset(4, 0).Style.Numberformat.Format = "0.00";
                        CurrentCell.Offset(4, 0).Style.Fill.PatternType = ExcelFillStyle.Solid;
                        CurrentCell.Offset(4, 0).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#F2F2F2"));

                        CurrentCell.Offset(5, 0).Value = CalculateFarthestRange(Log.SummaryGraphData);
                        CurrentCell.Offset(5, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        CurrentCell.Offset(5, 0).Style.Border.Right.Style = ExcelBorderStyle.Thick;
                        CurrentCell.Offset(5, 0).Style.Numberformat.Format = "0.00";
                        CurrentCell.Offset(5, 0).Style.Fill.PatternType = ExcelFillStyle.Solid;
                        CurrentCell.Offset(5, 0).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#F2F2F2"));

                        CurrentCell.Offset(6, 0).Value = CalculateDecodeRange(Log.SummaryGraphData);
                        CurrentCell.Offset(6, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        CurrentCell.Offset(6, 0).Style.Border.Right.Style = ExcelBorderStyle.Thick;
                        CurrentCell.Offset(6, 0).Style.Numberformat.Format = "0.00";
                        CurrentCell.Offset(6, 0).Style.Fill.PatternType = ExcelFillStyle.Solid;
                        CurrentCell.Offset(6, 0).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#F2F2F2"));

                        CurrentCell.Offset(7, 0).Value = CalculateAverageDecodeTime(Log.SummaryGraphData);
                        CurrentCell.Offset(7, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        CurrentCell.Offset(7, 0).Style.Border.Right.Style = ExcelBorderStyle.Thick;
                        CurrentCell.Offset(7, 0).Style.Numberformat.Format = "0.00";
                        CurrentCell.Offset(7, 0).Style.Fill.PatternType = ExcelFillStyle.Solid;
                        CurrentCell.Offset(7, 0).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#F2F2F2"));

                        //CurrentCell.Offset(8, 0).Value = CalculateStandardDeviationDecodeTime(Log.DecodeTimes);
                        CurrentCell.Offset(8, 0).Value = CalculateStandardDeviationDecodeTimeNEW(Log.SummaryGraphData);
                        CurrentCell.Offset(8, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        CurrentCell.Offset(8, 0).Style.Border.Right.Style = ExcelBorderStyle.Thick;
                        CurrentCell.Offset(8, 0).Style.Numberformat.Format = "0.00";
                        CurrentCell.Offset(8, 0).Style.Fill.PatternType = ExcelFillStyle.Solid;
                        CurrentCell.Offset(8, 0).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#F2F2F2"));

                        //CurrentCell.Offset(9, 0).Value = CalculateHighestDecodeTime(Log.DecodeTimes);
                        CurrentCell.Offset(9, 0).Value = CalculateHighestDecodeTimeNEW(Log.SummaryGraphData);
                        CurrentCell.Offset(9, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        CurrentCell.Offset(9, 0).Style.Border.Right.Style = ExcelBorderStyle.Thick;
                        CurrentCell.Offset(9, 0).Style.Numberformat.Format = "0.00";
                        CurrentCell.Offset(9, 0).Style.Fill.PatternType = ExcelFillStyle.Solid;
                        CurrentCell.Offset(9, 0).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#F2F2F2"));

                        //CurrentCell.Offset(10, 0).Value = CalculateLowestDecodeTime(Log.DecodeTimes);
                        CurrentCell.Offset(10, 0).Value = CalculateLowestDecodeTimeNEW(Log.SummaryGraphData);
                        CurrentCell.Offset(10, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        CurrentCell.Offset(10, 0).Style.Border.Right.Style = ExcelBorderStyle.Thick;
                        CurrentCell.Offset(10, 0).Style.Numberformat.Format = "0.00";
                        CurrentCell.Offset(10, 0).Style.Fill.PatternType = ExcelFillStyle.Solid;
                        CurrentCell.Offset(10, 0).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#F2F2F2"));

                        //snappy range results
                        try
                        {
                            CurrentCell.Offset(11, 0).Value = CalculateClosestRange(Log.SnappyData);
                            CurrentCell.Offset(11, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            CurrentCell.Offset(11, 0).Style.Border.Right.Style = ExcelBorderStyle.Thick;
                            CurrentCell.Offset(11, 0).Style.Numberformat.Format = "0.00";
                            CurrentCell.Offset(11, 0).Style.Fill.PatternType = ExcelFillStyle.Solid;
                            CurrentCell.Offset(11, 0).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D9D9D9"));

                            CurrentCell.Offset(12, 0).Value = CalculateFarthestRange(Log.SnappyData);
                            CurrentCell.Offset(12, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            CurrentCell.Offset(12, 0).Style.Border.Right.Style = ExcelBorderStyle.Thick;
                            CurrentCell.Offset(12, 0).Style.Numberformat.Format = "0.00";
                            CurrentCell.Offset(12, 0).Style.Fill.PatternType = ExcelFillStyle.Solid;
                            CurrentCell.Offset(12, 0).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D9D9D9"));

                            CurrentCell.Offset(13, 0).Value = CalculateDecodeRange(Log.SnappyData);
                            CurrentCell.Offset(13, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            CurrentCell.Offset(13, 0).Style.Border.Right.Style = ExcelBorderStyle.Thick;
                            CurrentCell.Offset(13, 0).Style.Numberformat.Format = "0.00";
                            CurrentCell.Offset(13, 0).Style.Fill.PatternType = ExcelFillStyle.Solid;
                            CurrentCell.Offset(13, 0).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D9D9D9"));

                            CurrentCell.Offset(14, 0).Value = CalculateAverageDecodeTime(Log.SnappyData);
                            CurrentCell.Offset(14, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            CurrentCell.Offset(14, 0).Style.Border.Right.Style = ExcelBorderStyle.Thick;
                            CurrentCell.Offset(14, 0).Style.Numberformat.Format = "0.00";
                            CurrentCell.Offset(14, 0).Style.Fill.PatternType = ExcelFillStyle.Solid;
                            CurrentCell.Offset(14, 0).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D9D9D9"));

                            //CurrentCell.Offset(15, 0).Value = CalculateStandardDeviationDecodeTime(Log.DecodeTimes);
                            CurrentCell.Offset(15, 0).Value = CalculateStandardDeviationDecodeTimeNEW(Log.SnappyData);
                            CurrentCell.Offset(15, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            CurrentCell.Offset(15, 0).Style.Border.Right.Style = ExcelBorderStyle.Thick;
                            CurrentCell.Offset(15, 0).Style.Numberformat.Format = "0.00";
                            CurrentCell.Offset(15, 0).Style.Fill.PatternType = ExcelFillStyle.Solid;
                            CurrentCell.Offset(15, 0).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D9D9D9"));

                            //CurrentCell.Offset(16, 0).Value = CalculateHighestDecodeTime(Log.DecodeTimes);
                            CurrentCell.Offset(16, 0).Value = CalculateHighestDecodeTimeNEW(Log.SnappyData);
                            CurrentCell.Offset(16, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            CurrentCell.Offset(16, 0).Style.Border.Right.Style = ExcelBorderStyle.Thick;
                            CurrentCell.Offset(16, 0).Style.Numberformat.Format = "0.00";
                            CurrentCell.Offset(16, 0).Style.Fill.PatternType = ExcelFillStyle.Solid;
                            CurrentCell.Offset(16, 0).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D9D9D9"));

                            //CurrentCell.Offset(17, 0).Value = CalculateLowestDecodeTime(Log.DecodeTimes);
                            CurrentCell.Offset(17, 0).Value = CalculateLowestDecodeTimeNEW(Log.SnappyData);
                            CurrentCell.Offset(17, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            CurrentCell.Offset(17, 0).Style.Border.Right.Style = ExcelBorderStyle.Thick;
                            CurrentCell.Offset(17, 0).Style.Numberformat.Format = "0.00";
                            CurrentCell.Offset(17, 0).Style.Fill.PatternType = ExcelFillStyle.Solid;
                            CurrentCell.Offset(17, 0).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D9D9D9"));
                        }
                        catch (Exception)
                        {
                            //issue detected with getting snappy data, use regular data instead
                            CurrentCell.Offset(11, 0).Value = CalculateClosestRange(Log.SummaryGraphData);
                            CurrentCell.Offset(11, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            CurrentCell.Offset(11, 0).Style.Border.Right.Style = ExcelBorderStyle.Thick;
                            CurrentCell.Offset(11, 0).Style.Numberformat.Format = "0.00";
                            CurrentCell.Offset(11, 0).Style.Fill.PatternType = ExcelFillStyle.Solid;
                            CurrentCell.Offset(11, 0).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D9D9D9"));

                            CurrentCell.Offset(12, 0).Value = CalculateFarthestRange(Log.SummaryGraphData);
                            CurrentCell.Offset(12, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            CurrentCell.Offset(12, 0).Style.Border.Right.Style = ExcelBorderStyle.Thick;
                            CurrentCell.Offset(12, 0).Style.Numberformat.Format = "0.00";
                            CurrentCell.Offset(12, 0).Style.Fill.PatternType = ExcelFillStyle.Solid;
                            CurrentCell.Offset(12, 0).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D9D9D9"));

                            CurrentCell.Offset(13, 0).Value = CalculateDecodeRange(Log.SummaryGraphData);
                            CurrentCell.Offset(13, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            CurrentCell.Offset(13, 0).Style.Border.Right.Style = ExcelBorderStyle.Thick;
                            CurrentCell.Offset(13, 0).Style.Numberformat.Format = "0.00";
                            CurrentCell.Offset(13, 0).Style.Fill.PatternType = ExcelFillStyle.Solid;
                            CurrentCell.Offset(13, 0).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D9D9D9"));

                            CurrentCell.Offset(14, 0).Value = CalculateAverageDecodeTime(Log.SummaryGraphData);
                            CurrentCell.Offset(14, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            CurrentCell.Offset(14, 0).Style.Border.Right.Style = ExcelBorderStyle.Thick;
                            CurrentCell.Offset(14, 0).Style.Numberformat.Format = "0.00";
                            CurrentCell.Offset(14, 0).Style.Fill.PatternType = ExcelFillStyle.Solid;
                            CurrentCell.Offset(14, 0).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D9D9D9"));

                            //CurrentCell.Offset(15, 0).Value = CalculateStandardDeviationDecodeTime(Log.DecodeTimes);
                            CurrentCell.Offset(15, 0).Value = CalculateStandardDeviationDecodeTimeNEW(Log.SummaryGraphData);
                            CurrentCell.Offset(15, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            CurrentCell.Offset(15, 0).Style.Border.Right.Style = ExcelBorderStyle.Thick;
                            CurrentCell.Offset(15, 0).Style.Numberformat.Format = "0.00";
                            CurrentCell.Offset(15, 0).Style.Fill.PatternType = ExcelFillStyle.Solid;
                            CurrentCell.Offset(15, 0).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D9D9D9"));

                            //CurrentCell.Offset(16, 0).Value = CalculateHighestDecodeTime(Log.DecodeTimes);
                            CurrentCell.Offset(16, 0).Value = CalculateHighestDecodeTimeNEW(Log.SummaryGraphData);
                            CurrentCell.Offset(16, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            CurrentCell.Offset(16, 0).Style.Border.Right.Style = ExcelBorderStyle.Thick;
                            CurrentCell.Offset(16, 0).Style.Numberformat.Format = "0.00";
                            CurrentCell.Offset(16, 0).Style.Fill.PatternType = ExcelFillStyle.Solid;
                            CurrentCell.Offset(16, 0).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D9D9D9"));

                            //CurrentCell.Offset(17, 0).Value = CalculateLowestDecodeTime(Log.DecodeTimes);
                            CurrentCell.Offset(17, 0).Value = CalculateLowestDecodeTimeNEW(Log.SummaryGraphData);
                            CurrentCell.Offset(17, 0).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            CurrentCell.Offset(17, 0).Style.Border.Right.Style = ExcelBorderStyle.Thick;
                            CurrentCell.Offset(17, 0).Style.Numberformat.Format = "0.00";
                            CurrentCell.Offset(17, 0).Style.Fill.PatternType = ExcelFillStyle.Solid;
                            CurrentCell.Offset(17, 0).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D9D9D9"));
                        }

                        break;
                    }
                    else
                    {
                        Column1++;
                    }
                }
            }
        }

        /// <summary>
        /// Highlights the best result in a CurrentRow with green and the worst result in red
        /// </summary>
        private void HighlightBestResultsHorizontal(ExcelWorksheet ScorecardSheet)
        {
            //rows where we are looking for the highest number to be the best
            int[] RowsToCheckHigh = { 18, 19, 25, 26, 38, 39, 45, 46, 58, 59, 65, 66, 78, 79, 85, 86, 98, 99, 105, 106 };

            foreach (int Row in RowsToCheckHigh)
            {
                int StartCol = ScorecardSheet.Dimension.Start.Column;
                int EndCol = ScorecardSheet.Dimension.End.Column;

                double HighestValue = double.MinValue;
                double LowestValue = double.MaxValue;

                //find the largest number
                for (int Col = StartCol; Col <= EndCol; Col++)
                {
                    ExcelRange Cell = ScorecardSheet.Cells[Row, Col];
                    if (Cell.Value != null && double.TryParse(Cell.Value.ToString(), out double CellValue))
                    {
                        if (CellValue > HighestValue)
                        {
                            HighestValue = CellValue;
                        }

                        if (CellValue < LowestValue)
                        {
                            LowestValue = CellValue;
                        }
                    }
                }

                //highlight the cell with the largest number
                for (int Col = StartCol; Col <= EndCol; Col++)
                {
                    ExcelRange Cell = ScorecardSheet.Cells[Row, Col];
                    if (Cell.Value != null && double.TryParse(Cell.Value.ToString(), out double CellValue))
                    {
                        if (CellValue == HighestValue)
                        {
                            Cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                            Cell.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Green);
                        }

                        if (CellValue == LowestValue)
                        {
                            Cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                            Cell.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Red);
                        }
                    }
                }
            }

            //rows where we are looking for the lowest number to be the best
            int[] RowsToCheckLow = { 17, 20, 21, 22, 23, 24, 27, 28, 29, 30, 37, 40, 41, 42, 43, 44, 47, 48, 49, 50, 57, 60, 61,
                62, 63, 64, 67, 68, 69, 70, 77, 80, 81, 82, 83, 84, 87, 88, 89, 90, 97, 100, 101, 102, 103, 104, 107, 108, 109, 110 };

            foreach (int Row in RowsToCheckLow)
            {
                int StartCol = ScorecardSheet.Dimension.Start.Column;
                int EndCol = ScorecardSheet.Dimension.End.Column;

                double LowestValue = double.MaxValue;
                double HighestValue = double.MinValue;

                //find the lowest number
                for (int Col = StartCol; Col <= EndCol; Col++)
                {
                    ExcelRange Cell = ScorecardSheet.Cells[Row, Col];
                    if (Cell.Value != null && double.TryParse(Cell.Value.ToString(), out double CellValue))
                    {
                        if (CellValue < LowestValue)
                        {
                            LowestValue = CellValue;
                        }

                        if (CellValue > HighestValue)
                        {
                            HighestValue = CellValue;
                        }
                    }
                }

                //highlight the cell with the lowest number
                for (int Col = StartCol; Col <= EndCol; Col++)
                {
                    ExcelRange Cell = ScorecardSheet.Cells[Row, Col];
                    if (Cell.Value != null && double.TryParse(Cell.Value.ToString(), out double CellValue))
                    {
                        if (CellValue == LowestValue)
                        {
                            Cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                            Cell.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Green);
                        }

                        if (CellValue == HighestValue)
                        {
                            Cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                            Cell.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Red);
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Highlights the best result in a column with green and the worst result in red
        /// </summary>
        /// <param name="ScorecardSheet"></param>
        private void HighlightBestResultsVertical(ExcelWorksheet ScorecardSheet)
        {
            //rows that need to be the lowest
            int[] RowsToCheckLow = { 17, 20, 21, 22, 23, 24, 27, 28, 29, 30 };

            //rows that need to be the highest
            int[] RowsToCheckHigh = { 18, 19, 25, 26 };

            //all columns that should be checked
            int[] ColumnsToCheck = { 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26 };

            //offset to look at values in each group
            int[] RowOffset = { 0, 20, 40, 60, 80 };

            //stores the highest and lowest cells the be highlighted
            double? HighestValue = null;
            double? LowestValue = null;
            ExcelRange HighestValueCell = null;
            ExcelRange LowestValueCell = null;

            //lowest number is best
            foreach (int Col in ColumnsToCheck)
            {
                foreach (int Row in RowsToCheckLow)
                {
                    foreach (int Offset in RowOffset)
                    {
                        var CellToCheck = ScorecardSheet.Cells[Row + Offset, Col];

                        if (CellToCheck.Value != null && double.TryParse(CellToCheck.Value.ToString(), out double CellValue))
                        {
                            if (HighestValue == null || CellValue >= HighestValue)
                            {
                                HighestValue = CellValue;
                                HighestValueCell = CellToCheck;
                            }
                            if (LowestValue == null || CellValue <= LowestValue)
                            {
                                LowestValue = CellValue;
                                LowestValueCell = CellToCheck;
                            }
                        }
                    }

                    if (HighestValueCell != null)
                    {
                        HighestValueCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        HighestValueCell.Style.Fill.BackgroundColor.SetColor(Color.Red);
                    }

                    if (LowestValueCell != null)
                    {
                        LowestValueCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        LowestValueCell.Style.Fill.BackgroundColor.SetColor(Color.Green);
                    }

                    HighestValue = null;
                    LowestValue = null;
                    HighestValueCell = null;
                    LowestValueCell = null;
                }
            }

            //Highest number is best
            foreach (int Col in ColumnsToCheck)
            {
                foreach (int Row in RowsToCheckHigh)
                {
                    foreach (int Offset in RowOffset)
                    {
                        var CellToCheck = ScorecardSheet.Cells[Row + Offset, Col];

                        if (CellToCheck.Value != null && double.TryParse(CellToCheck.Value.ToString(), out double CellValue))
                        {
                            if (HighestValue == null || CellValue >= HighestValue)
                            {
                                HighestValue = CellValue;
                                HighestValueCell = CellToCheck;
                            }
                            if (LowestValue == null || CellValue <= LowestValue)
                            {
                                LowestValue = CellValue;
                                LowestValueCell = CellToCheck;
                            }
                        }
                    }

                    if (HighestValueCell != null)
                    {
                        HighestValueCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        HighestValueCell.Style.Fill.BackgroundColor.SetColor(Color.Green);
                    }

                    if (LowestValueCell != null)
                    {
                        LowestValueCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        LowestValueCell.Style.Fill.BackgroundColor.SetColor(Color.Red);
                    }

                    HighestValue = null;
                    LowestValue = null;
                    HighestValueCell = null;
                    LowestValueCell = null;
                }
            }
        }

        /// <summary>
        /// Highlight ranking for best average decode times only vertically
        /// </summary>
        /// <param name="ScorecardSheet"></param>
        private void HighlightBestAverageDecodeTimeResultsVertical(ExcelWorksheet ScorecardSheet)
        {
            //rows that need to be the lowest
            int[] RowsToCheckLow = { 20, 27 };

            //all columns that should be checked
            int[] ColumnsToCheck = { 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26 };

            //offset to look at values in each group
            int[] RowOffset = { 0, 20, 40, 60, 80 };

            //stores the highest and lowest cells the be highlighted
            double? HighestValue = null;
            double? LowestValue = null;
            ExcelRange HighestValueCell = null;
            ExcelRange LowestValueCell = null;

            //lowest number is best
            foreach (int Col in ColumnsToCheck)
            {
                foreach (int Row in RowsToCheckLow)
                {
                    foreach (int Offset in RowOffset)
                    {
                        var CellToCheck = ScorecardSheet.Cells[Row + Offset, Col];

                        if (CellToCheck.Value != null && double.TryParse(CellToCheck.Value.ToString(), out double CellValue))
                        {
                            if (HighestValue == null || CellValue >= HighestValue)
                            {
                                HighestValue = CellValue;
                                HighestValueCell = CellToCheck;
                            }
                            if (LowestValue == null || CellValue <= LowestValue)
                            {
                                LowestValue = CellValue;
                                LowestValueCell = CellToCheck;
                            }
                        }
                    }

                    if (HighestValueCell != null)
                    {
                        HighestValueCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        HighestValueCell.Style.Fill.BackgroundColor.SetColor(Color.Red);
                    }

                    if (LowestValueCell != null)
                    {
                        LowestValueCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        LowestValueCell.Style.Fill.BackgroundColor.SetColor(Color.Green);
                    }

                    HighestValue = null;
                    LowestValue = null;
                    HighestValueCell = null;
                    LowestValueCell = null;
                }
            }
        }

        /// <summary>
        /// Highlight ranking for best average decode times only horizontally
        /// </summary>
        /// <param name="ScorecardSheet"></param>
        private void HighlightBestAverageDecodeTimeResultsHorizontal(ExcelWorksheet ScorecardSheet)
        {
            //rows where we are looking for the lowest number to be the best
            int[] RowsToCheckLow = { 20, 27, 40, 47, 60, 67, 80, 87, 100, 107 };

            foreach (int Row in RowsToCheckLow)
            {
                int StartCol = ScorecardSheet.Dimension.Start.Column;
                int EndCol = ScorecardSheet.Dimension.End.Column;

                double LowestValue = double.MaxValue;
                double HighestValue = double.MinValue;

                //find the lowest number
                for (int Col = StartCol; Col <= EndCol; Col++)
                {
                    ExcelRange Cell = ScorecardSheet.Cells[Row, Col];
                    if (Cell.Value != null && double.TryParse(Cell.Value.ToString(), out double CellValue))
                    {
                        if (CellValue < LowestValue)
                        {
                            LowestValue = CellValue;
                        }

                        if (CellValue > HighestValue)
                        {
                            HighestValue = CellValue;
                        }
                    }
                }

                //highlight the cell with the lowest number
                for (int Col = StartCol; Col <= EndCol; Col++)
                {
                    ExcelRange Cell = ScorecardSheet.Cells[Row, Col];
                    if (Cell.Value != null && double.TryParse(Cell.Value.ToString(), out double CellValue))
                    {
                        if (CellValue == LowestValue)
                        {
                            Cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                            Cell.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Green);
                        }

                        if (CellValue == HighestValue)
                        {
                            Cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                            Cell.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Red);
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Holds expected range results for each log
        /// </summary>
        private void ExpectedRange()
        {
            //todo
        }
        #endregion

        #region Helper_Functions
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
            //Adds an additional entry in each CurrentRow of the array to store the average for each distance
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

            //Adds the Average to the end of each CurrentRow
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

            //removes averages from rows that will not be included in the range
            int ZeroCount = 0;

            for (int i = 0; i < Rows2; i++)
            {
                for (int j = 0; j < Cols2; j++)
                {
                    if (BiggerFloatArray[i, j] == 0.0f)
                    {
                        ZeroCount++;
                    }
                }

                //if there are 3 or more missed decodes, set the average to zero so it will not be used in later functions
                if (ZeroCount >= 3)
                {
                    BiggerFloatArray[i, Cols2] = 0.0f;
                }

                ZeroCount = 0;
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
            int LastCol = Cols - 1;
            bool SnappyRangeFoundFlag = false;

            List<float[]> ValidContent = new List<float[]>();

            try
            {
                for (int i = 0; i < Rows; i++)
                {
                    float[] CurrentRow = new float[Cols];

                    //check if Current CurrentRow is in range of being considered snappy
                    if (FloatArray[i, LastCol] < 100 && FloatArray[i, LastCol] != 0)
                    {
                        for (int j = 0; j < Cols; j++)
                        {
                            CurrentRow[j] = FloatArray[i, j];
                        }

                        SnappyRangeFoundFlag = true;
                        ValidContent.Add(CurrentRow);
                        continue;
                    }
                    else if (FloatArray[i, LastCol] > 100 && SnappyRangeFoundFlag == false)
                    {
                        continue;
                    }
                    else if (FloatArray[i, LastCol] > 100 && SnappyRangeFoundFlag == true)
                    {
                        for (int j = 0; j < Cols; j++)
                        {
                            CurrentRow[j] = FloatArray[i + 1, j];
                        }

                        ValidContent.Add(CurrentRow);
                        SnappyRangeFoundFlag = false;
                        break;
                    }
                }
            }
            catch (Exception)
            {
                //throw;
            }

            float[,] SnappyLogContent = new float[ValidContent.Count, Cols];

            for (int i = 0; i < ValidContent.Count; i++)
            {
                for (int j = 0; j < Cols; j++)
                {
                    SnappyLogContent[i, j] = ValidContent[i][j];
                }
            }

            //Handles if snappy range is actually worth printing
            //if the detected snappy range is two or less distances, then its not actually a snappy range
            if (SnappyLogContent.GetLength(0) <= 3)
            {              
                return FloatArray;
            }

            return SnappyLogContent;
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
        /// Gets rid of rows that will not be a part of the range because 7 or more deocdes were not detected
        /// </summary>
        /// <param name="FloatArray"></param>
        /// <returns></returns>
        private float[,] GetRidOfUnUsableRows(float[,] FloatArray)
        {
            int Rows = FloatArray.GetLength(0);
            int Cols = FloatArray.GetLength(1);
            int LastCol = Cols - 1;

            List<float[]> ValidRows = new List<float[]>();

            //add only the uable rows
            for (int i = 0; i < Rows; i++)
            {
                if (FloatArray[i, LastCol] != 0.0f)
                {
                    // Add the CurrentRow to the list if the last entry is not zero
                    float[] CurrentRow = new float[Cols];
                    for (int j = 0; j < Cols; j++)
                    {
                        CurrentRow[j] = FloatArray[i, j];
                    }
                    ValidRows.Add(CurrentRow);
                }
            }

            // Convert list to 2D array
            float[,] UsableDataOnlyFloatArray = new float[ValidRows.Count, Cols];
            for (int i = 0; i < ValidRows.Count; i++)
            {
                for (int j = 0; j < Cols; j++)
                {
                    UsableDataOnlyFloatArray[i, j] = ValidRows[i][j];
                }
            }

            return UsableDataOnlyFloatArray;
        }

        /// <summary>
        /// Calculates the average decode time from the first 90% decode times in the log
        /// </summary>
        /// <param name="FloatArray"></param>
        /// <returns></returns>
        private float CalculateAverageDecodeTime(float[,] FloatArray)
        {
            float TotalSumDecodeTime = 0;
            float TotalAverageDecodeTime;

            //gets rid of unusable rows that are not apart of the range
            FloatArray = GetRidOfUnUsableRows(FloatArray);

            //added
            double TestRange = FloatArray.GetLength(0);

            //gets 90% of the decode range to be the length
            //double DecodeRange = (double)CalculateDecodeRange(FloatArray);
            //int NintyPercentRange = Convert.ToInt32(Math.Truncate(DecodeRange * 0.90));
            int NintyPercentRange = Convert.ToInt32(Math.Truncate(TestRange * 0.90));

            float[] AllDistanceAverages = new float[FloatArray.GetLength(0)];

            //populates all averages into array to be averaged for the decode time average
            for (int i = 0; i < FloatArray.GetLength(0); i++)
            {
                AllDistanceAverages[i] = FloatArray[i, FloatArray.GetLength(1) - 1];
            }

            //Gets rid of all zeroes
            AllDistanceAverages = AllDistanceAverages.Where(a => a != 0).ToArray();

            //added to get average of only the first 90% of the averages rounded up to the nearest whole number
            //Array.Resize(ref AllDistanceAverages, Convert.ToInt32(Math.Truncate(AllDistanceAverages.Length * 0.90)));
            Array.Resize(ref AllDistanceAverages, NintyPercentRange);

            //calculates the average
            foreach (float element in AllDistanceAverages)
            {
                TotalSumDecodeTime += element;
            }

            TotalAverageDecodeTime = TotalSumDecodeTime / AllDistanceAverages.Length;

            return TotalAverageDecodeTime;
        }

        /// <summary>
        /// Calculates the standard deviation decode time from all decodes in a log
        /// This is an old function that is unsued. Keeping for reference!!!!!!!!!!!!!!!!!!!!!!!!!
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
        /// Calculates the standard deviation decode time from the first 90% of decodes
        /// </summary>
        /// <param name="FloatArray"></param>
        /// <returns></returns>
        private float CalculateStandardDeviationDecodeTimeNEW(float[,] FloatArray)
        {
            //variables
            float StandardDeviationDecodeTime = 0.0f;
            float Sum = 0.0f;
            float Mean = 0.0f;
            float VarianceSum = 0.0f;
            float Variance = 0.0f;

            //gets rid of unusable rows that are not apart of the range
            FloatArray = GetRidOfUnUsableRows(FloatArray);

            //gets rid of the first and last entry in array to get rid of distances at the front and averages at the end of each row
            float [,] DecodeTimesOnly = GetDecodeTimesOnly(FloatArray);

            int Rows = DecodeTimesOnly.GetLength(0);
            int Cols = DecodeTimesOnly.GetLength(1);

            //gets 90% of rows
            int NintyPercentOfRows = Convert.ToInt32(Math.Truncate(Rows * 0.90));

            //new array with only the decode times from the first 90% of the decode range
            float[,] ResultArray = new float[NintyPercentOfRows, Cols];

            //moves times from old array to new array
            for (int i = 0; i < NintyPercentOfRows; i++)
            {
                for (int j = 0; j < Cols; j++)
                {
                    ResultArray[i, j] = DecodeTimesOnly[i, j];
                }
            }

            //List of all decode times
            List<float> DecodeTimesList = new List<float>();

            for (int i = 0; i < ResultArray.GetLength(0); i++)
            {
                for (int j = 0; j < ResultArray.GetLength(1); j++)
                {
                    DecodeTimesList.Add(DecodeTimesOnly[i, j]);
                }
            }

            //Gets sum of all decodes
            foreach (float DecodeTime in DecodeTimesList)
            {
                Sum += DecodeTime;
            }

            //Gets mean of all decode times
            Mean = Sum / DecodeTimesList.Count;

            //Gets variance of all decode times
            foreach (float DecodeTime in DecodeTimesList)
            {
                VarianceSum += (DecodeTime - Mean) * (DecodeTime - Mean);
            }
            Variance = VarianceSum / DecodeTimesList.Count;

            //gets square root of variance, which is the standard deviation
            StandardDeviationDecodeTime = (float)Math.Sqrt(Variance);

            return StandardDeviationDecodeTime;
        }

        /// <summary>
        /// Determines the highest decode time from all decodes in a log
        /// Old. Keeping for reference perposes!!!!!!!!!
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
        /// Determines the highest decode time from all decodes in the first 90% of the decode range
        /// </summary>
        /// <param name="FloatArray"></param>
        /// <returns></returns>
        private float CalculateHighestDecodeTimeNEW(float[,] FloatArray)
        {
            float HighestDecodeTime = 0.0f;

            //gets rid of unusable rows that are not apart of the range
            FloatArray = GetRidOfUnUsableRows(FloatArray);

            //gets rid of the first and last entry in array to get rid of distances at the front and averages at the end of each row
            float[,] DecodeTimesOnly = GetDecodeTimesOnly(FloatArray);

            int Rows = DecodeTimesOnly.GetLength(0);
            int Cols = DecodeTimesOnly.GetLength(1);

            //gets 90% of rows
            int NintyPercentOfRows = Convert.ToInt32(Math.Truncate(Rows * 0.90));

            //new array with only the decode times from the first 90% of the decode range
            float[,] ResultArray = new float[NintyPercentOfRows, Cols];

            //moves times from old array to new array
            for (int i = 0; i < NintyPercentOfRows; i++)
            {
                for (int j = 0; j < Cols; j++)
                {
                    ResultArray[i, j] = DecodeTimesOnly[i, j];
                }
            }

            //List of all decode times
            List<float> DecodeTimesList = new List<float>();

            for (int i = 0; i < ResultArray.GetLength(0); i++)
            {
                for (int j = 0; j < ResultArray.GetLength(1); j++)
                {
                    DecodeTimesList.Add(DecodeTimesOnly[i, j]);
                }
            }

            //removes zeros from list
            DecodeTimesList.RemoveAll(i => i == 0);

            //gets largest number
            HighestDecodeTime = DecodeTimesList.Max();

            return HighestDecodeTime;
        }

        /// <summary>
        /// Determines the lowest decode time from all decodes in a log
        /// Old. Keeping for reference perposes!!!!!!!!!
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
        /// Determines the highest decode time from all decodes in the first 90% of the decode range
        /// </summary>
        /// <param name="FloatArray"></param>
        /// <returns></returns>
        private float CalculateLowestDecodeTimeNEW(float[,] FloatArray)
        {
            float LowestDecodeTime = 0.0f;

            //gets rid of unusable rows that are not apart of the range
            FloatArray = GetRidOfUnUsableRows(FloatArray);

            //gets rid of the first and last entry in array to get rid of distances at the front and averages at the end of each row
            float[,] DecodeTimesOnly = GetDecodeTimesOnly(FloatArray);

            int Rows = DecodeTimesOnly.GetLength(0);
            int Cols = DecodeTimesOnly.GetLength(1);

            //gets 90% of rows
            int NintyPercentOfRows = Convert.ToInt32(Math.Truncate(Rows * 0.90));

            //new array with only the decode times from the first 90% of the decode range
            float[,] ResultArray = new float[NintyPercentOfRows, Cols];

            //moves times from old array to new array
            for (int i = 0; i < NintyPercentOfRows; i++)
            {
                for (int j = 0; j < Cols; j++)
                {
                    ResultArray[i, j] = DecodeTimesOnly[i, j];
                }
            }

            //List of all decode times
            List<float> DecodeTimesList = new List<float>();

            for (int i = 0; i < ResultArray.GetLength(0); i++)
            {
                for (int j = 0; j < ResultArray.GetLength(1); j++)
                {
                    DecodeTimesList.Add(DecodeTimesOnly[i, j]);
                }
            }

            //removes zeros from list
            DecodeTimesList.RemoveAll(i => i == 0);

            //gets largest number
            LowestDecodeTime = DecodeTimesList.Min();

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
        /// Gets Distance Averages in order from the array for the scorecard tab
        /// </summary>
        /// <param name="FloatArray"></param>
        /// <returns></returns>
        private float[] GetDistanceAveragesOnly(float[,] FloatArray)
        {
            FloatArray = GetRidOfUnUsableRows(FloatArray);

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
        /// Gets Distance Averages in order from the array for the Excel line graph
        /// </summary>
        /// <param name="FloatArray"></param>
        /// <returns></returns>
        private float[] GetDistanceAveragesOnlyForGraph(float[,] FloatArray)
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
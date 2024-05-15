using System;
using System.IO;
using System.Diagnostics;
using System.Windows.Forms;
using System.Linq;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using OfficeOpenXml.Drawing.Controls;
using System.Linq.Expressions;
using OfficeOpenXml.Drawing.Chart;

namespace Hakbal
{
    public partial class Form1 : Form
    {
        //Create an instance of the TargetLog class
        TargetLog TLog = new TargetLog();

        //Log file to gather decode times
        string SelectedLog = "blank";

        public Form1()
        {
            InitializeComponent();

            //Window will no longer be resizable
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
        }

        /// <summary>
        /// Sets the scanners name for the compiled log
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void SetScannerNameButton_Click(object sender, EventArgs e)
        {
            try
            {
                TLog.ScannerName = ScannerNameTextBox.Text;

                if (ScannerNameTextBox.Text == "")
                {
                    ErrorMessageTextBox.Text = "Please Enter Your Scanner's Name.";
                }
            }
            catch (Exception)
            {
                ErrorMessageTextBox.Text = "Error Setting Scanner's Name.";
            }
        }

        /// <summary>
        /// Sets the scanners make for the compiled log
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void SetScannerMakeButton_Click(object sender, EventArgs e)
        {
            try
            {
                TLog.ScannerMake = ScannerMakeTextBox.Text;

                if (ScannerMakeTextBox.Text == "")
                {
                    ErrorMessageTextBox.Text = "Please Enter Your Scanner's Make.";
                }
            }
            catch (Exception)
            {
                ErrorMessageTextBox.Text = "Error Setting Scanner's Make.";
            }
        }

        /// <summary>
        /// Sets the scanners serial number for the compiled log
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void SetScannerSerialNumberButton_Click(object sender, EventArgs e)
        {
            try
            {
                TLog.ScannerSerialNumber = ScannerSerialNumberTextBox.Text;

                if (ScannerSerialNumberTextBox.Text == "")
                {
                    ErrorMessageTextBox.Text = "Please Enter Your Scanner's Serial Number.";
                }
            }
            catch (Exception)
            {
                ErrorMessageTextBox.Text = "Error Setting Scanner's Serial Number.";
            }
        }

        /// <summary>
        /// Sets the barcode samples name for the compiled log
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void SetBarcodeSampleNameButton_Click(object sender, EventArgs e)
        {
            try
            {
                TLog.BarcodeSampleName = BarcodeSampleNameTextBox.Text;

                if (BarcodeSampleNameTextBox.Text == "")
                {
                    ErrorMessageTextBox.Text = "Please Enter Your Barcode Sample's Name.";
                }
            }
            catch (Exception)
            {
                ErrorMessageTextBox.Text = "Error Setting Barcode Sample's Name.";
            }
        }

        /// <summary>
        /// Opens up the file explorer for the user to select their log
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BrowseFilesButton_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            ofd.Filter = "log files (*.log)|*.log|All files (*.*)|*.*";
            ofd.Multiselect = false;

            try
            {
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    SelectedLog = ofd.FileName;
                    LogFilePathTextBox.Text = ofd.FileName;
                }
                else
                {
                    SelectedLog = "N/A";

                    ErrorMessageTextBox.Text = "The File You Tried to Find Does Not Exists In The Selected Directory.";
                }
            }
            catch (Exception)
            {
                ErrorMessageTextBox.Text = "Error Finding Log File.";
            }
        }

        /// <summary>
        /// Sets the log path to compile
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void SetLogPathButton_Click(object sender, EventArgs e)
        {
            try
            {
                TLog.LogFilePath = LogFilePathTextBox.Text;
            }
            catch (Exception)
            {
                ErrorMessageTextBox.Text = "Error Setting Log File.";
            }
        }

        /// <summary>
        /// Clears the error message text box when clicked
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ClearErrorMessageButton_Click(object sender, EventArgs e)
        {
            try
            {
                ErrorMessageTextBox.Text = String.Empty;
            }
            catch
            {

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
                ScannerNameTextBox.Text = String.Empty;
                ScannerMakeTextBox.Text = String.Empty;
                ScannerSerialNumberTextBox.Text = String.Empty;
                BarcodeSampleNameTextBox.Text = String.Empty;
                LogFilePathTextBox.Text = String.Empty;
            }
            catch
            {
                ErrorMessageTextBox.Text = "There was an issue clearing all text boxes.";
            }
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

            //Excell Spreadsheet Creation
            using (var excelPackage = new ExcelPackage())
            {
                var RawResultsSheet = excelPackage.Workbook.Worksheets.Add(TLog.ScannerName + "_" + TLog.BarcodeSampleName);

                var ScorecardsSheet = excelPackage.Workbook.Worksheets.Add("Scorecard");

                //Prints Device + Barcode information from user for Raw SummaryGraphData Tab
                try
                {
                    RawResultsSheet.Cells["A1"].Value = "Scanner Name: " + TLog.ScannerName;
                    RawResultsSheet.Cells["A1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    RawResultsSheet.Cells["A1"].Style.Fill.BackgroundColor.SetColor(Color.Orange);

                    RawResultsSheet.Cells["A2"].Value = "Scanner Make: " + TLog.ScannerMake;
                    RawResultsSheet.Cells["A2"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    RawResultsSheet.Cells["A2"].Style.Fill.BackgroundColor.SetColor(Color.Orange);

                    RawResultsSheet.Cells["A3"].Value = "Scanner Serial Number: " + TLog.ScannerSerialNumber;
                    RawResultsSheet.Cells["A3"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    RawResultsSheet.Cells["A3"].Style.Fill.BackgroundColor.SetColor(Color.Orange);

                    RawResultsSheet.Cells["A4"].Value = "Barcode Sample Name: " + TLog.BarcodeSampleName;
                    RawResultsSheet.Cells["A4"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    RawResultsSheet.Cells["A4"].Style.Fill.BackgroundColor.SetColor(Color.Orange);

                    RawResultsSheet.Cells["A1:A4"].AutoFitColumns();
                }
                catch (Exception)
                {
                    ErrorMessageTextBox.Text = "Error Printing Device + Barcode Summary In The Spreadsheet. Did You Set All Device Information Correctly?";
                }

                //Prints table headers for Raw SummaryGraphData Tab
                try
                {
                    var LogFileNameGraphHeader = RawResultsSheet.Cells["B5:M5"];
                    LogFileNameGraphHeader.Merge = true;
                    LogFileNameGraphHeader.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    LogFileNameGraphHeader.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    //RawResultsSheet.Cells["B5"].Value = LogFilePathTextBox.Text.ToString();
                    RawResultsSheet.Cells["B5"].Value = GetFileName(LogFilePathTextBox.Text.ToString());

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
                }
                catch (Exception)
                {
                    ErrorMessageTextBox.Text = "Error Printing Table Headers In The Spreadsheet.";
                }

                //Print distances + decode times from log to the Raw SummaryGraphData Tab
                try
                {
                    //Calls functions to get what we need from the log
                    TLog.SummaryGraphData = GetDecodeTimesAndAverages(TLog.LogFilePath);
                    TLog.DecodeTimes = Flatten2DArray(GetDecodeTimesOnly(TLog.SummaryGraphData));
                    TLog.SnappyData = GetSnappyTimes(GetDecodeTimesAndAverages(TLog.LogFilePath));
                    TLog.SnappyDecodeTimes = Flatten2DArray(GetDecodeTimesOnly(TLog.SnappyData));//test

                    int Rows = TLog.SummaryGraphData.GetLength(0);
                    int Cols = TLog.SummaryGraphData.GetLength(1);

                    for (int i = 0; i < Rows; i++)
                    {
                        for (int j = 0; j < Cols; j++)
                        {
                            //RawResultsSheet.Cells[i + 7, j + 2].Value = String.Format("{0:N2}", TLog.SummaryGraphData[i, j]);
                            RawResultsSheet.Cells[i + 7, j + 2].Value = TLog.SummaryGraphData[i, j];
                            RawResultsSheet.Cells[i + 7, j + 2].Style.Numberformat.Format = "0";
                        }
                    }

                    //Formatting for the graph data
                    RawResultsSheet.Cells["B:B"].Style.Numberformat.Format = "0.00";
                    RawResultsSheet.Cells["M:M"].Style.Numberformat.Format = "0.00";
                }
                catch (Exception)
                {
                    ErrorMessageTextBox.Text = "Error Printing Out Summary Of Decode Times In The Spreadsheet. Did You Select A Proper Log File?";
                }

                //Create a line graph in the Raw SummaryGraphData Tab
                try
                {
                    var LineGraph = RawResultsSheet.Drawings.AddChart("Results", OfficeOpenXml.Drawing.Chart.eChartType.LineMarkersStacked);

                    //From Row, From Col - To Row, To Col
                    var YRange = RawResultsSheet.Cells[7, 2, GetDistancesOnly(TLog.SummaryGraphData).Length + 6, 2];
                    var XRange = RawResultsSheet.Cells[7, 13, GetDistanceAveragesOnly(TLog.SummaryGraphData).Length + 6, 13];
                    var Series = LineGraph.Series.Add(XRange, YRange);

                    LineGraph.Title.Text = TLog.ScannerName + "_" + TLog.BarcodeSampleName + "_Average_Results";
                    LineGraph.XAxis.Title.Text = "Distance (In)";
                    LineGraph.YAxis.Title.Text = "Decode Time (MS)";
                    LineGraph.XAxis.MajorUnit = .1;
                    LineGraph.SetPosition(4, 10, 13, 10);
                    LineGraph.SetSize(1000, 600);

                }
                catch (Exception)
                {
                    ErrorMessageTextBox.Text = "Error Creating The Line Graph In The Spreadsheet. Did You Select A Proper Log File?";
                }

                //Print calculations from the log file to the Scorecard Tab
                try
                {
                    //scorecard header styling
                    ScorecardsSheet.Cells["B1"].Value = "Decode Range (In):";
                    ScorecardsSheet.Cells["B1"].Style.Font.Bold = true;
                    ScorecardsSheet.Cells["C1"].Value = "The largest range of distances that the scanner decoded at with no gap in decode. If the scanner decoded outside of this range, it is NOT included.";
                    ScorecardsSheet.Cells["C1"].Style.Font.Bold = true;

                    ScorecardsSheet.Cells["B2"].Value = "Minimum Distance (In):";
                    ScorecardsSheet.Cells["B2"].Style.Font.Bold = true;
                    ScorecardsSheet.Cells["C2"].Value = "The minimum distance of the Decode Range.";
                    ScorecardsSheet.Cells["C2"].Style.Font.Bold = true;

                    ScorecardsSheet.Cells["B3"].Value = "Maximum Distance (In):";
                    ScorecardsSheet.Cells["B3"].Style.Font.Bold = true;
                    ScorecardsSheet.Cells["C3"].Value = "The maximum distance of the Decode Range.";
                    ScorecardsSheet.Cells["C3"].Style.Font.Bold = true;

                    ScorecardsSheet.Cells["B4"].Value = "Total Average (MS):";
                    ScorecardsSheet.Cells["B4"].Style.Font.Bold = true;
                    ScorecardsSheet.Cells["C4"].Value = "The average decode time of the decode range.";
                    ScorecardsSheet.Cells["C4"].Style.Font.Bold = true;

                    ScorecardsSheet.Cells["B5"].Value = "Total Average (90%) (MS):";
                    ScorecardsSheet.Cells["B5"].Style.Font.Bold = true;
                    ScorecardsSheet.Cells["C5"].Value = "The average decode time of the first 90% of the decode range.";
                    ScorecardsSheet.Cells["C5"].Style.Font.Bold = true;

                    ScorecardsSheet.Cells["B6"].Value = "Standard Deviation (MS):";
                    ScorecardsSheet.Cells["B6"].Style.Font.Bold = true;
                    ScorecardsSheet.Cells["C6"].Value = "The standard deviation of the decode range.";
                    ScorecardsSheet.Cells["C6"].Style.Font.Bold = true;

                    ScorecardsSheet.Cells["B7"].Value = "Highest Decode Time (MS)";
                    ScorecardsSheet.Cells["B7"].Style.Font.Bold = true;
                    ScorecardsSheet.Cells["C7"].Value = "The highest decode time from the Decode Range.";
                    ScorecardsSheet.Cells["C7"].Style.Font.Bold = true;

                    ScorecardsSheet.Cells["B8"].Value = "Lowest Decode Time (MS):";
                    ScorecardsSheet.Cells["B8"].Style.Font.Bold = true;
                    ScorecardsSheet.Cells["C8"].Value = "The lowest decode time from the Deocde Range.";
                    ScorecardsSheet.Cells["C8"].Style.Font.Bold = true;

                    ScorecardsSheet.Cells["B10"].Value = "Total Range:";
                    ScorecardsSheet.Cells["B10"].Style.Font.Bold = true;
                    ScorecardsSheet.Cells["C10"].Value = "Statistics for the entire range of decodes consisting of any scans as long as there was a decode of any length.";
                    ScorecardsSheet.Cells["C10"].Style.Font.Bold = true;

                    ScorecardsSheet.Cells["B11"].Value = "Snappy Range:";
                    ScorecardsSheet.Cells["B11"].Style.Font.Bold = true;
                    ScorecardsSheet.Cells["C11"].Value = "Statistics for only the \"snappiest\" range where decode times are 500ms and under.";
                    ScorecardsSheet.Cells["C11"].Style.Font.Bold = true;

                    //prints sample name
                    ScorecardsSheet.Cells["C13"].Value = TLog.BarcodeSampleName;
                    ScorecardsSheet.Cells["C13"].Style.Font.Bold = true;
                    ScorecardsSheet.Cells["C13"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    ScorecardsSheet.Cells["C13"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ScorecardsSheet.Cells["C13"].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#FFF2CC"));

                    //prints scanner name
                    ScorecardsSheet.Cells["C14"].Value = TLog.ScannerMake + "_" + TLog.ScannerName;
                    ScorecardsSheet.Cells["C14"].Style.Font.Bold = true;
                    ScorecardsSheet.Cells["C14"].Style.Border.Top.Style = ExcelBorderStyle.Thick;
                    ScorecardsSheet.Cells["C14"].Style.Border.Right.Style = ExcelBorderStyle.Thick;
                    ScorecardsSheet.Cells["C14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    ScorecardsSheet.Cells["C14"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ScorecardsSheet.Cells["C14"].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#DDEBF7"));

                    //worksheet styling
                    var TotalRangeTag = ScorecardsSheet.Cells["A15:A21"];
                    TotalRangeTag.Merge = true;
                    TotalRangeTag.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    TotalRangeTag.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    ScorecardsSheet.Cells["A15"].Value = "Total Range";
                    ScorecardsSheet.Cells["A15"].Style.TextRotation = 45;

                    ScorecardsSheet.Column(1).Width = 14;
                    ScorecardsSheet.Column(2).Width = 30;
                    ScorecardsSheet.Column(3).Width = 35;

                    ScorecardsSheet.Cells["C15:C28"].Style.Numberformat.Format = "0.00";

                    var SnappyRangeTag = ScorecardsSheet.Cells["A22:A28"];
                    SnappyRangeTag.Merge = true;
                    SnappyRangeTag.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    SnappyRangeTag.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    ScorecardsSheet.Cells["A22"].Value = "Snappy Range";
                    ScorecardsSheet.Cells["A22"].Style.TextRotation = 45;

                    //total range results
                    ScorecardsSheet.Cells["B15"].Value = "Minimum Distance (In):";
                    ScorecardsSheet.Cells["B15"].Style.Font.Bold = true;
                    ScorecardsSheet.Cells["C15"].Value = CalculateClosestRange(TLog.SummaryGraphData);
                    ScorecardsSheet.Cells["C15"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    ScorecardsSheet.Cells["C15"].Style.Border.Right.Style = ExcelBorderStyle.Thick;

                    ScorecardsSheet.Cells["B16"].Value = "Maximum Distance (In):";
                    ScorecardsSheet.Cells["B16"].Style.Font.Bold = true;
                    ScorecardsSheet.Cells["C16"].Value = CalculateFarthestRange(TLog.SummaryGraphData);
                    ScorecardsSheet.Cells["C16"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    ScorecardsSheet.Cells["C16"].Style.Border.Right.Style = ExcelBorderStyle.Thick;

                    ScorecardsSheet.Cells["B17"].Value = "Decode Range (In):";
                    ScorecardsSheet.Cells["B17"].Style.Font.Bold = true;
                    ScorecardsSheet.Cells["C17"].Value = CalculateDecodeRange(TLog.SummaryGraphData);
                    ScorecardsSheet.Cells["C17"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    ScorecardsSheet.Cells["C17"].Style.Border.Right.Style = ExcelBorderStyle.Thick;

                    ScorecardsSheet.Cells["B18"].Value = "Total Average (MS):";
                    ScorecardsSheet.Cells["B18"].Style.Font.Bold = true;
                    ScorecardsSheet.Cells["C18"].Value = CalculateAverageDecodeTime(TLog.SummaryGraphData);
                    ScorecardsSheet.Cells["C18"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    ScorecardsSheet.Cells["C18"].Style.Border.Right.Style = ExcelBorderStyle.Thick;

                    //commented out because average 90% was no longer needed
                    /*
                    ScorecardsSheet.Cells["B"].Value = "Total Average (90%) (MS):";
                    ScorecardsSheet.Cells["B"].Style.Font.Bold = true;
                    ScorecardsSheet.Cells["C"].Value = CalculateNinetyPercentAverageDecodeTime(TLog.SummaryGraphData);
                    ScorecardsSheet.Cells["C"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    ScorecardsSheet.Cells["C"].Style.Border.Right.Style = ExcelBorderStyle.Thick;
                    */

                    ScorecardsSheet.Cells["B19"].Value = "Standard Deviation (MS):";
                    ScorecardsSheet.Cells["B19"].Style.Font.Bold = true;
                    ScorecardsSheet.Cells["C19"].Value = CalculateStandardDeviationDecodeTime(TLog.DecodeTimes);
                    ScorecardsSheet.Cells["C19"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    ScorecardsSheet.Cells["C19"].Style.Border.Right.Style = ExcelBorderStyle.Thick;

                    ScorecardsSheet.Cells["B20"].Value = "Highest Decode Time (MS)";
                    ScorecardsSheet.Cells["B20"].Style.Font.Bold = true;
                    ScorecardsSheet.Cells["C20"].Value = CalculateHighestDecodeTime(TLog.DecodeTimes);
                    ScorecardsSheet.Cells["C20"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    ScorecardsSheet.Cells["C20"].Style.Border.Right.Style = ExcelBorderStyle.Thick;

                    ScorecardsSheet.Cells["B21"].Value = "Lowest Decode Time (MS):";
                    ScorecardsSheet.Cells["B21"].Style.Font.Bold = true;
                    ScorecardsSheet.Cells["C21"].Value = CalculateLowestDecodeTime(TLog.DecodeTimes);
                    ScorecardsSheet.Cells["C21"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    ScorecardsSheet.Cells["C21"].Style.Border.Right.Style = ExcelBorderStyle.Thick;

                    //styling for total range results
                    var TotalRangeResults = ScorecardsSheet.Cells["A15:C21"];
                    TotalRangeResults.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    TotalRangeResults.Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#F2F2F2"));

                    //snappy range results
                    ScorecardsSheet.Cells["B22"].Value = "Snappy Minimum Distance (In):";
                    ScorecardsSheet.Cells["B22"].Style.Font.Bold = true;
                    ScorecardsSheet.Cells["C22"].Value = CalculateClosestRange(TLog.SnappyData);
                    ScorecardsSheet.Cells["C22"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    ScorecardsSheet.Cells["C22"].Style.Border.Right.Style = ExcelBorderStyle.Thick;

                    ScorecardsSheet.Cells["B23"].Value = "Snappy Maximum Distance (In):";
                    ScorecardsSheet.Cells["B23"].Style.Font.Bold = true;
                    ScorecardsSheet.Cells["C23"].Value = CalculateFarthestRange(TLog.SnappyData);
                    ScorecardsSheet.Cells["C23"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    ScorecardsSheet.Cells["C23"].Style.Border.Right.Style = ExcelBorderStyle.Thick;

                    ScorecardsSheet.Cells["B24"].Value = "Snappy Decode Range (In):";
                    ScorecardsSheet.Cells["B24"].Style.Font.Bold = true;
                    ScorecardsSheet.Cells["C24"].Value = CalculateDecodeRange(TLog.SnappyData);
                    ScorecardsSheet.Cells["C24"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    ScorecardsSheet.Cells["C24"].Style.Border.Right.Style = ExcelBorderStyle.Thick;

                    ScorecardsSheet.Cells["B25"].Value = "Snappy Total Average (MS):";
                    ScorecardsSheet.Cells["B25"].Style.Font.Bold = true;
                    ScorecardsSheet.Cells["C25"].Value = CalculateAverageDecodeTime(TLog.SnappyData);
                    ScorecardsSheet.Cells["C25"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    ScorecardsSheet.Cells["C25"].Style.Border.Right.Style = ExcelBorderStyle.Thick;

                    //commented out because average 90% was no longer needed
                    /*
                    ScorecardsSheet.Cells["B"].Value = "Snappy Total Average (90%) (MS):";
                    ScorecardsSheet.Cells["B"].Style.Font.Bold = true;
                    ScorecardsSheet.Cells["B"].Value = CalculateNinetyPercentAverageDecodeTime(TLog.SummaryGraphData);
                    ScorecardsSheet.Cells["C"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    ScorecardsSheet.Cells["C"].Style.Border.Right.Style = ExcelBorderStyle.Thick;
                    */

                    ScorecardsSheet.Cells["B26"].Value = "Snappy Standard Deviation (MS):";
                    ScorecardsSheet.Cells["B26"].Style.Font.Bold = true;
                    ScorecardsSheet.Cells["C26"].Value = CalculateStandardDeviationDecodeTime(TLog.DecodeTimes);
                    ScorecardsSheet.Cells["C26"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    ScorecardsSheet.Cells["C26"].Style.Border.Right.Style = ExcelBorderStyle.Thick;

                    ScorecardsSheet.Cells["B27"].Value = "Snappy Highest Decode Time (MS)";
                    ScorecardsSheet.Cells["B27"].Style.Font.Bold = true;
                    ScorecardsSheet.Cells["C27"].Value = CalculateHighestDecodeTime(TLog.DecodeTimes);
                    ScorecardsSheet.Cells["C27"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    ScorecardsSheet.Cells["C27"].Style.Border.Right.Style = ExcelBorderStyle.Thick;

                    ScorecardsSheet.Cells["B28"].Value = "Lowest Decode Time (MS):";
                    ScorecardsSheet.Cells["B28"].Style.Font.Bold = true;
                    ScorecardsSheet.Cells["C28"].Value = CalculateLowestDecodeTime(TLog.DecodeTimes);
                    ScorecardsSheet.Cells["C28"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    ScorecardsSheet.Cells["C28"].Style.Border.Right.Style = ExcelBorderStyle.Thick;

                    //styling for snappy range results
                    var SnappyRangeResults = ScorecardsSheet.Cells["A22:C28"];
                    SnappyRangeResults.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    SnappyRangeResults.Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D9D9D9"));

                    //ScorecardsSheet.Column(1).Width = 14;
                    //ScorecardsSheet.Column(2).Width = 30;
                    //ScorecardsSheet.Column(3).Width = 35;

                    //ScorecardsSheet.Cells["C15:C28"].Style.Numberformat.Format = "0.00";
                }
                catch (Exception)
                {
                    ErrorMessageTextBox.Text = "Error Printing Out Scorecard Data In The Spreadsheet. Did You Select A Proper Log File?";
                }

                //Saves the file
                try
                {
                    var FileInfo = new System.IO.FileInfo("Compiled Result_" + TLog.ScannerName + "_" + TLog.BarcodeSampleName + "_" + DateTime.Now.ToString("yyyyMMddHHmmssffff") + ".xlsx");
                    excelPackage.SaveAs(FileInfo);
                }
                catch (Exception)
                {
                    ErrorMessageTextBox.Text = "Error Saving Excell Spreadsheet.";
                }
            }
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

        /// <summary>
        /// Calculates the average decode time from the first 90% of all decode times in a log
        /// </summary>
        /// <returns></returns>
        private float CalculateNinetyPercentAverageDecodeTime(float[,] FloatArray)
        {
            float NinetyPercentAverageDecodeTime = 0;

            NinetyPercentAverageDecodeTime = CalculateAverageDecodeTime(FloatArray) * 0.90f;

            return NinetyPercentAverageDecodeTime;
        }

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
    }
}
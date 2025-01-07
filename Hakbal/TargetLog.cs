using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Hakbal
{
    /// <summary>
    /// Class to hold all information about each selcted log file to be compiled. Each log gets added to the data set
    /// </summary>
    internal class TargetLog
    {
        //Class Variables
        public string ScannerName;
        public string ScannerMake;
        public string ScannerSerialNumber;
        public string BarcodeSampleName;
        public string ScorecardGroupNumber;
        public string LogFilePath;
        public float[,] SummaryGraphData;
        public float[,] SnappyData;
        public float[] DecodeTimes;
        public float[] SnappyDecodeTimes;
        public float[] Distances;
        public float ClosestDistance;
        public float FarthestDistance;
        public float DecodeRange;
        public float TotalAverageDecodeTime;
        public float NinetyPercentAverageDecodeTime;
        public float STDDeviationDecodeTime;
        public float HighestDecodeTime;
        public float LowestDecodeTime;

        /// <summary>
        /// Default Constructor
        /// </summary>
        public TargetLog()
        {
            ScannerName = "blank";
            ScannerMake = "blank";
            ScannerSerialNumber = "blank";
            BarcodeSampleName = "blank";
            ScorecardGroupNumber = "0";
            LogFilePath = "blank";
            SummaryGraphData = new float[0, 0];
            SnappyData = new float[0, 0];
            DecodeTimes = new float[0];
            SnappyDecodeTimes = new float[0];
            Distances = new float[0];
            ClosestDistance = 0;
            FarthestDistance = 0;
            DecodeRange = 0;
            TotalAverageDecodeTime = 0;
            NinetyPercentAverageDecodeTime = 0;
            STDDeviationDecodeTime = 0;
            HighestDecodeTime = 0;
            LowestDecodeTime = 0;
        }

        /// <summary>
        /// Override of ToString to print out formatted data about the class
        /// </summary>
        /// <returns></returns>
        public override string ToString()
        {
            //only want to return basic info for the data set list box
            return "Make: " + ScannerMake + ", Scanner: " + ScannerName + ", Sample: " + BarcodeSampleName + ", Group: " + ScorecardGroupNumber;
        }
    }
}
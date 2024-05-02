using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Hakbal
{
    /// <summary>
    /// Class to hold all information about the selcted log file to be compiled
    /// </summary>
    internal class TargetLog
    {
        public string ScannerName;

        public string ScannerMake;

        public string ScannerSerialNumber;

        public string BarcodeSampleName;

        public string BarcodeSampleType;

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

        public TargetLog()
        {
            ScannerName = "blank";
            ScannerMake = "blank";
            ScannerSerialNumber = "blank";
            BarcodeSampleName = "blank";
            BarcodeSampleType = "blank";
            LogFilePath = "blank";
            SummaryGraphData = new float[0, 0];
            DecodeTimes = new float[0];
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
    }
}
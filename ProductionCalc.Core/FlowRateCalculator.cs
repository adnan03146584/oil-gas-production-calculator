using System;

namespace ProductionCalc.Core
{
    public class LiquidFlowCalculator
    {
        // Inputs
        public double MeterStart { get; set; }       // Start meter reading
        public double MeterEnd { get; set; }         // End meter reading
        public double Bsw { get; set; }              // BS&W %
        public double ShrinkageFactor { get; set; }  // Shrinkage / corrected meter factor
        public double TimeFactor { get; set; } = 24; // Default = 24 hours/day

        // Step 1: BLPD (Barrels of Liquid Per Day)
        public double CalculateBLPD()
        {
            double meterDiff = MeterEnd - MeterStart;
            return meterDiff * TimeFactor * ShrinkageFactor;
        }

        // Step 2: BOPD (Barrels of Oil Per Day)
        public double CalculateBOPD()
        {
            double blpd = CalculateBLPD();
            double percentOil = 100 - Bsw;
            return blpd * (percentOil / 100.0);
        }

        // Step 3: BWPD (Barrels of Water Per Day)
        public double CalculateBWPD()
        {
            double blpd = CalculateBLPD();
            return blpd * (Bsw / 100.0);
        }

        // Report
        public LiquidFlowResult CalculateReport()
        {
            return new LiquidFlowResult
            {
                BLPD = CalculateBLPD(),
                BOPD = CalculateBOPD(),
                BWPD = CalculateBWPD()
            };
        }
    }

    public class LiquidFlowResult
    {
        public double BLPD { get; set; }
        public double BOPD { get; set; }
        public double BWPD { get; set; }
    }
}

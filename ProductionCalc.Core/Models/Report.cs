using System;
using System.Text.Json.Serialization;

namespace ProductionCalc.Core.Models
{
    public class Report
    {
        public int Id { get; set; }

        // --- Identification & Paths ---
        public string ReportTitle { get; set; } = string.Empty;
        public string OperatorPath { get; set; } = string.Empty;
        public string SupervisorPath { get; set; } = string.Empty;

        // --- Personnel ---
        public string OperatorName { get; set; } = string.Empty;
        public string SupervisorComment { get; set; } = string.Empty;

        // --- Production Data ---
        public double ChokeSize { get; set; } = 0.0;
        public double GasQ { get; set; } = 0.0;       // Gas production (MMSCFD)
        public double LiqBOPD { get; set; } = 0.0;    // Liquid/oil production (BOPD)

        // --- Administrative ---
        public string Remarks { get; set; } = string.Empty;
        public string Status { get; set; } = "Pending";
        public DateTime DateSubmitted { get; set; } = DateTime.Now;

        // --- Flexible metadata (for future extension) ---
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
        public string MetadataJson { get; set; } = string.Empty;

        // --- Helper: Quick Summary (not serialized) ---
        [JsonIgnore]
        public string Summary => $"{ReportTitle} | {OperatorName} | {Status}";
    }
}

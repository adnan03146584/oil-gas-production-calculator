using OfficeOpenXml;
using OfficeOpenXml.Style;
using ProductionCalc.Core.Models;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text.Json;

namespace ProductionCalc.Core.Services
{
    public class ReportService
    {
        private readonly string _storagePath;

        public ReportService()
        {
            var baseDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "ProductionCalc");
            Directory.CreateDirectory(baseDir);
            _storagePath = Path.Combine(baseDir, "reports.json");
        }

        // --- ✅ Get all reports
        public List<Report> GetAllReports()
        {
            try { return LoadReports(); }
            catch { return new List<Report>(); }
        }

        // --- ✅ Save or update report
        public void SaveReport(Report report)
        {
            var reports = LoadReports();

            if (report.Id == 0)
            {
                int nextId = reports.Any() ? reports.Max(r => r.Id) + 1 : 1;
                report.Id = nextId;
            }

            var existing = reports.FirstOrDefault(r => r.Id == report.Id);
            if (existing != null)
                reports.Remove(existing);

            reports.Add(report);
            SaveReports(reports);
        }

        // --- ✅ Update report
        public bool UpdateReport(int id, string status, string supervisorComment)
        {
            var reports = LoadReports();
            var target = reports.FirstOrDefault(r => r.Id == id);
            if (target == null) return false;

            target.Status = status;
            target.SupervisorComment = supervisorComment;
            target.SupervisorPath = Environment.UserName;
            SaveReports(reports);
            return true;
        }

        // --- ✅ Load all reports
        public List<Report> LoadReports()
        {
            try
            {
                if (!File.Exists(_storagePath))
                    return new List<Report>();

                var json = File.ReadAllText(_storagePath);
                if (string.IsNullOrWhiteSpace(json))
                    return new List<Report>();

                var opts = new JsonSerializerOptions { PropertyNameCaseInsensitive = true };
                return JsonSerializer.Deserialize<List<Report>>(json, opts) ?? new List<Report>();
            }
            catch
            {
                return new List<Report>();
            }
        }

        // --- ✅ Save reports internally
        private void SaveReports(List<Report> reports)
        {
            try
            {
                var opts = new JsonSerializerOptions { WriteIndented = true };
                var json = JsonSerializer.Serialize(reports.OrderByDescending(r => r.DateSubmitted).ToList(), opts);
                File.WriteAllText(_storagePath, json);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error saving reports: {ex.Message}");
            }
        }

        // --- ✅ Delete report
        public bool DeleteReport(int id)
        {
            var reports = LoadReports();
            var existing = reports.FirstOrDefault(r => r.Id == id);
            if (existing == null) return false;

            reports.Remove(existing);
            SaveReports(reports);
            return true;
        }

        // --- ✅ Search reports
        public List<Report> SearchReports(string keyword)
        {
            var reports = LoadReports();
            if (string.IsNullOrWhiteSpace(keyword))
                return reports;

            keyword = keyword.ToLower();
            return reports.Where(r =>
                r.OperatorName.ToLower().Contains(keyword) ||
                r.Status.ToLower().Contains(keyword) ||
                r.ReportTitle.ToLower().Contains(keyword)
            ).ToList();
        }

        // --- ✅ Export all reports to Excel
        public bool ExportAllReportsToExcel(string? exportPath = null)
        {
            try
            {
                var reports = LoadReports();
                if (reports == null || !reports.Any())
                    return false;

                

                if (string.IsNullOrWhiteSpace(exportPath))
                {
                    exportPath = Path.Combine(
                        Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                        "AllReports.xlsx");
                }

                using (var package = new ExcelPackage())
                {
                    var ws = package.Workbook.Worksheets.Add("Reports");

                    // --- Header
                    ws.Cells[1, 1].Value = "ID";
                    ws.Cells[1, 2].Value = "Operator";
                    ws.Cells[1, 3].Value = "Title";
                    ws.Cells[1, 4].Value = "Status";
                    ws.Cells[1, 5].Value = "Date Submitted";
                    ws.Cells[1, 6].Value = "Supervisor Comment";

                    using (var range = ws.Cells[1, 1, 1, 6])
                    {
                        range.Style.Font.Bold = true;
                        range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        range.Style.Fill.BackgroundColor.SetColor(Color.LightGray);
                    }

                    int row = 2;
                    foreach (var report in reports)
                    {
                        ws.Cells[row, 1].Value = report.Id;
                        ws.Cells[row, 2].Value = report.OperatorName;
                        ws.Cells[row, 3].Value = report.ReportTitle;
                        ws.Cells[row, 4].Value = report.Status;
                        ws.Cells[row, 5].Value = report.DateSubmitted.ToString("g");
                        ws.Cells[row, 6].Value = report.SupervisorComment;
                        row++;
                    }

                    ws.Cells[ws.Dimension.Address].AutoFitColumns();
                    package.SaveAs(new FileInfo(exportPath));
                }

                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Export error: {ex.Message}");
                return false;
            }
        }
    }
}

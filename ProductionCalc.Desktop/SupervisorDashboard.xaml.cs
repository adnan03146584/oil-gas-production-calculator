using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using ProductionCalc.Core.Auth;
using ProductionCalc.Core.Models;
using ProductionCalc.Core.Services;

namespace ProductionCalc.Desktop
{
    public partial class SupervisorDashboard : Window
    {
        private readonly ReportService _reportService;
        private List<Report> _allReports = new();
        private readonly User _currentUser;

        // ✅ Constructor that accepts the logged-in User
        public SupervisorDashboard(User user)
        {
            InitializeComponent();
            _reportService = new ReportService();
            _currentUser = user;

            txtSupervisorDisplay.Text = $"Welcome, {_currentUser.Username}";

            LoadReports();
        }

        // Optional: parameterless constructor for XAML designer support
        public SupervisorDashboard() : this(SessionManager.CurrentUser ?? new User { Username = "Supervisor" })
        {
        }


        private void LoadReports()
        {
            try
            {
                _allReports = _reportService.GetAllReports();

                var pendingReports = _allReports
                    .Where(r => r.Status.Equals("Pending", StringComparison.OrdinalIgnoreCase)
                             || r.Status.Equals("Submitted", StringComparison.OrdinalIgnoreCase))
                    .OrderByDescending(r => r.DateSubmitted)
                    .ToList();

                ReportGrid.ItemsSource = pendingReports;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to load reports: {ex.Message}", "Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ReportGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (ReportGrid.SelectedItem is Report selected)
            {
                var sb = new StringBuilder();
                sb.AppendLine($"Report ID: {selected.Id}");
                sb.AppendLine($"Report Title: {selected.ReportTitle}");
                sb.AppendLine($"Operator: {selected.OperatorName}");
                sb.AppendLine($"Date Submitted: {selected.DateSubmitted:g}");
                sb.AppendLine($"Status: {selected.Status}");
                sb.AppendLine();
                sb.AppendLine("Operator Remarks:");
                sb.AppendLine(string.IsNullOrWhiteSpace(selected.Remarks) ? "(none)" : selected.Remarks);

                txtReportDetails.Text = sb.ToString();
            }
            else
            {
                txtReportDetails.Text = "Select a report to view details.";
            }
        }

        private void Approve_Click(object sender, RoutedEventArgs e)
        {
            UpdateReportStatus("Approved");
        }

        private void Reject_Click(object sender, RoutedEventArgs e)
        {
            UpdateReportStatus("Rejected");
        }

        private void UpdateReportStatus(string status)
        {
            if (ReportGrid.SelectedItem is not Report selected)
            {
                MessageBox.Show("Please select a report first.", "No Selection",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            string supervisorComment = txtSupervisorComments.Text.Trim();

            if (string.IsNullOrWhiteSpace(supervisorComment))
            {
                MessageBox.Show("Please enter your comment before submitting.", "Missing Comment",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // Update the report
            selected.Status = status;
            selected.SupervisorComment = supervisorComment;
            selected.SupervisorPath = SessionManager.CurrentUser?.Username ?? "Supervisor";
            selected.DateSubmitted = DateTime.Now;

            try
            {
                _reportService.SaveReport(selected); // ✅ Use SaveReport(), not UpdateReport()
                MessageBox.Show($"Report {status} successfully!", "Success",
                    MessageBoxButton.OK, MessageBoxImage.Information);

                txtSupervisorComments.Text = "";
                txtReportDetails.Text = "";
                LoadReports();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error updating report: {ex.Message}", "Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Search_TextChanged(object sender, TextChangedEventArgs e)
        {
            string query = txtSearch.Text.Trim().ToLower();

            if (string.IsNullOrEmpty(query))
            {
                ReportGrid.ItemsSource = _allReports
                    .Where(r => r.Status.Equals("Pending", StringComparison.OrdinalIgnoreCase)
                             || r.Status.Equals("Submitted", StringComparison.OrdinalIgnoreCase))
                    .OrderByDescending(r => r.DateSubmitted)
                    .ToList();
                return;
            }

            var filtered = _allReports
                .Where(r =>
                    (r.OperatorName?.ToLower().Contains(query) ?? false) ||
                    (r.DateSubmitted.ToString("g").ToLower().Contains(query)) ||
                    (r.ReportTitle?.ToLower().Contains(query) ?? false))
                .OrderByDescending(r => r.DateSubmitted)
                .ToList();

            ReportGrid.ItemsSource = filtered;
        }

        private void Logout_Click(object sender, RoutedEventArgs e)
        {
            SessionManager.EndSession();
            var login = new LoginWindow();
            login.Show();
            Close();
        }
    }
}

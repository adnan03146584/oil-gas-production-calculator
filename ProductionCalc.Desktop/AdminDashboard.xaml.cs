using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Input;
using ProductionCalc.Core.Models;
using ProductionCalc.Core.Services; // If you have services for loading data
using ProductionCalc.Core.Auth;

namespace ProductionCalc.Desktop
{
    public partial class AdminDashboard : Window
    {
        private readonly UserService _userService;
        private readonly FieldService _fieldService;
        private readonly ReportService _reportService;
        private readonly User _currentUser;

        // ✅ Constructor that accepts logged-in User
        public AdminDashboard(User user)
        {
            InitializeComponent();

            _currentUser = user;

            _userService = new UserService();
            _fieldService = new FieldService();
            _reportService = new ReportService();

            Loaded += AdminDashboard_Loaded;
        }

        // Optional parameterless constructor for XAML designer
        public AdminDashboard() : this(SessionManager.CurrentUser ?? new User { Username = "Admin" })
        {
        }

        private void AdminDashboard_Loaded(object sender, RoutedEventArgs e)
        {
            txtUserDisplay.Text = $"Logged in as: {_currentUser.Username ?? "Admin"}";

            LoadUsers();
            LoadFields();
            LoadReports();
        }

        // -------------------- USERS --------------------
        private void LoadUsers()
        {
            try
            {
                var users = _userService.GetAllUsers();
                UsersGrid.ItemsSource = users;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading users: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void AddUser_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("Add User clicked – implement user creation dialog here.", "Info");
        }

        private void EditUser_Click(object sender, RoutedEventArgs e)
        {
            if (UsersGrid.SelectedItem is User selectedUser)
            {
                MessageBox.Show($"Editing user: {selectedUser.Username}", "Edit User");
            }
            else
            {
                MessageBox.Show("Please select a user to edit.", "Warning");
            }
        }

        private void DeleteUser_Click(object sender, RoutedEventArgs e)
        {
            if (UsersGrid.SelectedItem is User selectedUser)
            {
                var confirm = MessageBox.Show($"Are you sure you want to delete '{selectedUser.Username}'?",
                    "Confirm Delete", MessageBoxButton.YesNo, MessageBoxImage.Question);

                if (confirm == MessageBoxResult.Yes)
                {
                    _userService.DeleteUser(selectedUser.Id);
                    LoadUsers();
                }
            }
            else
            {
                MessageBox.Show("Please select a user to delete.", "Warning");
            }
        }

        // -------------------- FIELDS --------------------
        private void LoadFields()
        {
            try
            {
                var fields = _fieldService.GetAllFields();
                FieldsGrid.ItemsSource = fields;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading fields: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void AddField_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("Add Field clicked – implement add field dialog.", "Info");
        }

        private void EditField_Click(object sender, RoutedEventArgs e)
        {
            if (FieldsGrid.SelectedItem is Field selectedField)
            {
                MessageBox.Show($"Editing field: {selectedField.Name}", "Edit Field");
            }
            else
            {
                MessageBox.Show("Please select a field to edit.", "Warning");
            }
        }

        private void DeleteField_Click(object sender, RoutedEventArgs e)
        {
            if (FieldsGrid.SelectedItem is Field selectedField)
            {
                var confirm = MessageBox.Show($"Delete field '{selectedField.Name}'?",
                    "Confirm Delete", MessageBoxButton.YesNo, MessageBoxImage.Question);

                if (confirm == MessageBoxResult.Yes)
                {
                    _fieldService.DeleteField(selectedField.Id);
                    LoadFields();
                }
            }
            else
            {
                MessageBox.Show("Please select a field to delete.", "Warning");
            }
        }

        // -------------------- REPORTS --------------------
        private void LoadReports()
        {
            try
            {
                var reports = _reportService.GetAllReports();
                AllReportsGrid.ItemsSource = reports;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading reports: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ExportAllReports_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                _reportService.ExportAllReportsToExcel();

                MessageBox.Show("Reports exported successfully!", "Success");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Export failed: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // -------------------- LOGOUT --------------------
        private void Logout_Click(object sender, RoutedEventArgs e)
        {
            var confirm = MessageBox.Show("Are you sure you want to logout?", "Logout",
                MessageBoxButton.YesNo, MessageBoxImage.Question);

            if (confirm == MessageBoxResult.Yes)
            {
                var login = new LoginWindow();
                login.Show();
                this.Close();
            }
        }

        // Optional: enable window dragging if borderless
        private void Window_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.LeftButton == MouseButtonState.Pressed)
                DragMove();
        }
    }
}

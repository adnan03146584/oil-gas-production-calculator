using System;
using System.Windows;
using ProductionCalc.Core.Auth;  // Access AuthService, User, SessionManager

namespace ProductionCalc.Desktop
{
    public partial class LoginWindow : Window
    {
        private readonly AuthService _authService;

        public LoginWindow()
        {
            InitializeComponent();
            _authService = new AuthService(); // Initialize AuthService
        }

        private void Login_Click(object sender, RoutedEventArgs e)
        {
            string username = txtUsername.Text.Trim();
            string password = txtPassword.Password;

            try
            {
                // Authenticate user
                var user = _authService.Login(username, password);

                if (user == null)
                {
                    MessageBox.Show("Invalid username or password.",
                                    "Login Failed",
                                    MessageBoxButton.OK,
                                    MessageBoxImage.Error);
                    return;
                }

                // ✅ Start user session
                SessionManager.StartSession(user);

                MessageBox.Show($"Welcome, {user.Role} {user.Username}!",
                                "Login Successful",
                                MessageBoxButton.OK,
                                MessageBoxImage.Information);

                // Role-based redirection
                Window nextWindow;

                switch (user.Role)
                {
                    case "Operator":
                        nextWindow = new MainWindow(); // ✅ Operator Dashboard
                        break;

                    case "Supervisor":
                        nextWindow = new SupervisorDashboard(user); // ✅ Supervisor Dashboard
                        break;

                    case "Admin":
                        nextWindow = new AdminDashboard(user);
                        break;

                    default:
                        MessageBox.Show("Unknown user role. Please contact support.",
                                        "Error",
                                        MessageBoxButton.OK,
                                        MessageBoxImage.Warning);
                        return;
                }

                // ✅ Show next window and close login
                nextWindow.Show();
                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred during login: {ex.Message}\n{ex.StackTrace}",
                                "Error",
                                MessageBoxButton.OK,
                                MessageBoxImage.Error);
            }
        }
    }
}

using ProductionCalc.Core.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;

namespace ProductionCalc.Core.Auth
{
    public class AuthService
    {
        private readonly List<User> _users;

        public AuthService()
        {
            // Preloaded sample users (mock database)
            _users = new List<User>
            {
                new User
                {
                    Username = "operator1",
                    PasswordHash = HashPassword("op123"),
                    Role = "Operator",
                    AssignedField = "Field A"
                },
                new User
                {
                    Username = "supervisor1",
                    PasswordHash = HashPassword("sup123"),
                    Role = "Supervisor",
                    AssignedField = "Field A"
                },
                new User
                {
                    Username = "admin",
                    PasswordHash = HashPassword("admin123"),
                    Role = "Admin",
                    AssignedField = "HQ"
                }
            };
        }

        public User? Login(string username, string password)
        {
            string hashedPassword = HashPassword(password);
            var user = _users.FirstOrDefault(u =>
                u.Username.Equals(username, StringComparison.OrdinalIgnoreCase) &&
                u.PasswordHash == hashedPassword &&
                u.IsActive);

            if (user != null)
                user.LastLogin = DateTime.Now;

            return user;
        }

        // Basic SHA256 password hashing
        private string HashPassword(string password)
        {
            using (var sha = SHA256.Create())
            {
                var bytes = sha.ComputeHash(Encoding.UTF8.GetBytes(password));
                return BitConverter.ToString(bytes).Replace("-", "").ToLower();
            }
        }
    }
}

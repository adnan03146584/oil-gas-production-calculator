using ProductionCalc.Core.Models;
using System.Collections.Generic;
using System.Linq;

namespace ProductionCalc.Core.Services
{
    public class UserService
    {
        private readonly List<User> _users;

        public UserService()
        {
            // Mock data
            _users = new List<User>
            {
                new User { Id = 1, Username = "admin", Role = "Administrator", AssignedField = "HQ" },
                new User { Id = 2, Username = "john.doe", Role = "Operator", AssignedField = "Field A" },
                new User { Id = 3, Username = "sarah", Role = "Supervisor", AssignedField = "Field B" }
            };
        }

        public List<User> GetAllUsers() => _users;

        public void DeleteUser(int id)
        {
            var user = _users.FirstOrDefault(u => u.Id == id);
            if (user != null)
                _users.Remove(user);
        }
    }
}

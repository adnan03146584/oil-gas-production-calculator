using ProductionCalc.Core.Models;
using System;

namespace ProductionCalc.Core.Auth
{
    public static class SessionManager
    {
        private static User? _currentUser;

        public static void StartSession(User user)
        {
            _currentUser = user;
            _currentUser.LastLogin = DateTime.Now;
        }

        public static User? CurrentUser => _currentUser;

        public static bool IsLoggedIn => _currentUser != null;

        public static void EndSession()
        {
            _currentUser = null;
        }
    }
}

using ProductionCalc.Core.Models;

namespace ProductionCalc.Core.Auth
{
    public static class AuthManager
    {
        public static User? CurrentUser { get; private set; }

        public static bool Login(User user)
        {
            if (user == null) return false;
            CurrentUser = user;
            user.LastLogin = DateTime.Now;
            return true;
        }

        public static void Logout()
        {
            CurrentUser = null;
        }

        public static bool IsLoggedIn => CurrentUser != null;
    }
}

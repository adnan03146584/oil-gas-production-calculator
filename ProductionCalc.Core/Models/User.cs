namespace ProductionCalc.Core.Models
{
    public class User
    {
        public int Id { get; set; }

        // --- Identity ---
        public string Username { get; set; } = string.Empty;
        public string PasswordHash { get; set; } = string.Empty;

        // --- Authorization ---
        public string Role { get; set; } = string.Empty; // Operator, Supervisor, Admin
        public string AssignedField { get; set; } = string.Empty;

        // --- Status & Audit ---
        public bool IsActive { get; set; } = true;
        public DateTime? LastLogin { get; set; }
        public DateTime CreatedAt { get; set; } = DateTime.Now;
    }
}

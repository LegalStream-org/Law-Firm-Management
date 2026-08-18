namespace Law_Firm_Management.Models
{
    public class ClientContactVm
    {
        public int ClientContactId { get; set; }
        public int ClientId { get; set; }
        public string? ContactType { get; set; }
        public string? ContactName { get; set; }
        public string? Title { get; set; }
        public string? Email { get; set; }
        public string? Phone { get; set; }
        public string? PreferredContactMethod { get; set; }
        public bool IsActive { get; set; }
        public string? Notes { get; set; }
    }
}
namespace Law_Firm_Management.Models
{
    public class LawFirmContactVm
    {
        public int LawFirmContactId { get; set; }
        public int LawFirmId { get; set; }
        public string ContactType { get; set; }
        public string ContactName { get; set; }
        public string Title { get; set; }
        public string Email { get; set; }
        public string Phone { get; set; }
        public string PreferredContactMethod { get; set; }
        public bool IsActive { get; set; }
        public string Notes { get; set; }
    }

}

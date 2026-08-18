namespace Law_Firm_Management.Models
{
    public class ClientListRowVm
    {
        public int ClientId { get; set; }
        public string? ClientName { get; set; }
        public string? ShortName { get; set; }
        public string? Status { get; set; }
        public int ContactCount { get; set; }
        public DateTime? UpdatedDate { get; set; }
    }
}
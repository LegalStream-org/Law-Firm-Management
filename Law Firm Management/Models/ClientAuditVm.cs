namespace Law_Firm_Management.Models
{
    public class ClientAuditVm
    {
        public int ClientAuditId { get; set; }
        public int ClientId { get; set; }
        public string? EntityType { get; set; }
        public int? EntityId { get; set; }
        public string? ActionType { get; set; }
        public string? AuditText { get; set; }
        public string? EnteredBy { get; set; }
        public DateTime EnteredDate { get; set; }
    }
}
namespace Law_Firm_Management.Models
{
    public class LawFirmAuditVm
    {
        public int LawFirmAuditId { get; set; }
        public int LawFirmId { get; set; }
        public string EntityType { get; set; }
        public int? EntityId { get; set; }
        public string ActionType { get; set; }
        public string ChangedBy { get; set; }
        public DateTime ChangedDate { get; set; }
        public string ChangeSummary { get; set; }
        public string OldValue { get; set; }
        public string NewValue { get; set; }
    }
}

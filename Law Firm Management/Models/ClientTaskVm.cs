namespace Law_Firm_Management.Models
{
    public class ClientTaskVm
    {
        public int ClientTaskId { get; set; }
        public int ClientId { get; set; }
        public string? TaskType { get; set; }
        public string? TaskName { get; set; }
        public string? Frequency { get; set; }
        public string? DueDay { get; set; }
        public string? AssignedTo { get; set; }
        public string? Status { get; set; }
        public string? DeliveryMethod { get; set; }
        public int? ClientSftpConfigId { get; set; }
        public string? SftpConfigName { get; set; }
        public string? Notes { get; set; }
        public DateTime? LastCompletedDate { get; set; }
        public DateTime? NextDueDate { get; set; }
    }
}
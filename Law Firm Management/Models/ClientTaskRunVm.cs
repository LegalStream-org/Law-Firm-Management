namespace Law_Firm_Management.Models
{
    public class ClientTaskRunVm
    {
        public int ClientTaskRunId { get; set; }
        public int ClientTaskId { get; set; }
        public DateTime RunDate { get; set; }
        public string? RunStatus { get; set; }
        public string? FileName { get; set; }
        public string? Notes { get; set; }
        public string? CompletedBy { get; set; }
    }
}
namespace Law_Firm_Management.Models
{
    public class ClientDocumentVm
    {
        public int ClientDocumentId { get; set; }
        public int ClientId { get; set; }
        public string? DocumentType { get; set; }
        public string? DocumentName { get; set; }
        public string? OriginalFileName { get; set; }
        public string? ContentType { get; set; }
        public long? FileSizeBytes { get; set; }
        public string? Notes { get; set; }
        public DateTime? EffectiveDate { get; set; }
        public DateTime? ExpirationDate { get; set; }
        public bool IsActive { get; set; }
        public string? UploadedBy { get; set; }
        public DateTime UploadedDate { get; set; }
    }
}
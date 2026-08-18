using System.ComponentModel.DataAnnotations;

namespace Law_Firm_Management.Models
{
    public class ClientEditVm
    {
        public int ClientId { get; set; }

        [Required]
        [StringLength(200)]
        public string? ClientName { get; set; }

        [StringLength(100)]
        public string? ShortName { get; set; }

        [Required]
        public string Status { get; set; } = "Active";

        public string? MainPhone { get; set; }
        public string? MainEmail { get; set; }
        public string? Website { get; set; }

        public string? Address1 { get; set; }
        public string? Address2 { get; set; }
        public string? City { get; set; }
        public string? State { get; set; }
        public string? Zip { get; set; }

        public string? ProfileNotes { get; set; }

        public List<ClientContactVm> Contacts { get; set; } = new();
        public List<ClientNoteVm> Notes { get; set; } = new();
        public List<ClientAuditVm> AuditHistory { get; set; } = new();
        public List<ClientDocumentVm> Documents { get; set; } = new();

        public List<ClientSftpConfigVm> SftpConfigs { get; set; } = new();
        public List<ClientTaskVm> Tasks { get; set; } = new();

        public List<string> ContactTypeOptions { get; set; } = new();

        public List<ClientPortfolioVm> Portfolios { get; set; } = new();
        public List<string> NoteOptions { get; set; } = new()
        {
            "General",
            "IT",
            "Accounting",
            "Operations"
        };

        public List<string> DocumentTypeOptions { get; set; } = new()
        {
            "Contract",
            "NDA",
            "SOW",
            "Amendment",
            "Insurance",
            "W9",
            "Other"
        };

        public List<string> AuthTypeOptions { get; set; } = new()
{
    "Password",
    "Key"
};

        public List<string> TaskTypeOptions { get; set; } = new()
{
    "Export",
    "Report",
    "Remit",
    "Other"
};

        public List<string> TaskFrequencyOptions { get; set; } = new()
{
    "Daily",
    "Weekly",
    "Monthly",
    "Ad Hoc"
};

        public List<string> TaskStatusOptions { get; set; } = new()
{
    "Active",
    "Inactive",
    "Paused"
};

        public List<string> DeliveryMethodOptions { get; set; } = new()
{
    "SFTP",
    "Email",
    "Manual Upload"
};
    }
}
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
namespace Law_Firm_Management.Models
{
    public class LawFirmEditVm
    {
        public int LawFirmId { get; set; }

        [Required]
        [StringLength(200)]
        public string FirmName { get; set; }

        [StringLength(100)]
        public string ShortName { get; set; }

        [Required]
        public string Status { get; set; } = "Active";

        public string MainPhone { get; set; }
        public string MainEmail { get; set; }
        public string? Website { get; set; }

        public string Address1 { get; set; }
        public string? Address2 { get; set; }
        public string City { get; set; }
        public string State { get; set; }
        public string Zip { get; set; }

        public string? ProfileNotes { get; set; }

        public List<LawFirmStateMapVm> StateMappings { get; set; } = new();
        public List<LawFirmContactVm> Contacts { get; set; } = new();
        public List<LawFirmNoteVm> Notes { get; set; } = new();
        public List<LawFirmAuditVm> AuditHistory { get; set; } = new();

        public List<string> StateCodeOptions { get; set; } = new();
        public List<string> ContactTypeOptions { get; set; } = new();

        public List<LawFirmAccountingMapVm> AccountingMaps { get; set; } = new();
        public List<string> AccountingFirmOptions { get; set; } = new();

        public LawFirmAccountingActivityVm AccountingActivity { get; set; } = new();


        public List<string> PortfolioOptions { get; set; } = new()

{
    "CCMR",
    "NCB",
    "MTB",
    "Contingency",
    "Other",
    "Collections"
};

        public List<string> NoteOptions { get; set; } = new()
{
    "General",
    "IT",
    "Accounting",
    "Operations"
};

        public List<LawFirmDocumentVm> Documents { get; set; } = new();

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

        public class LawFirmImportRow
        {
            public string FirmName { get; set; }
            public string ShortName { get; set; }
            public string Status { get; set; }
            public string MainPhone { get; set; }
            public string MainEmail { get; set; }
            public string Website { get; set; }
            public string Address1 { get; set; }
            public string Address2 { get; set; }
            public string City { get; set; }
            public string State { get; set; }
            public string Zip { get; set; }
            public string ProfileNotes { get; set; }
        }

        public class LawFirm
        {
            public int LawFirmId { get; set; }
            public string FirmName { get; set; }
            public string ShortName { get; set; }
            public string Status { get; set; }
            public string MainPhone { get; set; }
            public string MainEmail { get; set; }
            public string Website { get; set; }
            public string Address1 { get; set; }
            public string Address2 { get; set; }
            public string City { get; set; }
            public string State { get; set; }
            public string Zip { get; set; }
            public string ProfileNotes { get; set; }

        }

    }



}

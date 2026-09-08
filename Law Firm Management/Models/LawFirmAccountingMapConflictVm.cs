namespace Law_Firm_Management.Models
{
    /// <summary>
    /// Sep 08 - LawFirm-side mirror of ClientAccountingMapConflictVm: carries a refused accounting-
    /// mapping attempt across the redirect back to Edit so the page can re-offer it behind a warning.
    /// </summary>
    public class LawFirmAccountingMapConflictVm
    {
        public int LawFirmId { get; set; }
        public string? AccountingFirmName { get; set; }
        public bool IsActive { get; set; }

        /// <summary>0 for an Add; the row being edited for an Update.</summary>
        public int LawFirmAccountingMapId { get; set; }

        /// <summary>"AddAccountingMap" or "UpdateAccountingMap" - which action Continue replays.
        /// Explicit so the view needn't re-derive it from LawFirmAccountingMapId being 0.</summary>
        public string? PostAction { get; set; }

        /// <summary>Firm(s) currently holding this mapping, already formatted for display - see
        /// LawFirmController.DescribeConflictOwners. More than one is genuinely possible here: this
        /// table predates any duplicate guard.</summary>
        public string? ExistingFirmName { get; set; }
    }
}

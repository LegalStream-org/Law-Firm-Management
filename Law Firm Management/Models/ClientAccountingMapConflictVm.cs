namespace Law_Firm_Management.Models
{
    /// <summary>
    /// Sep 08 - carries a refused accounting-mapping attempt across the redirect back to Edit, so the
    /// page can re-offer the identical post behind a confirmation naming the client that holds it.
    /// </summary>
    public class ClientAccountingMapConflictVm
    {
        public int ClientId { get; set; }
        public string? AccountingClientName { get; set; }
        public bool IsActive { get; set; }

        /// <summary>0 for an Add; the row being edited for an Update.</summary>
        public int ClientAccountingMapId { get; set; }

        /// <summary>"AddAccountingMap" or "UpdateAccountingMap" - which action Continue replays.
        /// Explicit so the view needn't re-derive it from ClientAccountingMapId being 0.</summary>
        public string? PostAction { get; set; }

        /// <summary>Client(s) currently holding this mapping, already formatted for display - see
        /// ClientController.DescribeConflictOwners.</summary>
        public string? ExistingClientName { get; set; }
    }
}

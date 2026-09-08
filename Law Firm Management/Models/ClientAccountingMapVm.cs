namespace Law_Firm_Management.Models
{
    /// <summary>
    /// Mirrors LawFirmAccountingMapVm's exact shape - the Client-side equivalent, one row per
    /// accounting-side name variant mapped to this real Client.
    /// </summary>
    public class ClientAccountingMapVm
    {
        public int ClientAccountingMapId { get; set; }
        public string? AccountingClientName { get; set; }
        public bool IsActive { get; set; }
    }
}
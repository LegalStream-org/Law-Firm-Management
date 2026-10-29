namespace Law_Firm_Management.Models
{
    /// <summary>Allowed values of dbo.LawFirm.EntityType. Named "record type" in C# to avoid
    /// confusion with LawFirmAudit's unrelated EntityType column.</summary>
    public static class LawFirmRecordType
    {
        public const string Firm = "Firm";
        public const string Agency = "Agency";
    }

    /// <summary>Screen wording per record type, so the shared LawFirm views can show either.</summary>
    public record LawFirmRecordLabels(string Short, string ShortPlural, string Full, string FullPlural)
    {
        public static readonly LawFirmRecordLabels Firm = new("Firm", "Firms", "Law Firm", "Law Firms");
        public static readonly LawFirmRecordLabels Agency = new("Agency", "Agencies", "Agency", "Agencies");

        public static LawFirmRecordLabels For(string recordType) =>
            recordType == LawFirmRecordType.Agency ? Agency : Firm;
    }
}
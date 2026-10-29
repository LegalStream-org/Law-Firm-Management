namespace Law_Firm_Management.Models
{
    /// <summary>Allowed values of dbo.LawFirm.EntityType. Named "record type" in C# to avoid
    /// confusion with LawFirmAudit's unrelated EntityType column.</summary>
    public static class LawFirmRecordType
    {
        public const string Firm = "Firm";
        public const string Agency = "Agency";
    }
}
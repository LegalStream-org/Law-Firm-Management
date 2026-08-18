namespace Law_Firm_Management.Models
{
    public class LawFirmNoteVm
    {
        public int LawFirmNoteId { get; set; }
        public int LawFirmId { get; set; }
        public string NoteType { get; set; }
        public string NoteText { get; set; }
        public string EnteredBy { get; set; }
        public DateTime EnteredDate { get; set; }

        public List<string> NoteOptions { get; set; } = new()
{
    "General",
    "IT",
    "Accounting",
    "Operations"
};
    }

}

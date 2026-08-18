namespace Law_Firm_Management.Models
{
    public class ClientNoteVm
    {
        public int ClientNoteId { get; set; }
        public int ClientId { get; set; }
        public string? NoteType { get; set; }
        public string? NoteText { get; set; }
        public string? EnteredBy { get; set; }
        public DateTime EnteredDate { get; set; }
    }
}
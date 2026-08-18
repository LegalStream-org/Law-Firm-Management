namespace Law_Firm_Management.Models
{
    
        public class FirmNcbVm
        {
            public int Id { get; set; }

            public string Firm { get; set; } = "";
            public string States { get; set; } = "";
            public int? COCO_NO { get; set; }
            public string? YGC_ID { get; set; }
            public string? FTP { get; set; }
            public string Status { get; set; } = "";
            public string? Manager { get; set; }
            public int? Number { get; set; }
            public string? Main_Contact { get; set; }

            public List<string> StatusOptions { get; set; } = new() { "Active", "Non Active" };
        }
    
}

namespace Law_Firm_Management.Models
{
        public class LawFirmListRowVm
        {
            public int LawFirmId { get; set; }
            public string FirmName { get; set; }
            public string ShortName { get; set; }
            public string Status { get; set; }
            public int StateCount { get; set; }
            public int ContactCount { get; set; }
            public DateTime? UpdatedDate { get; set; }

        public DateTime? LastInvoice { get; set; }
        public DateTime? LastRemit { get; set; }
        public DateTime? LastCostFile { get; set; }
    }
    

}

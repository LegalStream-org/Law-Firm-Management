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

            /// <summary>State maps with a COCO number; 0 means the firm can't appear on the Exceptions dashboard.</summary>
            public int CocoNoCount { get; set; }
            public DateTime? UpdatedDate { get; set; }

        public DateTime? LastInvoice { get; set; }
        public DateTime? LastRemit { get; set; }
        public DateTime? LastCostFile { get; set; }
    }


}
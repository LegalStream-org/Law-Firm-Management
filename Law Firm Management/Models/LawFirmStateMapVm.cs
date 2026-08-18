namespace Law_Firm_Management.Models
{

    public class LawFirmStateMapVm
    {
        public int LawFirmStateMapId { get; set; }
        public int LawFirmId { get; set; }

        public string StateCode { get; set; }
        public string COCO_NO { get; set; }
        public string ExportId { get; set; }
        public decimal? PlacementPercent { get; set; }

        public string Portfolio { get; set; }   // ✅ ADD THIS

        public string SftpFolderPath { get; set; }
        public bool IsActive { get; set; }
        public string Notes { get; set; }
    }
}


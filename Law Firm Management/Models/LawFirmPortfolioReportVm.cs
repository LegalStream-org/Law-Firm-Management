namespace Law_Firm_Management.Models
{
    public class LawFirmPortfolioReportVm
    {
        public string SelectedPortfolio { get; set; }
        public List<string> PortfolioOptions { get; set; } = new();
        public List<LawFirmPortfolioReportRowVm> Rows { get; set; } = new();
    }

    public class LawFirmPortfolioReportRowVm
    {
        public string StateCode { get; set; }
        public string FirmName { get; set; }
        public string Portfolio { get; set; }
        public string COCO_NO { get; set; }
        public string ExportId { get; set; }
        public string SftpFolderPath { get; set; }
        public decimal? PlacementPercent { get; set; }
    }
}
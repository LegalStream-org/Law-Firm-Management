namespace Law_Firm_Management.Models
{
    public class ClientPortfolioVm
    {
        public int ClientPortfolioId { get; set; }
        public int ClientId { get; set; }
        public string? PortfolioName { get; set; }
        public string? PortfolioCode { get; set; }
        public string? ForwNo { get; set; }
        public bool IsActive { get; set; }
        public DateTime? CreatedDate { get; set; }
        public DateTime? UpdatedDate { get; set; }
    }
}
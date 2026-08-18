using System.Collections.Generic;

namespace Law_Firm_Management.Models
{
    public class HomeDashboardVm
    {
        public int TotalLawFirms { get; set; }
        public int ActiveLawFirms { get; set; }
        public int TotalClients { get; set; }
        public int ActiveClients { get; set; }

        public List<LawFirmListRowVm> RecentLawFirms { get; set; } = new();
        public List<ClientListRowVm> RecentClients { get; set; } = new();
    }
}
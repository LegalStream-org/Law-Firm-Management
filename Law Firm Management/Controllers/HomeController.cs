using Law_Firm_Management.Models;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;

namespace Law_Firm_Management.Controllers
{
    public class HomeController : Controller
    {
        private readonly string _connectionString;

        public HomeController(IConfiguration configuration)
        {
            _connectionString = configuration.GetConnectionString("Connection")
                ?? throw new InvalidOperationException("Missing SQL Connection connection string.");
        }

        public IActionResult Index()
        {
            var vm = new HomeDashboardVm();

            using var conn = new SqlConnection(_connectionString);
            conn.Open();

            using (var cmd = new SqlCommand("SELECT COUNT(*) FROM dbo.LawFirm;", conn))
                vm.TotalLawFirms = Convert.ToInt32(cmd.ExecuteScalar());

            using (var cmd = new SqlCommand("SELECT COUNT(*) FROM dbo.LawFirm WHERE Status = 'Active';", conn))
                vm.ActiveLawFirms = Convert.ToInt32(cmd.ExecuteScalar());

            using (var cmd = new SqlCommand("SELECT COUNT(*) FROM dbo.Client;", conn))
                vm.TotalClients = Convert.ToInt32(cmd.ExecuteScalar());

            using (var cmd = new SqlCommand("SELECT COUNT(*) FROM dbo.Client WHERE Status = 'Active';", conn))
                vm.ActiveClients = Convert.ToInt32(cmd.ExecuteScalar());

            using (var cmd = new SqlCommand(@"
SELECT TOP 5
    LawFirmId,
    FirmName,
    ShortName,
    Status,
    0 AS StateCount,
    0 AS ContactCount,
    ISNULL(UpdatedDate, CreatedDate) AS UpdatedDate
FROM dbo.LawFirm
ORDER BY ISNULL(UpdatedDate, CreatedDate) DESC;", conn))
            using (var rdr = cmd.ExecuteReader())
            {
                while (rdr.Read())
                {
                    vm.RecentLawFirms.Add(new LawFirmListRowVm
                    {
                        LawFirmId = Convert.ToInt32(rdr["LawFirmId"]),
                        FirmName = rdr["FirmName"]?.ToString(),
                        ShortName = rdr["ShortName"]?.ToString(),
                        Status = rdr["Status"]?.ToString(),
                        StateCount = 0,
                        ContactCount = 0,
                        UpdatedDate = rdr["UpdatedDate"] == DBNull.Value ? null : (DateTime?)Convert.ToDateTime(rdr["UpdatedDate"])
                    });
                }
            }

            using (var cmd = new SqlCommand(@"
SELECT TOP 5
    ClientId,
    ClientName,
    ShortName,
    Status,
    0 AS ContactCount,
    ISNULL(UpdatedDate, CreatedDate) AS UpdatedDate
FROM dbo.Client
ORDER BY ISNULL(UpdatedDate, CreatedDate) DESC;", conn))
            using (var rdr = cmd.ExecuteReader())
            {
                while (rdr.Read())
                {
                    vm.RecentClients.Add(new ClientListRowVm
                    {
                        ClientId = Convert.ToInt32(rdr["ClientId"]),
                        ClientName = rdr["ClientName"]?.ToString(),
                        ShortName = rdr["ShortName"]?.ToString(),
                        Status = rdr["Status"]?.ToString(),
                        ContactCount = 0,
                        UpdatedDate = rdr["UpdatedDate"] == DBNull.Value ? null : (DateTime?)Convert.ToDateTime(rdr["UpdatedDate"])
                    });
                }
            }

            return View(vm);
        }
    }
}
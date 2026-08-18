using Law_Firm_Management.Models;
using Microsoft.Data.SqlClient;
using System.Text;
using System.Data;

namespace Law_Firm_Management.Service
{
    public class FirmManagementService : IFirmManagementService
    {
        private readonly IConfiguration _configuration;

        public FirmManagementService(IConfiguration configuration)
        {
            _configuration = configuration;
        }

        private string ConnectionString =>
            _configuration.GetConnectionString("Connection");

        public async Task<List<FirmNcbVm>> GetFirmsAsync(string? state = null, string? status = null, string? search = null)
        {
            var list = new List<FirmNcbVm>();

            var sql = new StringBuilder(@"
SELECT
    Id,
    Firm,
    States,
    COCO_NO,
    YGC_ID,
    FTP,
    Status,
    Manager,
    [#],
    Main_Contact
FROM LegalStream.dbo.Firms_NCB
WHERE 1 = 1
");

            if (!string.IsNullOrWhiteSpace(state))
                sql.Append(" AND States = @States");

            if (!string.IsNullOrWhiteSpace(status))
            {
                if (status == "Non Active")
                    sql.Append(" AND ISNULL(Status, '') IN ('Non Active', 'No')");
                else
                    sql.Append(" AND Status = @Status");
            }

            if (!string.IsNullOrWhiteSpace(search))
            {
                sql.Append(@"
 AND (
        Firm LIKE @Search
        OR States LIKE @Search
        OR ISNULL(YGC_ID,'') LIKE @Search
        OR ISNULL(FTP,'') LIKE @Search
        OR ISNULL(Status,'') LIKE @Search
        OR ISNULL(Manager,'') LIKE @Search
        OR ISNULL(Main_Contact,'') LIKE @Search
     )");
            }

            sql.Append(" ORDER BY States, Firm, [#]");

            using var conn = new SqlConnection(ConnectionString);
            using var cmd = new SqlCommand(sql.ToString(), conn);

            if (!string.IsNullOrWhiteSpace(state))
                cmd.Parameters.AddWithValue("@States", state);

            if (!string.IsNullOrWhiteSpace(status) && status != "Non Active")
                cmd.Parameters.AddWithValue("@Status", status);

            if (!string.IsNullOrWhiteSpace(search))
                cmd.Parameters.AddWithValue("@Search", "%" + search + "%");

            await conn.OpenAsync();
            using var rdr = await cmd.ExecuteReaderAsync();

            while (await rdr.ReadAsync())
            {
                list.Add(new FirmNcbVm
                {
                    Id = rdr["Id"] == DBNull.Value ? 0 : Convert.ToInt32(rdr["Id"]),
                    Firm = rdr["Firm"]?.ToString() ?? "",
                    States = rdr["States"]?.ToString() ?? "",
                    COCO_NO = rdr["COCO_NO"] == DBNull.Value ? null : Convert.ToInt32(rdr["COCO_NO"]),
                    YGC_ID = rdr["YGC_ID"] == DBNull.Value ? null : rdr["YGC_ID"].ToString(),
                    FTP = rdr["FTP"] == DBNull.Value ? null : rdr["FTP"].ToString(),
                    Status = NormalizeStatus(rdr["Status"]?.ToString()),
                    Manager = rdr["Manager"] == DBNull.Value ? null : rdr["Manager"].ToString(),
                    Number = rdr["#"] == DBNull.Value ? null : Convert.ToInt32(rdr["#"]),
                    Main_Contact = rdr["Main_Contact"] == DBNull.Value ? null : rdr["Main_Contact"].ToString()
                });
            }

            return list;
        }

        public async Task<FirmNcbVm?> GetFirmByIdAsync(int id)
        {
            const string sql = @"
SELECT TOP 1
    Id,
    Firm,
    States,
    COCO_NO,
    YGC_ID,
    FTP,
    Status,
    Manager,
    [#],
    Main_Contact
FROM LegalStream.dbo.Firms_NCB
WHERE Id = @Id;";

            using var conn = new SqlConnection(ConnectionString);
            using var cmd = new SqlCommand(sql, conn);
            cmd.Parameters.AddWithValue("@Id", id);

            await conn.OpenAsync();
            using var rdr = await cmd.ExecuteReaderAsync();

            if (await rdr.ReadAsync())
            {
                return new FirmNcbVm
                {
                    Id = rdr["Id"] == DBNull.Value ? 0 : Convert.ToInt32(rdr["Id"]),
                    Firm = rdr["Firm"]?.ToString() ?? "",
                    States = rdr["States"]?.ToString() ?? "",
                    COCO_NO = rdr["COCO_NO"] == DBNull.Value ? null : Convert.ToInt32(rdr["COCO_NO"]),
                    YGC_ID = rdr["YGC_ID"] == DBNull.Value ? null : rdr["YGC_ID"].ToString(),
                    FTP = rdr["FTP"] == DBNull.Value ? null : rdr["FTP"].ToString(),
                    Status = NormalizeStatus(rdr["Status"]?.ToString()),
                    Manager = rdr["Manager"] == DBNull.Value ? null : rdr["Manager"].ToString(),
                    Number = rdr["#"] == DBNull.Value ? null : Convert.ToInt32(rdr["#"]),
                    Main_Contact = rdr["Main_Contact"] == DBNull.Value ? null : rdr["Main_Contact"].ToString()
                };
            }

            return null;
        }

        public async Task CreateFirmAsync(FirmNcbVm model)
        {
            const string sql = @"
INSERT INTO LegalStream.dbo.Firms_NCB
(
    Firm,
    States,
    COCO_NO,
    YGC_ID,
    FTP,
    Status,
    Manager,
    [#],
    Main_Contact
)
VALUES
(
    @Firm,
    @States,
    @COCO_NO,
    @YGC_ID,
    @FTP,
    @Status,
    @Manager,
    @Number,
    @Main_Contact
);";

            using var conn = new SqlConnection(ConnectionString);
            using var cmd = new SqlCommand(sql, conn);

            cmd.Parameters.AddWithValue("@Firm", (object?)model.Firm ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@States", (object?)model.States ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@COCO_NO", (object?)model.COCO_NO ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@YGC_ID", (object?)model.YGC_ID ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@FTP", (object?)model.FTP ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@Status", NormalizeStatus(model.Status));
            cmd.Parameters.AddWithValue("@Manager", (object?)model.Manager ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@Number", (object?)model.Number ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@Main_Contact", (object?)model.Main_Contact ?? DBNull.Value);

            await conn.OpenAsync();
            await cmd.ExecuteNonQueryAsync();
        }

        public async Task UpdateFirmAsync(FirmNcbVm model)
        {
            const string sql = @"
UPDATE LegalStream.dbo.Firms_NCB
SET
    Firm = @Firm,
    States = @States,
    COCO_NO = @COCO_NO,
    YGC_ID = @YGC_ID,
    FTP = @FTP,
    Status = @Status,
    Manager = @Manager,
    [#] = @Number,
    Main_Contact = @Main_Contact
WHERE Id = @Id;";

            using var conn = new SqlConnection(ConnectionString);
            using var cmd = new SqlCommand(sql, conn);

            cmd.Parameters.AddWithValue("@Id", model.Id);
            cmd.Parameters.AddWithValue("@Firm", (object?)model.Firm ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@States", (object?)model.States ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@COCO_NO", (object?)model.COCO_NO ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@YGC_ID", (object?)model.YGC_ID ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@FTP", (object?)model.FTP ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@Status", NormalizeStatus(model.Status));
            cmd.Parameters.AddWithValue("@Manager", (object?)model.Manager ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@Number", (object?)model.Number ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@Main_Contact", (object?)model.Main_Contact ?? DBNull.Value);

            await conn.OpenAsync();
            await cmd.ExecuteNonQueryAsync();
        }

        private static string NormalizeStatus(string? status)
        {
            if (string.IsNullOrWhiteSpace(status))
                return "Non Active";

            if (status.Equals("No", StringComparison.OrdinalIgnoreCase))
                return "Non Active";

            return status;
        }

        public async Task<DataTable> GetFirmsDataTableAsync(string? state = null, string? status = null, string? search = null)
        {
            var dt = new DataTable();

            var sql = new StringBuilder(@"
SELECT
    Firm,
    States,
    COCO_NO,
    YGC_ID,
    FTP,
    CASE 
        WHEN ISNULL(Status, '') = 'No' THEN 'Non Active'
        ELSE Status
    END AS Status,
    Manager,
    [#] AS [Number],
    Main_Contact
FROM LegalStream.dbo.Firms_NCB
WHERE 1 = 1
");

            if (!string.IsNullOrWhiteSpace(state))
                sql.Append(" AND States = @States");

            if (!string.IsNullOrWhiteSpace(status))
            {
                if (status == "Non Active")
                    sql.Append(" AND ISNULL(Status, '') IN ('Non Active', 'No')");
                else
                    sql.Append(" AND Status = @Status");
            }

            if (!string.IsNullOrWhiteSpace(search))
            {
                sql.Append(@"
 AND (
        Firm LIKE @Search
        OR States LIKE @Search
        OR ISNULL(YGC_ID,'') LIKE @Search
        OR ISNULL(FTP,'') LIKE @Search
        OR ISNULL(Status,'') LIKE @Search
        OR ISNULL(Manager,'') LIKE @Search
        OR ISNULL(Main_Contact,'') LIKE @Search
     )");
            }

            sql.Append(" ORDER BY States, Firm, [#]");

            using var conn = new SqlConnection(ConnectionString);
            using var cmd = new SqlCommand(sql.ToString(), conn);

            if (!string.IsNullOrWhiteSpace(state))
                cmd.Parameters.AddWithValue("@States", state);

            if (!string.IsNullOrWhiteSpace(status) && status != "Non Active")
                cmd.Parameters.AddWithValue("@Status", status);

            if (!string.IsNullOrWhiteSpace(search))
                cmd.Parameters.AddWithValue("@Search", "%" + search + "%");

            using var da = new SqlDataAdapter(cmd);
            await conn.OpenAsync();
            da.Fill(dt);

            return dt;
        }

        public async Task<List<string>> GetPortfolioOptionsAsync()
        {
            var list = new List<string>();

            using var conn = new SqlConnection(ConnectionString);
            using var cmd = new SqlCommand(@"
        SELECT DISTINCT Portfolio
        FROM BR_App.dbo.LawFirmStateMap
        WHERE Portfolio IS NOT NULL
          AND LTRIM(RTRIM(Portfolio)) <> ''
        ORDER BY Portfolio;", conn);

            await conn.OpenAsync();

            using var rdr = await cmd.ExecuteReaderAsync();

            while (await rdr.ReadAsync())
            {
                list.Add(rdr["Portfolio"].ToString());
            }

            return list;
        }

        public async Task<List<LawFirmPortfolioReportRowVm>> GetPortfolioReportAsync(string portfolio)
        {
            var rows = new List<LawFirmPortfolioReportRowVm>();

            using var conn = new SqlConnection(ConnectionString);
            using var cmd = new SqlCommand(@"
        SELECT 
            b.StateCode,
            a.FirmName,
            b.Portfolio,
            b.COCO_NO,
            b.ExportId,
            b.SftpFolderPath,
            b.PlacementPercent
        FROM BR_App.dbo.LawFirm a
        INNER JOIN BR_App.dbo.LawFirmStateMap b 
            ON b.LawFirmId = a.LawFirmId
        WHERE (@Portfolio IS NULL OR @Portfolio = '' OR b.Portfolio = @Portfolio)
        AND a.IsActive = 1
        AND b.IsActive = 1

        ORDER BY b.StateCode, a.FirmName;", conn);

            cmd.Parameters.AddWithValue("@Portfolio", (object)portfolio ?? DBNull.Value);

            await conn.OpenAsync();

            using var rdr = await cmd.ExecuteReaderAsync();

            while (await rdr.ReadAsync())
            {
                rows.Add(new LawFirmPortfolioReportRowVm
                {
                    StateCode = rdr["StateCode"]?.ToString(),
                    FirmName = rdr["FirmName"]?.ToString(),
                    Portfolio = rdr["Portfolio"]?.ToString(),
                    COCO_NO = rdr["COCO_NO"]?.ToString(),
                    ExportId = rdr["ExportId"]?.ToString(),
                    SftpFolderPath = rdr["SftpFolderPath"]?.ToString(),
                    PlacementPercent = rdr["PlacementPercent"] == DBNull.Value ? null : Convert.ToDecimal(rdr["PlacementPercent"])
                });
            }

            return rows;
        }

    }
}

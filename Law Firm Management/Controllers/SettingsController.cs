using Law_Firm_Management.Models;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;

namespace Law_Firm_Management.Controllers
{
    // Settings-area controller. Client Type is the first lookup managed here;
    // follow the same pattern for future programmable lists.
    public class SettingsController : Controller
    {
        private readonly string _connectionString;

        public SettingsController(IConfiguration configuration)
        {
            _connectionString = configuration.GetConnectionString("Connection")
                ?? throw new InvalidOperationException("Missing SQL Connection connection string.");
        }

        [HttpGet]
        public IActionResult ClientTypes()
        {
            var rows = GetClientTypes();
            return View(rows);
        }

        [HttpGet]
        public IActionResult CreateClientType()
        {
            return View(new ClientTypeEditVm { IsActive = true });
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public IActionResult CreateClientType(ClientTypeEditVm vm)
        {
            if (!ModelState.IsValid)
                return View(vm);

            if (ClientTypeNameExists(vm.TypeName, excludeId: null))
            {
                ModelState.AddModelError(nameof(vm.TypeName), "A client type with this name already exists.");
                return View(vm);
            }

            using var conn = new SqlConnection(_connectionString);
            using var cmd = new SqlCommand(@"
INSERT INTO dbo.ClientType
(
    TypeName,
    Description,
    SuppressCommissionByDefault,
    IsActive,
    CreatedBy,
    CreatedDate,
    UpdatedBy,
    UpdatedDate
)
VALUES
(
    @TypeName,
    @Description,
    @SuppressCommissionByDefault,
    @IsActive,
    @CreatedBy,
    GETDATE(),
    @UpdatedBy,
    GETDATE()
);", conn);

            cmd.Parameters.AddWithValue("@TypeName", vm.TypeName!.Trim());
            cmd.Parameters.AddWithValue("@Description", (object?)vm.Description ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@SuppressCommissionByDefault", vm.SuppressCommissionByDefault);
            cmd.Parameters.AddWithValue("@IsActive", vm.IsActive);
            cmd.Parameters.AddWithValue("@CreatedBy", User?.Identity?.Name ?? "System");
            cmd.Parameters.AddWithValue("@UpdatedBy", User?.Identity?.Name ?? "System");

            conn.Open();
            cmd.ExecuteNonQuery();

            TempData["Message"] = "Client type created.";
            return RedirectToAction(nameof(ClientTypes));
        }

        [HttpGet]
        public IActionResult EditClientType(int id)
        {
            var vm = GetClientType(id);
            if (vm == null)
                return NotFound();

            return View(vm);
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public IActionResult EditClientType(ClientTypeEditVm vm)
        {
            if (!ModelState.IsValid)
                return View(vm);

            if (ClientTypeNameExists(vm.TypeName, excludeId: vm.ClientTypeId))
            {
                ModelState.AddModelError(nameof(vm.TypeName), "A client type with this name already exists.");
                return View(vm);
            }

            using var conn = new SqlConnection(_connectionString);
            using var cmd = new SqlCommand(@"
UPDATE dbo.ClientType
SET
    TypeName = @TypeName,
    Description = @Description,
    SuppressCommissionByDefault = @SuppressCommissionByDefault,
    IsActive = @IsActive,
    UpdatedBy = @UpdatedBy,
    UpdatedDate = GETDATE()
WHERE ClientTypeId = @ClientTypeId;", conn);

            cmd.Parameters.AddWithValue("@ClientTypeId", vm.ClientTypeId);
            cmd.Parameters.AddWithValue("@TypeName", vm.TypeName!.Trim());
            cmd.Parameters.AddWithValue("@Description", (object?)vm.Description ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@SuppressCommissionByDefault", vm.SuppressCommissionByDefault);
            cmd.Parameters.AddWithValue("@IsActive", vm.IsActive);
            cmd.Parameters.AddWithValue("@UpdatedBy", User?.Identity?.Name ?? "System");

            conn.Open();
            cmd.ExecuteNonQuery();

            TempData["Message"] = "Client type updated.";
            return RedirectToAction(nameof(ClientTypes));
        }

        // ---- helpers ----

        private bool ClientTypeNameExists(string? typeName, int? excludeId)
        {
            if (string.IsNullOrWhiteSpace(typeName))
                return false;

            using var conn = new SqlConnection(_connectionString);
            using var cmd = new SqlCommand(@"
SELECT COUNT(1)
FROM dbo.ClientType
WHERE LTRIM(RTRIM(TypeName)) = LTRIM(RTRIM(@TypeName))
  AND (@ExcludeId IS NULL OR ClientTypeId <> @ExcludeId);", conn);

            cmd.Parameters.AddWithValue("@TypeName", typeName.Trim());
            cmd.Parameters.AddWithValue("@ExcludeId", (object?)excludeId ?? DBNull.Value);

            conn.Open();
            int count = (int)cmd.ExecuteScalar();
            return count > 0;
        }

        private List<ClientTypeListRowVm> GetClientTypes()
        {
            var rows = new List<ClientTypeListRowVm>();

            using var conn = new SqlConnection(_connectionString);
            using var cmd = new SqlCommand(@"
SELECT
    ct.ClientTypeId,
    ct.TypeName,
    ct.Description,
    ct.SuppressCommissionByDefault,
    ct.IsActive,
    ct.UpdatedDate,
    COUNT(c.ClientId) AS ClientCount
FROM dbo.ClientType ct
LEFT JOIN dbo.Client c ON c.ClientTypeId = ct.ClientTypeId
GROUP BY
    ct.ClientTypeId, ct.TypeName, ct.Description,
    ct.SuppressCommissionByDefault, ct.IsActive, ct.UpdatedDate
ORDER BY ct.TypeName;", conn);

            conn.Open();
            using var rdr = cmd.ExecuteReader();
            while (rdr.Read())
            {
                rows.Add(new ClientTypeListRowVm
                {
                    ClientTypeId = Convert.ToInt32(rdr["ClientTypeId"]),
                    TypeName = rdr["TypeName"]?.ToString(),
                    Description = rdr["Description"]?.ToString(),
                    SuppressCommissionByDefault = Convert.ToBoolean(rdr["SuppressCommissionByDefault"]),
                    IsActive = Convert.ToBoolean(rdr["IsActive"]),
                    UpdatedDate = rdr["UpdatedDate"] == DBNull.Value ? null : Convert.ToDateTime(rdr["UpdatedDate"]),
                    ClientCount = Convert.ToInt32(rdr["ClientCount"])
                });
            }

            return rows;
        }

        private ClientTypeEditVm? GetClientType(int id)
        {
            using var conn = new SqlConnection(_connectionString);
            using var cmd = new SqlCommand(
                "SELECT * FROM dbo.ClientType WHERE ClientTypeId = @ClientTypeId;", conn);
            cmd.Parameters.AddWithValue("@ClientTypeId", id);

            conn.Open();
            using var rdr = cmd.ExecuteReader();
            if (!rdr.Read())
                return null;

            return new ClientTypeEditVm
            {
                ClientTypeId = Convert.ToInt32(rdr["ClientTypeId"]),
                TypeName = rdr["TypeName"]?.ToString(),
                Description = rdr["Description"]?.ToString(),
                SuppressCommissionByDefault = Convert.ToBoolean(rdr["SuppressCommissionByDefault"]),
                IsActive = Convert.ToBoolean(rdr["IsActive"])
            };
        }

        // Shared by ClientController so the Client Create/Edit dropdown always reflects
        // the current Settings-managed list. Kept here so Settings stays the one owner
        // of ClientType reads/writes.
        public static List<(int ClientTypeId, string TypeName)> GetActiveClientTypeOptions(string connectionString)
        {
            var list = new List<(int, string)>();

            using var conn = new SqlConnection(connectionString);
            using var cmd = new SqlCommand(@"
SELECT ClientTypeId, TypeName
FROM dbo.ClientType
WHERE IsActive = 1
ORDER BY TypeName;", conn);

            conn.Open();
            using var rdr = cmd.ExecuteReader();
            while (rdr.Read())
            {
                list.Add((Convert.ToInt32(rdr["ClientTypeId"]), rdr["TypeName"]?.ToString() ?? ""));
            }

            return list;
        }
    }
}

using ClosedXML.Excel;
using DocumentFormat.OpenXml.InkML;
using Law_Firm_Management.Models;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Filters;
using Microsoft.Data.SqlClient;
using Microsoft.VisualBasic.FileIO;
using System.Data;
using System.IO;
using System.Text.Json;
using static Law_Firm_Management.Models.LawFirmEditVm;
using static System.Runtime.InteropServices.JavaScript.JSType;

namespace Law_Firm_Management.Controllers
{
    public partial class LawFirmController : Controller
    {
        private readonly string _connectionString;

        /// <summary>Sep 08 - TempData key the reassignment prompt travels under, named once so the
        /// writing actions and the Edit GET that reads it can't drift.</summary>
        private const string AccountingMapConflictKey = "AccountingMapConflict";

        /// <summary>Which LawFirm rows this controller manages (EntityType). AgencyController overrides it.</summary>
        protected virtual string RecordType => LawFirmRecordType.Firm;

        /// <summary>Wording for this type ("Law Firm" / "Agency"), also passed to every view as ViewData["Labels"].</summary>
        protected LawFirmRecordLabels Labels => LawFirmRecordLabels.For(RecordType);

        [NonAction]
        public override void OnActionExecuting(ActionExecutingContext context)
        {
            ViewData["Labels"] = Labels;
            base.OnActionExecuting(context);
        }

        /// <summary>VarChar(20) to match the column, so the EntityType index can be used.</summary>
        private void AddRecordTypeParam(SqlCommand cmd) =>
            cmd.Parameters.Add("@RecordType", SqlDbType.VarChar, 20).Value = RecordType;

        /// <summary>Finds a record by name within this controller's type; says so if it belongs to the other type.</summary>
        private (int? LawFirmId, string? Error) FindRecordIdByName(SqlConnection conn, string name)
        {
            using var cmd = new SqlCommand(@"
SELECT TOP 1 LawFirmId, EntityType
FROM dbo.LawFirm
WHERE FirmName = @FirmName
ORDER BY CASE WHEN EntityType = @RecordType THEN 0 ELSE 1 END, LawFirmId;", conn);

            cmd.Parameters.AddWithValue("@FirmName", name);
            AddRecordTypeParam(cmd);

            using var rdr = cmd.ExecuteReader();
            if (!rdr.Read())
                return (null, $"'{name}' was not found.");

            var foundType = rdr["EntityType"]?.ToString();
            if (!string.Equals(foundType, RecordType, StringComparison.OrdinalIgnoreCase))
                return (null, $"'{name}' is set up as {DescribeType(foundType)}, so it can't be imported here.");

            return (Convert.ToInt32(rdr["LawFirmId"]), null);
        }

        /// <summary>The record's EntityType, or null if the record doesn't exist.</summary>
        private string? GetRecordTypeById(int lawFirmId)
        {
            using var conn = new SqlConnection(_connectionString);
            using var cmd = new SqlCommand("SELECT EntityType FROM dbo.LawFirm WHERE LawFirmId = @LawFirmId;", conn);
            cmd.Parameters.AddWithValue("@LawFirmId", lawFirmId);
            conn.Open();
            return cmd.ExecuteScalar() as string;
        }

        /// <summary>Sends a record of the other type to its own pages, e.g. LawFirm/Edit/5 -> Agency/Edit/5.</summary>
        private IActionResult? RedirectIfOtherType(int lawFirmId, string action)
        {
            var type = GetRecordTypeById(lawFirmId);
            if (type == null || string.Equals(type, RecordType, StringComparison.OrdinalIgnoreCase))
                return null;

            var controller = type == LawFirmRecordType.Agency ? "Agency" : "LawFirm";
            return RedirectToAction(action, controller, new { id = lawFirmId });
        }

        private static string DescribeType(string? recordType) =>
            string.Equals(recordType, LawFirmRecordType.Agency, StringComparison.OrdinalIgnoreCase)
                ? "an agency"
                : "a law firm";

        private static readonly List<string> StateCodeOptions = new()
        {
            "AL","AK","AZ","AR","CA","CO","CT","DE","FL","GA",
            "HI","ID","IL","IN","IA","KS","KY","LA","ME","MD",
            "MA","MI","MN","MS","MO","MT","NE","NV","NH","NJ",
            "NM","NY","NC","ND","OH","OK","OR","PA","PR","RI","SC",
            "SD","TN","TX","UT","VT","VA","WA","WV","WI","WY","DC"
        };

        private static readonly List<string> ContactTypeOptions = new()
        {
            "IT",
            "Accounting",
            "Day-to-Day",
            "Escalation",
            "Operations",
            "Legal",
            "Other"
        };

        public LawFirmController(IConfiguration configuration)
        {
            _connectionString = configuration.GetConnectionString("Connection")
                ?? throw new InvalidOperationException("Missing SQL Connection connection string.");
        }



        // =========================================================
        // INDEX
        // =========================================================
        public IActionResult Index(string search, string status, string sort = "FIRM_NAME", string dir = "ASC")
        {
            var rows = GetLawFirms();

            if (!string.IsNullOrWhiteSpace(search))
            {
                rows = rows
                    .Where(x =>
                        (x.FirmName ?? "").Contains(search, StringComparison.OrdinalIgnoreCase) ||
                        (x.ShortName ?? "").Contains(search, StringComparison.OrdinalIgnoreCase))
                    .ToList();
            }

            if (!string.IsNullOrWhiteSpace(status))
            {
                rows = rows
                    .Where(x => string.Equals(x.Status, status, StringComparison.OrdinalIgnoreCase))
                    .ToList();
            }

            rows = SortRows(rows, sort, dir);

            ViewBag.Search = search ?? "";
            ViewBag.Status = status ?? "";
            ViewBag.Sort = sort;
            ViewBag.Dir = dir;

            return View(rows);
        }

        public IActionResult Export(string search, string status, string sort = "FIRM_NAME", string dir = "ASC")
        {
            var rows = GetLawFirms();

            if (!string.IsNullOrWhiteSpace(search))
            {
                rows = rows
                    .Where(x =>
                        (x.FirmName ?? "").Contains(search, StringComparison.OrdinalIgnoreCase) ||
                        (x.ShortName ?? "").Contains(search, StringComparison.OrdinalIgnoreCase))
                    .ToList();
            }

            if (!string.IsNullOrWhiteSpace(status))
            {
                rows = rows
                    .Where(x => string.Equals(x.Status, status, StringComparison.OrdinalIgnoreCase))
                    .ToList();
            }

            rows = SortRows(rows, sort, dir);

            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add(Labels.FullPlural);

            ws.Cell(1, 1).Value = $"{Labels.Short} Name";
            ws.Cell(1, 2).Value = "Short Name";
            ws.Cell(1, 3).Value = "Status";
            ws.Cell(1, 4).Value = "State Count";
            ws.Cell(1, 5).Value = "Contact Count";
            ws.Cell(1, 6).Value = "Updated Date";

            int row = 2;
            foreach (var item in rows)
            {
                ws.Cell(row, 1).Value = item.FirmName;
                ws.Cell(row, 2).Value = item.ShortName;
                ws.Cell(row, 3).Value = item.Status;
                ws.Cell(row, 4).Value = item.StateCount;
                ws.Cell(row, 5).Value = item.ContactCount;
                ws.Cell(row, 6).Value = item.UpdatedDate;
                row++;
            }

            ws.Row(1).Style.Font.Bold = true;
            ws.Columns().AdjustToContents();
            ws.RangeUsed().SetAutoFilter();

            using var stream = new MemoryStream();
            wb.SaveAs(stream);

            return File(
                stream.ToArray(),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                $"{Labels.FullPlural.Replace(" ", "")}List_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");
        }

        // =========================================================
        // CREATE
        // =========================================================
        [HttpGet]
        public IActionResult Create()
        {
            return View(new LawFirmEditVm
            {
                Status = "Active",
                StateCodeOptions = StateCodeOptions,
                ContactTypeOptions = ContactTypeOptions
            });
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public IActionResult Create(LawFirmEditVm vm)
        {
            if (!ModelState.IsValid)
            {
                vm.StateCodeOptions = StateCodeOptions;
                vm.ContactTypeOptions = ContactTypeOptions;
                return View(vm);
            }

            using var conn = new SqlConnection(_connectionString);
            conn.Open();

            if (EinAlreadyInUse(conn, vm.EIN, 0))
            {
                ModelState.AddModelError(nameof(vm.EIN), "Another law firm or agency already uses this EIN.");
                vm.StateCodeOptions = StateCodeOptions;
                vm.ContactTypeOptions = ContactTypeOptions;
                return View(vm);
            }

            using var cmd = new SqlCommand(@"
INSERT INTO dbo.LawFirm
(
    FirmName,
    ShortName,
    Status,
    MainPhone,
    MainEmail,
    Website,
    Address1,
    Address2,
    City,
    [State],
    Zip,
    EIN,
    Notes,
    EntityType,
    CreatedBy,
    CreatedDate,
    UpdatedBy,
    UpdatedDate
)
VALUES
(
    @FirmName,
    @ShortName,
    @Status,
    @MainPhone,
    @MainEmail,
    @Website,
    @Address1,
    @Address2,
    @City,
    @State,
    @Zip,
    @EIN,
    @ProfileNotes,
    @RecordType,
    @CreatedBy,
    GETDATE(),
    @UpdatedBy,
    GETDATE()
);

SELECT CAST(SCOPE_IDENTITY() AS INT);", conn);

            AddRecordTypeParam(cmd);

            cmd.Parameters.AddWithValue("@FirmName", (object?)vm.FirmName ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@ShortName", (object?)vm.ShortName ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@Status", (object?)vm.Status ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@MainPhone", (object?)vm.MainPhone ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@MainEmail", (object?)vm.MainEmail ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@Website", (object?)vm.Website ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@Address1", (object?)vm.Address1 ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@Address2", (object?)vm.Address2 ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@City", (object?)vm.City ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@State", (object?)vm.State ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@Zip", (object?)vm.Zip ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@EIN", (object?)NormalizeEin(vm.EIN) ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@ProfileNotes", (object?)vm.ProfileNotes ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@CreatedBy", User?.Identity?.Name ?? "System");
            cmd.Parameters.AddWithValue("@UpdatedBy", User?.Identity?.Name ?? "System");

            int newId = (int)cmd.ExecuteScalar();

            AddAudit(newId, "LawFirm", newId, "Insert", $"Created {Labels.Full.ToLower()} {vm.FirmName}");

            TempData["Message"] = $"{Labels.Full} created.";
            return RedirectToAction(nameof(Edit), new { id = newId });
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public IActionResult CreateNcb(FirmNcbVm vm)
        {
            if (!ModelState.IsValid)
            {
                return View("Create", vm);
            }

            using var conn = new SqlConnection(_connectionString);
            using var cmd = new SqlCommand(@"
INSERT INTO dbo.Firm_NCB
(
    Firm,
    States,
    COCO_NO,
    Number,
    YGC_ID,
    FTP,
    Status,
    Manager,
    Main_Contact
)
VALUES
(
    @Firm,
    @States,
    @COCO_NO,
    @Number,
    @YGC_ID,
    @FTP,
    @Status,
    @Manager,
    @Main_Contact
);", conn);

            cmd.Parameters.AddWithValue("@Firm", (object?)vm.Firm ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@States", (object?)vm.States ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@COCO_NO", (object?)vm.COCO_NO ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@Number", (object?)vm.Number ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@YGC_ID", (object?)vm.YGC_ID ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@FTP", (object?)vm.FTP ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@Status", (object?)vm.Status ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@Manager", (object?)vm.Manager ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@Main_Contact", (object?)vm.Main_Contact ?? DBNull.Value);

            conn.Open();
            cmd.ExecuteNonQuery();

            TempData["Message"] = "Firm added.";
            return RedirectToAction(nameof(Index));
        }

        // =========================================================
        // EDIT / DETAILS
        // =========================================================
        [HttpGet]
        public IActionResult Edit(int id)
        {
            var otherType = RedirectIfOtherType(id, nameof(Edit));
            if (otherType != null)
                return otherType;

            var vm = GetLawFirmEditVm(id);
            if (vm == null)
                return NotFound();

            // Sep 08 - set by Add/UpdateAccountingMap when the chosen name is already mapped elsewhere.
            // Read here rather than in GetLawFirmEditVm (shared with Details and the POST re-render) so
            // the prompt only appears on the page that raised it; the LawFirmId guard blocks stale ones.
            var pendingConflict = TempData[AccountingMapConflictKey] as string;
            if (!string.IsNullOrWhiteSpace(pendingConflict))
            {
                var conflictVm = JsonSerializer.Deserialize<LawFirmAccountingMapConflictVm>(pendingConflict);
                if (conflictVm != null && conflictVm.LawFirmId == id)
                    vm.PendingAccountingMapConflict = conflictVm;
            }

            return View(vm);



        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public IActionResult Edit(LawFirmEditVm vm)
        {
            if (!ModelState.IsValid)
            {
                vm.StateCodeOptions = StateCodeOptions;
                vm.ContactTypeOptions = ContactTypeOptions;
                vm.StateMappings = GetStateMappings(vm.LawFirmId);
                vm.Contacts = GetContacts(vm.LawFirmId);
                vm.Notes = GetNotes(vm.LawFirmId);
                vm.AuditHistory = GetAuditHistory(vm.LawFirmId);
                return View(vm);
            }

            using var conn = new SqlConnection(_connectionString);
            conn.Open();

            if (EinAlreadyInUse(conn, vm.EIN, vm.LawFirmId))
            {
                ModelState.AddModelError(nameof(vm.EIN), "Another law firm or agency already uses this EIN.");
                vm.StateCodeOptions = StateCodeOptions;
                vm.ContactTypeOptions = ContactTypeOptions;
                vm.StateMappings = GetStateMappings(vm.LawFirmId);
                vm.Contacts = GetContacts(vm.LawFirmId);
                vm.Notes = GetNotes(vm.LawFirmId);
                vm.AuditHistory = GetAuditHistory(vm.LawFirmId);
                return View(vm);
            }

            using var cmd = new SqlCommand(@"
UPDATE dbo.LawFirm
SET
    FirmName = @FirmName,
    ShortName = @ShortName,
    Status = @Status,
    MainPhone = @MainPhone,
    MainEmail = @MainEmail,
    Website = @Website,
    Address1 = @Address1,
    Address2 = @Address2,
    City = @City,
    [State] = @State,
    Zip = @Zip,
    EIN = @EIN,
    Notes = @ProfileNotes,
    UpdatedBy = @UpdatedBy,
    UpdatedDate = GETDATE()
WHERE LawFirmId = @LawFirmId;", conn);

            cmd.Parameters.AddWithValue("@LawFirmId", vm.LawFirmId);
            cmd.Parameters.AddWithValue("@FirmName", (object?)vm.FirmName ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@ShortName", (object?)vm.ShortName ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@Status", (object?)vm.Status ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@MainPhone", (object?)vm.MainPhone ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@MainEmail", (object?)vm.MainEmail ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@Website", (object?)vm.Website ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@Address1", (object?)vm.Address1 ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@Address2", (object?)vm.Address2 ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@City", (object?)vm.City ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@State", (object?)vm.State ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@Zip", (object?)vm.Zip ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@EIN", (object?)NormalizeEin(vm.EIN) ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@ProfileNotes", (object?)vm.ProfileNotes ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@UpdatedBy", User?.Identity?.Name ?? "System");

            cmd.ExecuteNonQuery();



            AddAudit(vm.LawFirmId, "LawFirm", vm.LawFirmId, "Update", $"Updated {Labels.Full.ToLower()} {vm.FirmName}");

            TempData["Message"] = $"{Labels.Full} updated.";
            return RedirectToAction(nameof(Edit), new { id = vm.LawFirmId });
        }

        [HttpGet]
        public IActionResult Details(int id)
        {
            var otherType = RedirectIfOtherType(id, nameof(Details));
            if (otherType != null)
                return otherType;

            var vm = GetLawFirmEditVm(id);
            if (vm == null)
                return NotFound();

            return View(vm);
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public IActionResult DeleteFirm(int lawFirmId)
        {
            using var conn = new SqlConnection(_connectionString);
            conn.Open();

            using (var cmd = new SqlCommand("DELETE FROM dbo.LawFirmNote WHERE LawFirmId = @LawFirmId;", conn))
            {
                cmd.Parameters.AddWithValue("@LawFirmId", lawFirmId);
                cmd.ExecuteNonQuery();
            }

            using (var cmd = new SqlCommand("DELETE FROM dbo.LawFirmContact WHERE LawFirmId = @LawFirmId;", conn))
            {
                cmd.Parameters.AddWithValue("@LawFirmId", lawFirmId);
                cmd.ExecuteNonQuery();
            }

            using (var cmd = new SqlCommand("DELETE FROM dbo.LawFirmStateMap WHERE LawFirmId = @LawFirmId;", conn))
            {
                cmd.Parameters.AddWithValue("@LawFirmId", lawFirmId);
                cmd.ExecuteNonQuery();
            }

            using (var cmd = new SqlCommand("DELETE FROM dbo.LawFirmAudit WHERE LawFirmId = @LawFirmId;", conn))
            {
                cmd.Parameters.AddWithValue("@LawFirmId", lawFirmId);
                cmd.ExecuteNonQuery();
            }

            using (var cmd = new SqlCommand("DELETE FROM dbo.LawFirm WHERE LawFirmId = @LawFirmId;", conn))
            {
                cmd.Parameters.AddWithValue("@LawFirmId", lawFirmId);
                cmd.ExecuteNonQuery();
            }

            TempData["Message"] = $"{Labels.Full} deleted.";
            return RedirectToAction(nameof(Index));
        }

        // =========================================================
        // PRIVATE HELPERS
        // =========================================================
        private List<LawFirmListRowVm> GetLawFirms()
        {
            var rows = new List<LawFirmListRowVm>();

            using var conn = new SqlConnection(_connectionString);
            using var cmd = new SqlCommand(@"
WITH InvoiceLast AS
(
    SELECT 
        b.LawFirmId,
        MAX(a.TranDate) AS LastInvoice
    FROM Accounting_Data.dbo.FirmInvoices a
    INNER JOIN dbo.LawFirmAccountingMap b
        ON b.AccountingFirmName = a.FIRM
       AND b.IsActive = 1
    GROUP BY b.LawFirmId
),
RemitLast AS
(
    SELECT 
        b.LawFirmId,
        MAX(a.TranDate) AS LastRemit
    FROM Accounting_Data.dbo.FirmRemits a
    INNER JOIN dbo.LawFirmAccountingMap b
        ON b.AccountingFirmName = a.FIRM
       AND b.IsActive = 1
    GROUP BY b.LawFirmId
),
CostLast AS
(
    SELECT 
        b.LawFirmId,
        MAX(a.TranDate) AS LastCostFile
    FROM Accounting_Data.dbo.FirmCost a
    INNER JOIN dbo.LawFirmAccountingMap b
        ON b.AccountingFirmName = a.FIRM
       AND b.IsActive = 1
    GROUP BY b.LawFirmId
),
StateCounts AS
(
    SELECT LawFirmId, COUNT(*) AS StateCount
    FROM dbo.LawFirmStateMap
    GROUP BY LawFirmId
),
ContactCounts AS
(
    SELECT LawFirmId, COUNT(*) AS ContactCount
    FROM dbo.LawFirmContact
    GROUP BY LawFirmId
)
SELECT
    lf.LawFirmId,
    lf.FirmName,
    lf.ShortName,
    lf.Status,
    ISNULL(lf.UpdatedDate, lf.CreatedDate) AS UpdatedDate,
    ISNULL(sc.StateCount, 0) AS StateCount,
    ISNULL(cc.ContactCount, 0) AS ContactCount,
    il.LastInvoice,
    rl.LastRemit,
    cl.LastCostFile
FROM dbo.LawFirm lf
LEFT JOIN StateCounts sc ON sc.LawFirmId = lf.LawFirmId
LEFT JOIN ContactCounts cc ON cc.LawFirmId = lf.LawFirmId
LEFT JOIN InvoiceLast il ON il.LawFirmId = lf.LawFirmId
LEFT JOIN RemitLast rl ON rl.LawFirmId = lf.LawFirmId
LEFT JOIN CostLast cl ON cl.LawFirmId = lf.LawFirmId
WHERE lf.EntityType = @RecordType
ORDER BY lf.FirmName;", conn);

            AddRecordTypeParam(cmd);

            conn.Open();

            using var rdr = cmd.ExecuteReader();

            while (rdr.Read())
            {
                rows.Add(new LawFirmListRowVm
                {
                    LawFirmId = Convert.ToInt32(rdr["LawFirmId"]),
                    FirmName = rdr["FirmName"]?.ToString(),
                    ShortName = rdr["ShortName"]?.ToString(),
                    Status = rdr["Status"]?.ToString(),
                    StateCount = Convert.ToInt32(rdr["StateCount"]),
                    ContactCount = Convert.ToInt32(rdr["ContactCount"]),
                    UpdatedDate = rdr["UpdatedDate"] == DBNull.Value ? null : Convert.ToDateTime(rdr["UpdatedDate"]),
                    LastInvoice = rdr["LastInvoice"] == DBNull.Value ? null : Convert.ToDateTime(rdr["LastInvoice"]),
                    LastRemit = rdr["LastRemit"] == DBNull.Value ? null : Convert.ToDateTime(rdr["LastRemit"]),
                    LastCostFile = rdr["LastCostFile"] == DBNull.Value ? null : Convert.ToDateTime(rdr["LastCostFile"])
                });
            }

            return rows;
        }

        private LawFirmEditVm? GetLawFirmEditVm(int id)
        {
            LawFirmEditVm? vm = null;

            using var conn = new SqlConnection(_connectionString);
            using var cmd = new SqlCommand(@"
SELECT *
FROM dbo.LawFirm
WHERE LawFirmId = @LawFirmId;", conn);

            cmd.Parameters.AddWithValue("@LawFirmId", id);

            conn.Open();
            using (var rdr = cmd.ExecuteReader())
            {
                if (rdr.Read())
                {
                    vm = new LawFirmEditVm
                    {
                        LawFirmId = Convert.ToInt32(rdr["LawFirmId"]),
                        FirmName = rdr["FirmName"]?.ToString(),
                        ShortName = rdr["ShortName"]?.ToString(),
                        Status = rdr["Status"]?.ToString(),
                        MainPhone = rdr["MainPhone"]?.ToString(),
                        MainEmail = rdr["MainEmail"]?.ToString(),
                        Website = rdr["Website"]?.ToString(),
                        Address1 = rdr["Address1"]?.ToString(),
                        Address2 = rdr["Address2"]?.ToString(),
                        City = rdr["City"]?.ToString(),
                        State = rdr["State"]?.ToString(),
                        Zip = rdr["Zip"]?.ToString(),
                        EIN = rdr["EIN"]?.ToString(),
                        ProfileNotes = rdr["Notes"]?.ToString()
                    };
                }
            }

            if (vm == null)
                return null;

            vm.StateMappings = GetStateMappings(id);
            vm.Contacts = GetContacts(id);
            vm.Notes = GetNotes(id);
            vm.AuditHistory = GetAuditHistory(id);
            vm.StateCodeOptions = StateCodeOptions;
            vm.ContactTypeOptions = ContactTypeOptions;
            vm.Documents = GetDocuments(id);
            vm.AccountingMaps = GetAccountingMaps(id);
            vm.AccountingActivity = GetAccountingActivity(id);

            // =========================
            // LOAD ACCOUNTING FIRMS
            // =========================
            // Sep 08 - was every Firm name ever seen in the Accounting transaction tables; now that
            // app's own canonical list, ported from FirmLookupService.GetActiveFirmNamesAsync with only
            // the DB prefixes swapped (it runs on Accounting_Data reaching into BR_App, we're the
            // reverse, same server). CommandTimeout carried over too.
            //
            // Ours, not the source's: lf.Status = 'Active' on the JOIN so an Inactive firm falls back to
            // the raw COCO_NAME rather than vanishing (hence the CASE on lf.LawFirmId); the third branch
            // re-admitting mapped aliases, so an existing row's stored value still matches an option
            // (the Aug 21 bug on the Client side); and the NULL/blank guard the source applies in C#.
            vm.AccountingFirmOptions = new List<string>();

            using (var acctCmd = new SqlCommand(@"
SELECT DISTINCT FirmName FROM (
    SELECT
        CASE
            WHEN lf.LawFirmId IS NOT NULL
                THEN lf.FirmName
            ELSE af.COCO_NAME
        END AS FirmName
    FROM Accounting_Data.dbo.AccountingFirms af
    LEFT JOIN dbo.LawFirmAccountingMap lam
        ON lam.AccountingFirmName = af.COCO_NAME
        AND lam.IsActive = 1
    LEFT JOIN dbo.LawFirm lf
        ON lf.LawFirmId = lam.LawFirmId
        AND lf.Status = 'Active'

    UNION

    SELECT lf.FirmName
    FROM dbo.LawFirm lf
    WHERE lf.Status = 'Active'
      AND NOT EXISTS
    (
        SELECT 1
        FROM dbo.LawFirmAccountingMap lam
        WHERE lam.LawFirmId = lf.LawFirmId
          AND lam.IsActive = 1
    )

    UNION

    SELECT lam2.AccountingFirmName AS FirmName
    FROM dbo.LawFirmAccountingMap lam2
    WHERE lam2.IsActive = 1
) b
WHERE FirmName IS NOT NULL AND LTRIM(RTRIM(FirmName)) <> ''
ORDER BY FirmName;", conn) { CommandTimeout = 500 })
            {
                using var acctRdr = acctCmd.ExecuteReader();
                while (acctRdr.Read())
                {
                    vm.AccountingFirmOptions.Add(acctRdr["FirmName"].ToString());
                }
            }



            return vm;
        }

        private List<LawFirmAuditVm> GetAuditHistory(int lawFirmId)
        {
            var list = new List<LawFirmAuditVm>();

            using var conn = new SqlConnection(_connectionString);
            using var cmd = new SqlCommand(@"
SELECT *
FROM dbo.LawFirmAudit
WHERE LawFirmId = @LawFirmId
ORDER BY ChangedDate DESC, LawFirmAuditId DESC;", conn);

            cmd.Parameters.AddWithValue("@LawFirmId", lawFirmId);

            conn.Open();
            using var rdr = cmd.ExecuteReader();
            while (rdr.Read())
            {
                list.Add(new LawFirmAuditVm
                {
                    LawFirmAuditId = Convert.ToInt32(rdr["LawFirmAuditId"]),
                    LawFirmId = Convert.ToInt32(rdr["LawFirmId"]),
                    EntityType = rdr["EntityType"]?.ToString(),
                    EntityId = rdr["EntityId"] == DBNull.Value ? null : (int?)Convert.ToInt32(rdr["EntityId"]),
                    ActionType = rdr["ActionType"]?.ToString(),
                    ChangedBy = rdr["ChangedBy"]?.ToString(),
                    ChangedDate = rdr["ChangedDate"] == DBNull.Value
                        ? DateTime.MinValue
                        : Convert.ToDateTime(rdr["ChangedDate"]),
                    ChangeSummary = rdr["ChangeSummary"]?.ToString()
                });
            }

            return list;
        }

        private void AddAudit(int lawFirmId, string entityType, int? entityId, string actionType, string summary)
        {
            using var conn = new SqlConnection(_connectionString);
            using var cmd = new SqlCommand(@"
INSERT INTO dbo.LawFirmAudit
(
    LawFirmId,
    EntityType,
    EntityId,
    ActionType,
    ChangedBy,
    ChangedDate,
    ChangeSummary
)
VALUES
(
    @LawFirmId,
    @EntityType,
    @EntityId,
    @ActionType,
    @ChangedBy,
    GETDATE(),
    @ChangeSummary
);", conn);

            cmd.Parameters.AddWithValue("@LawFirmId", lawFirmId);
            cmd.Parameters.AddWithValue("@EntityType", entityType);
            cmd.Parameters.AddWithValue("@EntityId", (object?)entityId ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@ActionType", actionType);
            cmd.Parameters.AddWithValue("@ChangedBy", User?.Identity?.Name ?? "System");
            cmd.Parameters.AddWithValue("@ChangeSummary", summary);

            conn.Open();
            cmd.ExecuteNonQuery();
        }

        /// <summary>EINs get pasted with stray spaces as often as typed clean - trim before the
        /// write, and save blank as NULL rather than "" so an empty EIN has one clear state.
        /// Mirrors ClientController's version.</summary>
        private static string? NormalizeEin(string? ein) =>
            string.IsNullOrWhiteSpace(ein) ? null : ein.Trim();

        /// <summary>Whether an EIN value is already on another law firm's row - two vendor
        /// records sharing one taxpayer number is exactly what a 1099 can't survive, so Create/Edit
        /// block on it. Runs on the caller's already-open connection so each POST pays for one
        /// connection, not two. The column is compared as-is: NormalizeEin trims the incoming
        /// value and every write path stores it through NormalizeEin, so stored EINs carry no
        /// padding. Blank EINs never conflict, since a missing EIN is common and not a duplicate.
        /// excludeLawFirmId lets Edit check every OTHER row, not itself.
        /// Mirrors ClientController's version.</summary>
        private bool EinAlreadyInUse(SqlConnection conn, string? ein, int excludeLawFirmId = 0)
        {
            var normalized = NormalizeEin(ein);
            if (normalized == null)
                return false;

            using var cmd = new SqlCommand(@"
SELECT COUNT(*)
FROM dbo.LawFirm
WHERE EIN = @Ein
  AND LawFirmId <> @ExcludeLawFirmId;", conn);

            // Typed as VarChar to match the varchar(20) column - an NVarChar parameter would
            // make SQL Server convert the column and discard any index seek on EIN.
            cmd.Parameters.Add("@Ein", SqlDbType.VarChar, 20).Value = normalized;
            cmd.Parameters.AddWithValue("@ExcludeLawFirmId", excludeLawFirmId);

            return Convert.ToInt32(cmd.ExecuteScalar()) > 0;
        }

        private static List<LawFirmListRowVm> SortRows(List<LawFirmListRowVm> rows, string sort, string dir)
        {
            bool desc = string.Equals(dir, "DESC", StringComparison.OrdinalIgnoreCase);

            return (sort ?? "").ToUpperInvariant() switch
            {
                "STATUS" => desc
                    ? rows.OrderByDescending(x => x.Status).ThenBy(x => x.FirmName).ToList()
                    : rows.OrderBy(x => x.Status).ThenBy(x => x.FirmName).ToList(),

                "STATE_COUNT" => desc
                    ? rows.OrderByDescending(x => x.StateCount).ThenBy(x => x.FirmName).ToList()
                    : rows.OrderBy(x => x.StateCount).ThenBy(x => x.FirmName).ToList(),

                "CONTACT_COUNT" => desc
                    ? rows.OrderByDescending(x => x.ContactCount).ThenBy(x => x.FirmName).ToList()
                    : rows.OrderBy(x => x.ContactCount).ThenBy(x => x.FirmName).ToList(),

                "UPDATED_DATE" => desc
                    ? rows.OrderByDescending(x => x.UpdatedDate).ThenBy(x => x.FirmName).ToList()
                    : rows.OrderBy(x => x.UpdatedDate).ThenBy(x => x.FirmName).ToList(),

                _ => desc
                    ? rows.OrderByDescending(x => x.FirmName).ToList()
                    : rows.OrderBy(x => x.FirmName).ToList()
            };
        }

        /// <summary>Sep 08 - the current firm's own name, for the audit text on both sides of a
        /// reassignment.</summary>
        private string? GetFirmNameById(int lawFirmId)
        {
            using var conn = new SqlConnection(_connectionString);
            using var cmd = new SqlCommand("SELECT FirmName FROM dbo.LawFirm WHERE LawFirmId = @LawFirmId;", conn);
            cmd.Parameters.AddWithValue("@LawFirmId", lawFirmId);
            conn.Open();
            return cmd.ExecuteScalar() as string;
        }



    }



}
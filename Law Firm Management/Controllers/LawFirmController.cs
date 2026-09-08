using ClosedXML.Excel;
using DocumentFormat.OpenXml.InkML;
using Law_Firm_Management.Models;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using Microsoft.VisualBasic.FileIO;
using System.Data;
using System.IO;
using System.Text.Json;
using static Law_Firm_Management.Models.LawFirmEditVm;
using static System.Runtime.InteropServices.JavaScript.JSType;

namespace Law_Firm_Management.Controllers
{
    public class LawFirmController : Controller
    {
        private readonly string _connectionString;

        /// <summary>Sep 08 - TempData key the reassignment prompt travels under, named once so the
        /// writing actions and the Edit GET that reads it can't drift.</summary>
        private const string AccountingMapConflictKey = "AccountingMapConflict";

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

        private List<LawFirmDocumentVm> GetDocuments(int lawFirmId)
        {
            var list = new List<LawFirmDocumentVm>();

            using var conn = new SqlConnection(_connectionString);
            using var cmd = new SqlCommand(@"
SELECT
    LawFirmDocumentId,
    LawFirmId,
    DocumentType,
    DocumentName,
    OriginalFileName,
    ContentType,
    FileSizeBytes,
    Notes,
    EffectiveDate,
    ExpirationDate,
    IsActive,
    UploadedBy,
    UploadedDate
FROM dbo.LawFirmDocument
WHERE LawFirmId = @LawFirmId
ORDER BY UploadedDate DESC, LawFirmDocumentId DESC;", conn);

            cmd.Parameters.AddWithValue("@LawFirmId", lawFirmId);

            conn.Open();
            using var rdr = cmd.ExecuteReader();
            while (rdr.Read())
            {
                list.Add(new LawFirmDocumentVm
                {
                    LawFirmDocumentId = Convert.ToInt32(rdr["LawFirmDocumentId"]),
                    LawFirmId = Convert.ToInt32(rdr["LawFirmId"]),
                    DocumentType = rdr["DocumentType"]?.ToString(),
                    DocumentName = rdr["DocumentName"]?.ToString(),
                    OriginalFileName = rdr["OriginalFileName"]?.ToString(),
                    ContentType = rdr["ContentType"]?.ToString(),
                    FileSizeBytes = rdr["FileSizeBytes"] == DBNull.Value ? null : (long?)Convert.ToInt64(rdr["FileSizeBytes"]),
                    Notes = rdr["Notes"]?.ToString(),
                    EffectiveDate = rdr["EffectiveDate"] == DBNull.Value ? null : (DateTime?)Convert.ToDateTime(rdr["EffectiveDate"]),
                    ExpirationDate = rdr["ExpirationDate"] == DBNull.Value ? null : (DateTime?)Convert.ToDateTime(rdr["ExpirationDate"]),
                    IsActive = rdr["IsActive"] != DBNull.Value && Convert.ToBoolean(rdr["IsActive"]),
                    UploadedBy = rdr["UploadedBy"]?.ToString(),
                    UploadedDate = rdr["UploadedDate"] == DBNull.Value ? DateTime.MinValue : Convert.ToDateTime(rdr["UploadedDate"])
                });
            }

            return list;
        }

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
            var ws = wb.Worksheets.Add("Law Firms");

            ws.Cell(1, 1).Value = "Firm Name";
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
                $"LawFirmList_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");
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
    Notes,
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
    @ProfileNotes,
    @CreatedBy,
    GETDATE(),
    @UpdatedBy,
    GETDATE()
);

SELECT CAST(SCOPE_IDENTITY() AS INT);", conn);

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
            cmd.Parameters.AddWithValue("@ProfileNotes", (object?)vm.ProfileNotes ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@CreatedBy", User?.Identity?.Name ?? "System");
            cmd.Parameters.AddWithValue("@UpdatedBy", User?.Identity?.Name ?? "System");

            conn.Open();
            int newId = (int)cmd.ExecuteScalar();

            AddAudit(newId, "LawFirm", newId, "Insert", $"Created law firm {vm.FirmName}");

            TempData["Message"] = "Law firm created.";
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
            cmd.Parameters.AddWithValue("@ProfileNotes", (object?)vm.ProfileNotes ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@UpdatedBy", User?.Identity?.Name ?? "System");

            conn.Open();
            cmd.ExecuteNonQuery();



            AddAudit(vm.LawFirmId, "LawFirm", vm.LawFirmId, "Update", $"Updated law firm {vm.FirmName}");

            TempData["Message"] = "Law firm updated.";
            return RedirectToAction(nameof(Edit), new { id = vm.LawFirmId });
        }

        [HttpGet]
        public IActionResult Details(int id)
        {
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

            TempData["Message"] = "Law firm deleted.";
            return RedirectToAction(nameof(Index));
        }

        // =========================================================
        // STATE MAPPINGS
        // =========================================================
        [HttpPost]
        [ValidateAntiForgeryToken]
        public IActionResult AddStateMap(LawFirmStateMapVm vm)
        {
            if (string.IsNullOrWhiteSpace(vm.StateCode))
            {
                TempData["Error"] = "State is required.";
                return RedirectToAction(nameof(Edit), new { id = vm.LawFirmId });
            }

            if (string.IsNullOrWhiteSpace(vm.Portfolio))
            {
                TempData["Error"] = "Portfolio is required.";
                return RedirectToAction(nameof(Edit), new { id = vm.LawFirmId });
            }

            if (vm.PlacementPercent.HasValue && (vm.PlacementPercent < 0 || vm.PlacementPercent > 100))
            {
                TempData["Error"] = "Placement % must be between 0 and 100.";
                return RedirectToAction(nameof(Edit), new { id = vm.LawFirmId });
            }

            using var conn = new SqlConnection(_connectionString);
            conn.Open();

            using (var dupCmd = new SqlCommand(@"
SELECT COUNT(*)
FROM dbo.LawFirmStateMap
WHERE LawFirmId = @LawFirmId
  AND StateCode = @StateCode
  AND ISNULL(Portfolio,'') = ISNULL(@Portfolio,'')
  AND LawFirmStateMapId <> @LawFirmStateMapId;", conn))
            {
                dupCmd.Parameters.AddWithValue("@LawFirmId", vm.LawFirmId);
                dupCmd.Parameters.AddWithValue("@StateCode", vm.StateCode);
                dupCmd.Parameters.AddWithValue("@Portfolio", (object?)vm.Portfolio ?? DBNull.Value);
                dupCmd.Parameters.AddWithValue("@LawFirmStateMapId", vm.LawFirmStateMapId);

                int exists = Convert.ToInt32(dupCmd.ExecuteScalar());

                if (exists > 0)
                {
                    TempData["Error"] = "That state and portfolio are already mapped for this firm.";
                    return RedirectToAction(nameof(Edit), new { id = vm.LawFirmId });
                }
            }

            using var cmd = new SqlCommand(@"
INSERT INTO dbo.LawFirmStateMap
(
    LawFirmId,
    StateCode,
    COCO_NO,
    ExportId,
    PlacementPercent,
    Portfolio,
    SftpFolderPath,
    IsActive,
    Notes,
    UpdatedDate
)
VALUES
(
    @LawFirmId,
    @StateCode,
    @COCO_NO,
    @ExportId,
    @PlacementPercent,
    @Portfolio,
    @SftpFolderPath,
    @IsActive,
    @Notes,
    GETDATE()
);", conn);

            cmd.Parameters.AddWithValue("@LawFirmId", vm.LawFirmId);
            cmd.Parameters.AddWithValue("@StateCode", vm.StateCode.Trim().ToUpper());
            cmd.Parameters.AddWithValue("@COCO_NO", (object?)vm.COCO_NO ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@ExportId", (object?)vm.ExportId ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@PlacementPercent", (object?)vm.PlacementPercent ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@Portfolio", vm.Portfolio.Trim());
            cmd.Parameters.AddWithValue("@SftpFolderPath", (object?)vm.SftpFolderPath ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@IsActive", true);
            cmd.Parameters.AddWithValue("@Notes", (object?)vm.Notes ?? DBNull.Value);

            cmd.ExecuteNonQuery();

            AddAudit(vm.LawFirmId, "StateMap", null, "Insert", $"Added state mapping {vm.StateCode} - {vm.Portfolio}");

            TempData["Message"] = "State mapping added.";
            return RedirectToAction(nameof(Edit), new { id = vm.LawFirmId });
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public IActionResult UpdateStateMap(LawFirmStateMapVm vm)
        {
            if (string.IsNullOrWhiteSpace(vm.StateCode))
            {
                TempData["Error"] = "State is required.";
                return RedirectToAction(nameof(Edit), new { id = vm.LawFirmId });
            }

            if (vm.PlacementPercent.HasValue && (vm.PlacementPercent < 0 || vm.PlacementPercent > 100))
            {
                TempData["Error"] = "Placement % must be between 0 and 100.";
                return RedirectToAction(nameof(Edit), new { id = vm.LawFirmId });
            }

            using var conn = new SqlConnection(_connectionString);
            conn.Open();

            using (var dupCmd = new SqlCommand(@"
SELECT COUNT(*)
FROM dbo.LawFirmStateMap
WHERE LawFirmId = @LawFirmId
  AND StateCode = @StateCode
  AND ISNULL(Portfolio, '') = ISNULL(@Portfolio, '')
  AND LawFirmStateMapId <> @LawFirmStateMapId;", conn))
            {
                dupCmd.Parameters.AddWithValue("@LawFirmId", vm.LawFirmId);
                dupCmd.Parameters.AddWithValue("@StateCode", vm.StateCode.Trim().ToUpper());
                dupCmd.Parameters.AddWithValue("@Portfolio", (object?)vm.Portfolio?.Trim() ?? DBNull.Value);
                dupCmd.Parameters.AddWithValue("@LawFirmStateMapId", vm.LawFirmStateMapId);

                int exists = Convert.ToInt32(dupCmd.ExecuteScalar());

                if (exists > 0)
                {
                    TempData["Error"] = "That state and portfolio are already mapped for this firm.";
                    return RedirectToAction(nameof(Edit), new { id = vm.LawFirmId });
                }
            }

            using var cmd = new SqlCommand(@"
UPDATE dbo.LawFirmStateMap
SET
    StateCode = @StateCode,
    COCO_NO = @COCO_NO,
    ExportId = @ExportId,
    PlacementPercent = @PlacementPercent,
    Portfolio = @Portfolio,
    SftpFolderPath = @SftpFolderPath,
    IsActive = @IsActive,
    Notes = @Notes,
    UpdatedDate = GETDATE()
WHERE LawFirmStateMapId = @LawFirmStateMapId;", conn);

            cmd.Parameters.AddWithValue("@LawFirmStateMapId", vm.LawFirmStateMapId);
            cmd.Parameters.AddWithValue("@StateCode", (object?)vm.StateCode ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@COCO_NO", (object?)vm.COCO_NO ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@ExportId", (object?)vm.ExportId ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@PlacementPercent", (object?)vm.PlacementPercent ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@Portfolio", (object?)vm.Portfolio ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@SftpFolderPath", (object?)vm.SftpFolderPath ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@IsActive", vm.IsActive);
            cmd.Parameters.AddWithValue("@Notes", (object?)vm.Notes ?? DBNull.Value);

            cmd.ExecuteNonQuery();

            AddAudit(vm.LawFirmId, "StateMap", vm.LawFirmStateMapId, "Update", $"Updated state mapping {vm.StateCode}");

            TempData["Message"] = "State mapping updated.";
            return RedirectToAction(nameof(Edit), new { id = vm.LawFirmId });
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public IActionResult DeleteStateMap(int lawFirmStateMapId, int lawFirmId)
        {
            using var conn = new SqlConnection(_connectionString);
            using var cmd = new SqlCommand(@"
DELETE FROM dbo.LawFirmStateMap
WHERE LawFirmStateMapId = @LawFirmStateMapId;", conn);

            cmd.Parameters.AddWithValue("@LawFirmStateMapId", lawFirmStateMapId);

            conn.Open();
            cmd.ExecuteNonQuery();

            AddAudit(lawFirmId, "StateMap", lawFirmStateMapId, "Delete", "Deleted state mapping");

            TempData["Message"] = "State mapping deleted.";
            return RedirectToAction(nameof(Edit), new { id = lawFirmId });
        }

        // =========================================================
        // CONTACTS
        // =========================================================
        [HttpPost]
        [ValidateAntiForgeryToken]
        public IActionResult AddContact(LawFirmContactVm vm)
        {
            using var conn = new SqlConnection(_connectionString);
            using var cmd = new SqlCommand(@"
INSERT INTO dbo.LawFirmContact
(
    LawFirmId,
    ContactType,
    ContactName,
    Title,
    Email,
    Phone,
    PreferredContactMethod,
    IsActive,
    Notes,
    UpdatedDate
)
VALUES
(
    @LawFirmId,
    @ContactType,
    @ContactName,
    @Title,
    @Email,
    @Phone,
    @PreferredContactMethod,
    @IsActive,
    @Notes,
    GETDATE()
);", conn);

            cmd.Parameters.AddWithValue("@LawFirmId", vm.LawFirmId);
            cmd.Parameters.AddWithValue("@ContactType", (object?)vm.ContactType ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@ContactName", (object?)vm.ContactName ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@Title", (object?)vm.Title ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@Email", (object?)vm.Email ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@Phone", (object?)vm.Phone ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@PreferredContactMethod", (object?)vm.PreferredContactMethod ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@IsActive", vm.IsActive);
            cmd.Parameters.AddWithValue("@Notes", (object?)vm.Notes ?? DBNull.Value);

            conn.Open();
            cmd.ExecuteNonQuery();

            AddAudit(vm.LawFirmId, "Contact", null, "Insert", $"Added contact {vm.ContactName}");

            TempData["Message"] = "Contact added.";
            return RedirectToAction(nameof(Edit), new { id = vm.LawFirmId });
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public IActionResult UpdateContact(LawFirmContactVm vm)
        {
            using var conn = new SqlConnection(_connectionString);
            using var cmd = new SqlCommand(@"
UPDATE dbo.LawFirmContact
SET
    ContactType = @ContactType,
    ContactName = @ContactName,
    Title = @Title,
    Email = @Email,
    Phone = @Phone,
    PreferredContactMethod = @PreferredContactMethod,
    IsActive = @IsActive,
    Notes = @Notes,
    UpdatedDate = GETDATE()
WHERE LawFirmContactId = @LawFirmContactId;", conn);

            cmd.Parameters.AddWithValue("@LawFirmContactId", vm.LawFirmContactId);
            cmd.Parameters.AddWithValue("@ContactType", (object?)vm.ContactType ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@ContactName", (object?)vm.ContactName ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@Title", (object?)vm.Title ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@Email", (object?)vm.Email ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@Phone", (object?)vm.Phone ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@PreferredContactMethod", (object?)vm.PreferredContactMethod ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@IsActive", vm.IsActive);
            cmd.Parameters.AddWithValue("@Notes", (object?)vm.Notes ?? DBNull.Value);

            conn.Open();
            cmd.ExecuteNonQuery();

            AddAudit(vm.LawFirmId, "Contact", vm.LawFirmContactId, "Update", $"Updated contact {vm.ContactName}");

            TempData["Message"] = "Contact updated.";
            return RedirectToAction(nameof(Edit), new { id = vm.LawFirmId });
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public IActionResult DeleteContact(int lawFirmContactId, int lawFirmId)
        {
            using var conn = new SqlConnection(_connectionString);
            using var cmd = new SqlCommand(@"
DELETE FROM dbo.LawFirmContact
WHERE LawFirmContactId = @LawFirmContactId;", conn);

            cmd.Parameters.AddWithValue("@LawFirmContactId", lawFirmContactId);

            conn.Open();
            cmd.ExecuteNonQuery();

            AddAudit(lawFirmId, "Contact", lawFirmContactId, "Delete", "Deleted contact");

            TempData["Message"] = "Contact deleted.";
            return RedirectToAction(nameof(Edit), new { id = lawFirmId });
        }

        // =========================================================
        // NOTES
        // =========================================================
        [HttpPost]
        [ValidateAntiForgeryToken]
        public IActionResult AddNote(LawFirmNoteVm vm, string returnAction = "Edit")
        {
            if (vm.LawFirmId <= 0)
            {
                TempData["Error"] = "Invalid law firm.";
                return RedirectToAction(returnAction, new { id = vm.LawFirmId });
            }

            if (string.IsNullOrWhiteSpace(vm.NoteText))
            {
                TempData["Error"] = "Note text is required.";
                return RedirectToAction(returnAction, new { id = vm.LawFirmId });
            }

            using var conn = new SqlConnection(_connectionString);
            using var cmd = new SqlCommand(@"
INSERT INTO dbo.LawFirmNote
(
    LawFirmId,
    NoteType,
    NoteText,
    EnteredBy,
    EnteredDate
)
VALUES
(
    @LawFirmId,
    @NoteType,
    @NoteText,
    @EnteredBy,
    GETDATE()
);", conn);

            cmd.Parameters.AddWithValue("@LawFirmId", vm.LawFirmId);
            cmd.Parameters.AddWithValue("@NoteType", (object?)vm.NoteType ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@NoteText", vm.NoteText.Trim());
            cmd.Parameters.AddWithValue("@EnteredBy", User?.Identity?.Name ?? "System");

            conn.Open();
            cmd.ExecuteNonQuery();

            AddAudit(vm.LawFirmId, "Note", null, "Insert", $"Added note {vm.NoteType}");

            TempData["Message"] = "Note added.";
            return RedirectToAction(returnAction, new { id = vm.LawFirmId });
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public IActionResult UpdateNote(LawFirmNoteVm vm)
        {
            using var conn = new SqlConnection(_connectionString);
            using var cmd = new SqlCommand(@"
UPDATE dbo.LawFirmNote
SET
    NoteType = @NoteType,
    NoteText = @NoteText
WHERE LawFirmNoteId = @LawFirmNoteId;", conn);

            cmd.Parameters.AddWithValue("@LawFirmNoteId", vm.LawFirmNoteId);
            cmd.Parameters.AddWithValue("@NoteType", (object?)vm.NoteType ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@NoteText", (object?)vm.NoteText ?? DBNull.Value);

            conn.Open();
            cmd.ExecuteNonQuery();

            AddAudit(vm.LawFirmId, "Note", vm.LawFirmNoteId, "Update", $"Updated note {vm.NoteType}");

            TempData["Message"] = "Note updated.";
            return RedirectToAction(nameof(Edit), new { id = vm.LawFirmId });
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public IActionResult DeleteNote(int lawFirmNoteId, int lawFirmId)
        {
            using var conn = new SqlConnection(_connectionString);
            using var cmd = new SqlCommand(@"
DELETE FROM dbo.LawFirmNote
WHERE LawFirmNoteId = @LawFirmNoteId;", conn);

            cmd.Parameters.AddWithValue("@LawFirmNoteId", lawFirmNoteId);

            conn.Open();
            cmd.ExecuteNonQuery();

            AddAudit(lawFirmId, "Note", lawFirmNoteId, "Delete", "Deleted note");

            TempData["Message"] = "Note deleted.";
            return RedirectToAction(nameof(Edit), new { id = lawFirmId });
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
ORDER BY lf.FirmName;", conn);

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

        [HttpPost]
        [ValidateAntiForgeryToken]
        public IActionResult AddAccountingMap(int lawFirmId, string accountingFirmName, bool isActive = true, bool confirmReassign = false)
        {
            if (lawFirmId <= 0 || string.IsNullOrWhiteSpace(accountingFirmName))
            {
                TempData["Error"] = "Please select an accounting firm.";
                return RedirectToAction(nameof(Edit), new { id = lawFirmId });
            }

            if (isActive)
            {
                // Sep 08 - re-runs on the CONFIRMED post too: the confirmation went out to the browser
                // and back, so the rows it described are a claim about the past. These are the real ones.
                var conflicts = FindConflictingLawFirmMapRows(accountingFirmName);
                if (conflicts.Count > 0)
                {
                    if (!confirmReassign)
                    {
                        return PromptAccountingMapReassign(
                            nameof(AddAccountingMap), lawFirmId, 0, accountingFirmName, isActive, conflicts);
                    }

                    return ReassignAccountingMap(lawFirmId, accountingFirmName, isActive, 0, conflicts,
                        successMessage: "Accounting mapping added.");
                }
            }

            using var conn = new SqlConnection(_connectionString);
            using var cmd = new SqlCommand(@"
INSERT INTO dbo.LawFirmAccountingMap
(
    LawFirmId,
    AccountingFirmName,
    IsActive,
    CreatedDate
)
VALUES
(
    @LawFirmId,
    @AccountingFirmName,
    @IsActive,
    GETDATE()
);", conn);

            cmd.Parameters.AddWithValue("@LawFirmId", lawFirmId);
            cmd.Parameters.AddWithValue("@AccountingFirmName", accountingFirmName);
            cmd.Parameters.AddWithValue("@IsActive", isActive);

            conn.Open();
            cmd.ExecuteNonQuery();

            TempData["Message"] = "Accounting mapping added.";
            return RedirectToAction(nameof(Edit), new { id = lawFirmId });
        }

        private List<LawFirmStateMapVm> GetStateMappings(int lawFirmId)
        {
            var list = new List<LawFirmStateMapVm>();

            using var conn = new SqlConnection(_connectionString);
            using var cmd = new SqlCommand(@"
SELECT *
FROM dbo.LawFirmStateMap
WHERE LawFirmId = @LawFirmId
ORDER BY StateCode;", conn);

            cmd.Parameters.AddWithValue("@LawFirmId", lawFirmId);

            conn.Open();
            using var rdr = cmd.ExecuteReader();
            while (rdr.Read())
            {
                list.Add(new LawFirmStateMapVm
                {
                    LawFirmStateMapId = Convert.ToInt32(rdr["LawFirmStateMapId"]),
                    LawFirmId = Convert.ToInt32(rdr["LawFirmId"]),
                    StateCode = rdr["StateCode"]?.ToString(),
                    COCO_NO = rdr["COCO_NO"]?.ToString(),
                    ExportId = rdr["ExportId"]?.ToString(),
                    PlacementPercent = rdr["PlacementPercent"] == DBNull.Value ? null : (decimal?)Convert.ToDecimal(rdr["PlacementPercent"]),
                    Portfolio = rdr["Portfolio"]?.ToString(),
                    SftpFolderPath = rdr["SftpFolderPath"]?.ToString(),
                    IsActive = rdr["IsActive"] != DBNull.Value && Convert.ToBoolean(rdr["IsActive"]),
                    Notes = rdr["Notes"]?.ToString()
                });
            }

            return list;
        }

        private List<LawFirmContactVm> GetContacts(int lawFirmId)
        {
            var list = new List<LawFirmContactVm>();

            using var conn = new SqlConnection(_connectionString);
            using var cmd = new SqlCommand(@"
SELECT *
FROM dbo.LawFirmContact
WHERE LawFirmId = @LawFirmId
ORDER BY ContactType, ContactName;", conn);

            cmd.Parameters.AddWithValue("@LawFirmId", lawFirmId);

            conn.Open();
            using var rdr = cmd.ExecuteReader();
            while (rdr.Read())
            {
                list.Add(new LawFirmContactVm
                {
                    LawFirmContactId = Convert.ToInt32(rdr["LawFirmContactId"]),
                    LawFirmId = Convert.ToInt32(rdr["LawFirmId"]),
                    ContactType = rdr["ContactType"]?.ToString(),
                    ContactName = rdr["ContactName"]?.ToString(),
                    Title = rdr["Title"]?.ToString(),
                    Email = rdr["Email"]?.ToString(),
                    Phone = rdr["Phone"]?.ToString(),
                    PreferredContactMethod = rdr["PreferredContactMethod"]?.ToString(),
                    IsActive = rdr["IsActive"] != DBNull.Value && Convert.ToBoolean(rdr["IsActive"]),
                    Notes = rdr["Notes"]?.ToString()
                });
            }

            return list;
        }

        private List<LawFirmNoteVm> GetNotes(int lawFirmId)
        {
            var list = new List<LawFirmNoteVm>();

            using var conn = new SqlConnection(_connectionString);
            using var cmd = new SqlCommand(@"
SELECT *
FROM dbo.LawFirmNote
WHERE LawFirmId = @LawFirmId
ORDER BY EnteredDate DESC, LawFirmNoteId DESC;", conn);

            cmd.Parameters.AddWithValue("@LawFirmId", lawFirmId);

            conn.Open();
            using var rdr = cmd.ExecuteReader();
            while (rdr.Read())
            {
                list.Add(new LawFirmNoteVm
                {
                    LawFirmNoteId = Convert.ToInt32(rdr["LawFirmNoteId"]),
                    LawFirmId = Convert.ToInt32(rdr["LawFirmId"]),
                    NoteType = rdr["NoteType"]?.ToString(),
                    NoteText = rdr["NoteText"]?.ToString(),
                    EnteredBy = rdr["EnteredBy"]?.ToString(),
                    EnteredDate = rdr["EnteredDate"] == DBNull.Value
                        ? DateTime.MinValue
                        : Convert.ToDateTime(rdr["EnteredDate"])
                });
            }

            return list;
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

        [HttpPost]
        [ValidateAntiForgeryToken]
        public IActionResult AddDocument(LawFirmDocumentVm vm, string returnAction = "Edit")
        {
            if (vm.LawFirmId <= 0)
            {
                TempData["Error"] = "Invalid law firm.";
                return RedirectToAction(returnAction, new { id = vm.LawFirmId });
            }

            if (string.IsNullOrWhiteSpace(vm.DocumentType) || string.IsNullOrWhiteSpace(vm.DocumentName))
            {
                TempData["Error"] = "Document type and document name are required.";
                return RedirectToAction(returnAction, new { id = vm.LawFirmId });
            }

            if (vm.UploadFile == null || vm.UploadFile.Length == 0)
            {
                TempData["Error"] = "Please choose a file to upload.";
                return RedirectToAction(returnAction, new { id = vm.LawFirmId });
            }

            byte[] fileBytes;
            using (var ms = new MemoryStream())
            {
                vm.UploadFile.CopyTo(ms);
                fileBytes = ms.ToArray();
            }

            using var conn = new SqlConnection(_connectionString);
            using var cmd = new SqlCommand(@"
INSERT INTO dbo.LawFirmDocument
(
    LawFirmId,
    DocumentType,
    DocumentName,
    OriginalFileName,
    ContentType,
    FileData,
    FileSizeBytes,
    Notes,
    EffectiveDate,
    ExpirationDate,
    IsActive,
    UploadedBy,
    UploadedDate
)
VALUES
(
    @LawFirmId,
    @DocumentType,
    @DocumentName,
    @OriginalFileName,
    @ContentType,
    @FileData,
    @FileSizeBytes,
    @Notes,
    @EffectiveDate,
    @ExpirationDate,
    @IsActive,
    @UploadedBy,
    GETDATE()
);", conn);

            cmd.Parameters.AddWithValue("@LawFirmId", vm.LawFirmId);
            cmd.Parameters.AddWithValue("@DocumentType", vm.DocumentType);
            cmd.Parameters.AddWithValue("@DocumentName", vm.DocumentName);
            cmd.Parameters.AddWithValue("@OriginalFileName", vm.UploadFile.FileName);
            cmd.Parameters.AddWithValue("@ContentType", (object?)vm.UploadFile.ContentType ?? DBNull.Value);
            cmd.Parameters.Add("@FileData", SqlDbType.VarBinary, -1).Value = fileBytes;
            cmd.Parameters.AddWithValue("@FileSizeBytes", vm.UploadFile.Length);
            cmd.Parameters.AddWithValue("@Notes", (object?)vm.Notes ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@EffectiveDate", (object?)vm.EffectiveDate ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@ExpirationDate", (object?)vm.ExpirationDate ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@IsActive", vm.IsActive);
            cmd.Parameters.AddWithValue("@UploadedBy", User?.Identity?.Name ?? "System");

            conn.Open();
            cmd.ExecuteNonQuery();

            AddAudit(vm.LawFirmId, "Document", null, "Insert", $"Added document {vm.DocumentName}");

            TempData["Message"] = "Document uploaded.";
            return RedirectToAction(returnAction, new { id = vm.LawFirmId });
        }

        [HttpGet]
        public IActionResult DownloadDocument(int id)
        {
            using var conn = new SqlConnection(_connectionString);
            using var cmd = new SqlCommand(@"
SELECT
    OriginalFileName,
    ContentType,
    FileData
FROM dbo.LawFirmDocument
WHERE LawFirmDocumentId = @LawFirmDocumentId;", conn);

            cmd.Parameters.AddWithValue("@LawFirmDocumentId", id);

            conn.Open();
            using var rdr = cmd.ExecuteReader();

            if (!rdr.Read())
                return NotFound();

            var fileName = rdr["OriginalFileName"]?.ToString() ?? "document";
            var contentType = rdr["ContentType"]?.ToString();
            var fileData = (byte[])rdr["FileData"];

            return File(
                fileData,
                string.IsNullOrWhiteSpace(contentType) ? "application/octet-stream" : contentType,
                fileName);
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public IActionResult DeleteDocument(int lawFirmDocumentId, int lawFirmId, string returnAction = "Edit")
        {
            using var conn = new SqlConnection(_connectionString);
            conn.Open();

            string? documentName = null;

            using (var getCmd = new SqlCommand(@"
SELECT DocumentName
FROM dbo.LawFirmDocument
WHERE LawFirmDocumentId = @LawFirmDocumentId;", conn))
            {
                getCmd.Parameters.AddWithValue("@LawFirmDocumentId", lawFirmDocumentId);
                var result = getCmd.ExecuteScalar();
                documentName = result?.ToString();
            }

            using (var cmd = new SqlCommand(@"
DELETE FROM dbo.LawFirmDocument
WHERE LawFirmDocumentId = @LawFirmDocumentId;", conn))
            {
                cmd.Parameters.AddWithValue("@LawFirmDocumentId", lawFirmDocumentId);
                cmd.ExecuteNonQuery();
            }

            AddAudit(lawFirmId, "Document", lawFirmDocumentId, "Delete", $"Deleted document {documentName}");

            TempData["Message"] = "Document deleted.";
            return RedirectToAction(returnAction, new { id = lawFirmId });
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public IActionResult ReplaceDocument(LawFirmDocumentVm vm, string returnAction = "Edit")
        {
            if (vm.LawFirmId <= 0 || vm.LawFirmDocumentId <= 0)
            {
                TempData["Error"] = "Invalid document.";
                return RedirectToAction(returnAction, new { id = vm.LawFirmId });
            }

            if (vm.UploadFile == null || vm.UploadFile.Length == 0)
            {
                TempData["Error"] = "Please choose a replacement file.";
                return RedirectToAction(returnAction, new { id = vm.LawFirmId });
            }

            byte[] fileBytes;
            using (var ms = new MemoryStream())
            {
                vm.UploadFile.CopyTo(ms);
                fileBytes = ms.ToArray();
            }

            using var conn = new SqlConnection(_connectionString);
            using var cmd = new SqlCommand(@"
UPDATE dbo.LawFirmDocument
SET
    DocumentType = @DocumentType,
    DocumentName = @DocumentName,
    OriginalFileName = @OriginalFileName,
    ContentType = @ContentType,
    FileData = @FileData,
    FileSizeBytes = @FileSizeBytes,
    Notes = @Notes,
    EffectiveDate = @EffectiveDate,
    ExpirationDate = @ExpirationDate,
    IsActive = @IsActive,
    UploadedBy = @UploadedBy,
    UploadedDate = GETDATE()
WHERE LawFirmDocumentId = @LawFirmDocumentId;", conn);

            cmd.Parameters.AddWithValue("@LawFirmDocumentId", vm.LawFirmDocumentId);
            cmd.Parameters.AddWithValue("@DocumentType", (object?)vm.DocumentType ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@DocumentName", (object?)vm.DocumentName ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@OriginalFileName", vm.UploadFile.FileName);
            cmd.Parameters.AddWithValue("@ContentType", (object?)vm.UploadFile.ContentType ?? DBNull.Value);
            cmd.Parameters.Add("@FileData", SqlDbType.VarBinary, -1).Value = fileBytes;
            cmd.Parameters.AddWithValue("@FileSizeBytes", vm.UploadFile.Length);
            cmd.Parameters.AddWithValue("@Notes", (object?)vm.Notes ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@EffectiveDate", (object?)vm.EffectiveDate ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@ExpirationDate", (object?)vm.ExpirationDate ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@IsActive", vm.IsActive);
            cmd.Parameters.AddWithValue("@UploadedBy", User?.Identity?.Name ?? "System");

            conn.Open();
            cmd.ExecuteNonQuery();

            AddAudit(vm.LawFirmId, "Document", vm.LawFirmDocumentId, "Update", $"Replaced document {vm.DocumentName}");

            TempData["Message"] = "Document replaced.";
            return RedirectToAction(returnAction, new { id = vm.LawFirmId });
        }



        [HttpGet]
        public IActionResult Import()
        {
            return View();
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public IActionResult Import(IFormFile file)
        {
            if (file == null || file.Length == 0)
            {
                TempData["Error"] = "Please upload a CSV file.";
                return View();
            }

            var errors = new List<string>();
            int inserted = 0;
            int rowNumber = 0;

            using var stream = file.OpenReadStream();
            using var parser = new TextFieldParser(stream);

            parser.TextFieldType = FieldType.Delimited;
            parser.SetDelimiters(",");
            parser.HasFieldsEnclosedInQuotes = true;

            // Skip header row
            if (!parser.EndOfData)
            {
                parser.ReadFields();
                rowNumber++;
            }

            using var conn = new SqlConnection(_connectionString);
            conn.Open();

            while (!parser.EndOfData)
            {
                rowNumber++;

                try
                {
                    var fields = parser.ReadFields();

                    if (fields == null || fields.Length < 12)
                    {
                        errors.Add($"Row {rowNumber}: Not enough columns.");
                        continue;
                    }

                    var firmName = fields[0]?.Trim();
                    var shortName = fields[1]?.Trim();
                    var status = string.IsNullOrWhiteSpace(fields[2]) ? "Active" : fields[2].Trim();
                    var mainPhone = fields[3]?.Trim();
                    var mainEmail = fields[4]?.Trim();
                    var website = fields[5]?.Trim();
                    var address1 = fields[6]?.Trim();
                    var address2 = fields[7]?.Trim();
                    var city = fields[8]?.Trim();
                    var state = fields[9]?.Trim();
                    var zip = fields[10]?.Trim();
                    var profileNotes = fields[11]?.Trim();

                    if (string.IsNullOrWhiteSpace(firmName))
                    {
                        errors.Add($"Row {rowNumber}: FirmName is required.");
                        continue;
                    }

                    if (firmName.Length > 200)
                    {
                        errors.Add($"Row {rowNumber}: FirmName exceeds 200 characters.");
                        continue;
                    }

                    if (!string.IsNullOrWhiteSpace(shortName) && shortName.Length > 100)
                    {
                        errors.Add($"Row {rowNumber}: ShortName exceeds 100 characters.");
                        continue;
                    }

                    if (!string.Equals(status, "Active", StringComparison.OrdinalIgnoreCase) &&
                        !string.Equals(status, "Inactive", StringComparison.OrdinalIgnoreCase))
                    {
                        errors.Add($"Row {rowNumber}: Status must be Active or Inactive.");
                        continue;
                    }

                    using (var dupCmd = new SqlCommand(@"
SELECT COUNT(*)
FROM dbo.LawFirm
WHERE FirmName = @FirmName;", conn))
                    {
                        dupCmd.Parameters.AddWithValue("@FirmName", firmName);

                        int exists = Convert.ToInt32(dupCmd.ExecuteScalar());
                        if (exists > 0)
                        {
                            errors.Add($"Row {rowNumber}: Firm '{firmName}' already exists.");
                            continue;
                        }
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
    Notes,
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
    @ProfileNotes,
    @CreatedBy,
    GETDATE(),
    @UpdatedBy,
    GETDATE()
);

SELECT CAST(SCOPE_IDENTITY() AS INT);", conn);

                    cmd.Parameters.AddWithValue("@FirmName", (object?)firmName ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("@ShortName", (object?)shortName ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("@Status", (object?)status ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("@MainPhone", (object?)mainPhone ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("@MainEmail", (object?)mainEmail ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("@Website", (object?)website ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("@Address1", (object?)address1 ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("@Address2", (object?)address2 ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("@City", (object?)city ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("@State", (object?)state ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("@Zip", (object?)zip ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("@ProfileNotes", (object?)profileNotes ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("@CreatedBy", User?.Identity?.Name ?? "Import");
                    cmd.Parameters.AddWithValue("@UpdatedBy", User?.Identity?.Name ?? "Import");

                    int newId = Convert.ToInt32(cmd.ExecuteScalar());

                    AddAudit(newId, "LawFirm", newId, "Insert", $"Imported law firm {firmName}");

                    inserted++;
                }
                catch (Exception ex)
                {
                    errors.Add($"Row {rowNumber}: {ex.Message}");
                }
            }

            TempData["Message"] = $"{inserted} law firm(s) imported.";
            if (errors.Any())
            {
                TempData["Error"] = string.Join("<br/>", errors);
            }

            return RedirectToAction(nameof(Index));
        }

        [HttpGet]
        public IActionResult DownloadLawFirmTemplate()
        {
            var csv = new System.Text.StringBuilder();
            csv.AppendLine("FirmName,ShortName,Status,MainPhone,MainEmail,Website,Address1,Address2,City,State,Zip,ProfileNotes");
            csv.AppendLine("Test Firm,SJ Law,Active,555-111-2222,info@sjlaw.com,https://www.sjlaw.com,123 Main St,,Kansas City,MO,64101,Primary partner firm");

            var bytes = System.Text.Encoding.UTF8.GetBytes(csv.ToString());
            return File(bytes, "text/csv", "LawFirmImportTemplate.csv");
        }

    

    [HttpGet]
        public IActionResult ImportContacts()
        {
            return View();
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public IActionResult ImportContacts(IFormFile file)
        {
            if (file == null || file.Length == 0)
            {
                TempData["Error"] = "Please upload a CSV file.";
                return View();
            }

            var errors = new List<string>();
            int inserted = 0;
            int rowNumber = 0;

            using var stream = file.OpenReadStream();
            using var parser = new TextFieldParser(stream);

            parser.TextFieldType = FieldType.Delimited;
            parser.SetDelimiters(",");
            parser.HasFieldsEnclosedInQuotes = true;

            if (!parser.EndOfData)
            {
                parser.ReadFields(); // header
                rowNumber++;
            }

            using var conn = new SqlConnection(_connectionString);
            conn.Open();

            while (!parser.EndOfData)
            {
                rowNumber++;

                try
                {
                    var fields = parser.ReadFields();

                    if (fields == null || fields.Length < 9)
                    {
                        errors.Add($"Row {rowNumber}: Not enough columns.");
                        continue;
                    }

                    var firmName = fields[0]?.Trim();
                    var contactType = fields[1]?.Trim();
                    var contactName = fields[2]?.Trim();
                    var title = fields[3]?.Trim();
                    var email = fields[4]?.Trim();
                    var phone = fields[5]?.Trim();
                    var preferredContactMethod = fields[6]?.Trim();
                    var isActiveText = fields[7]?.Trim();
                    var notes = fields[8]?.Trim();

                    if (string.IsNullOrWhiteSpace(firmName))
                    {
                        errors.Add($"Row {rowNumber}: FirmName is required.");
                        continue;
                    }

                    if (string.IsNullOrWhiteSpace(contactName))
                    {
                        errors.Add($"Row {rowNumber}: ContactName is required.");
                        continue;
                    }

                    bool isActive = true;
                    if (!string.IsNullOrWhiteSpace(isActiveText) && !bool.TryParse(isActiveText, out isActive))
                    {
                        if (isActiveText == "1")
                            isActive = true;
                        else if (isActiveText == "0")
                            isActive = false;
                        else
                        {
                            errors.Add($"Row {rowNumber}: IsActive must be true, false, 1, or 0.");
                            continue;
                        }
                    }

                    int lawFirmId;
                    using (var firmCmd = new SqlCommand(@"
SELECT LawFirmId
FROM dbo.LawFirm
WHERE FirmName = @FirmName;", conn))
                    {
                        firmCmd.Parameters.AddWithValue("@FirmName", firmName);

                        var firmIdObj = firmCmd.ExecuteScalar();
                        if (firmIdObj == null || firmIdObj == DBNull.Value)
                        {
                            errors.Add($"Row {rowNumber}: Firm '{firmName}' was not found.");
                            continue;
                        }

                        lawFirmId = Convert.ToInt32(firmIdObj);
                    }

                    using (var dupCmd = new SqlCommand(@"
SELECT COUNT(*)
FROM dbo.LawFirmContact
WHERE LawFirmId = @LawFirmId
  AND ISNULL(ContactName, '') = @ContactName
  AND ISNULL(Email, '') = @Email;", conn))
                    {
                        dupCmd.Parameters.AddWithValue("@LawFirmId", lawFirmId);
                        dupCmd.Parameters.AddWithValue("@ContactName", contactName ?? "");
                        dupCmd.Parameters.AddWithValue("@Email", email ?? "");

                        int exists = Convert.ToInt32(dupCmd.ExecuteScalar());
                        if (exists > 0)
                        {
                            errors.Add($"Row {rowNumber}: Contact '{contactName}' already exists for '{firmName}'.");
                            continue;
                        }
                    }

                    using var cmd = new SqlCommand(@"
INSERT INTO dbo.LawFirmContact
(
    LawFirmId,
    ContactType,
    ContactName,
    Title,
    Email,
    Phone,
    PreferredContactMethod,
    IsActive,
    Notes,
    CreatedDate,
    UpdatedDate
)
VALUES
(
    @LawFirmId,
    @ContactType,
    @ContactName,
    @Title,
    @Email,
    @Phone,
    @PreferredContactMethod,
    @IsActive,
    @Notes,
    GETDATE(),
    GETDATE()
);", conn);

                    cmd.Parameters.AddWithValue("@LawFirmId", lawFirmId);
                    cmd.Parameters.AddWithValue("@ContactType", (object?)contactType ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("@ContactName", (object?)contactName ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("@Title", (object?)title ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("@Email", (object?)email ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("@Phone", (object?)phone ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("@PreferredContactMethod", (object?)preferredContactMethod ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("@IsActive", isActive);
                    cmd.Parameters.AddWithValue("@Notes", (object?)notes ?? DBNull.Value);

                    cmd.ExecuteNonQuery();
                    inserted++;
                }
                catch (Exception ex)
                {
                    errors.Add($"Row {rowNumber}: {ex.Message}");
                }
            }

            TempData["Message"] = $"{inserted} contact(s) imported.";
            if (errors.Any())
                TempData["Error"] = string.Join("<br/>", errors);

            return RedirectToAction(nameof(Index));
        }

        [HttpGet]
        public IActionResult DownloadContactTemplate()
        {
            var csv = new System.Text.StringBuilder();
            csv.AppendLine("FirmName,ContactType,ContactName,Title,Email,Phone,PreferredContactMethod,IsActive,Notes");
            csv.AppendLine("Test Firm,IT,John Smith,Manager,john@test.com,555-111-2222,Email,true,Main IT contact");

            var bytes = System.Text.Encoding.UTF8.GetBytes(csv.ToString());
            return File(bytes, "text/csv", "LawFirmContactImportTemplate.csv");
        }

        [HttpGet]
        public IActionResult ImportStateMappings()
        {
            return View();
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public IActionResult ImportStateMappings(IFormFile file)
        {
            if (file == null || file.Length == 0)
            {
                TempData["Error"] = "Please upload a CSV file.";
                return View();
            }

            var errors = new List<string>();
            int inserted = 0;
            int rowNumber = 0;

            using var stream = file.OpenReadStream();
            using var parser = new TextFieldParser(stream);

            parser.TextFieldType = FieldType.Delimited;
            parser.SetDelimiters(",");
            parser.HasFieldsEnclosedInQuotes = true;

            if (!parser.EndOfData)
            {
                parser.ReadFields(); // skip header
                rowNumber++;
            }

            using var conn = new SqlConnection(_connectionString);
            conn.Open();

            while (!parser.EndOfData)
            {
                rowNumber++;

                try
                {
                    var fields = parser.ReadFields();

                    if (fields == null || fields.Length < 9)
                    {
                        errors.Add($"Row {rowNumber}: Not enough columns.");
                        continue;
                    }

                    var firmName = fields[0]?.Trim();
                    var stateCode = fields[1]?.Trim()?.ToUpper();
                    var cocoNo = fields[2]?.Trim();
                    var exportId = fields[3]?.Trim();
                    var placementPercentText = fields[4]?.Trim();
                    var portfolio = fields[5]?.Trim();
                    var sftpFolderPath = fields[6]?.Trim();
                    var isActiveText = fields[7]?.Trim();
                    var notes = fields[8]?.Trim();

                    if (string.IsNullOrWhiteSpace(firmName))
                    {
                        errors.Add($"Row {rowNumber}: FirmName is required.");
                        continue;
                    }

                    if (string.IsNullOrWhiteSpace(stateCode))
                    {
                        errors.Add($"Row {rowNumber}: StateCode is required.");
                        continue;
                    }

                    if (stateCode.Length != 2)
                    {
                        errors.Add($"Row {rowNumber}: StateCode must be 2 characters.");
                        continue;
                    }

                    if (!StateCodeOptions.Contains(stateCode))
                    {
                        errors.Add($"Row {rowNumber}: StateCode '{stateCode}' is not valid.");
                        continue;
                    }

                    decimal? placementPercent = null;
                    if (!string.IsNullOrWhiteSpace(placementPercentText))
                    {
                        if (!decimal.TryParse(placementPercentText, out var parsedPercent))
                        {
                            errors.Add($"Row {rowNumber}: PlacementPercent must be numeric.");
                            continue;
                        }

                        if (parsedPercent < 0 || parsedPercent > 100)
                        {
                            errors.Add($"Row {rowNumber}: PlacementPercent must be between 0 and 100.");
                            continue;
                        }

                        placementPercent = parsedPercent;
                    }

                    bool isActive = true;
                    if (!string.IsNullOrWhiteSpace(isActiveText) && !bool.TryParse(isActiveText, out isActive))
                    {
                        if (isActiveText == "1")
                            isActive = true;
                        else if (isActiveText == "0")
                            isActive = false;
                        else
                        {
                            errors.Add($"Row {rowNumber}: IsActive must be true, false, 1, or 0.");
                            continue;
                        }
                    }

                    if (!string.IsNullOrWhiteSpace(portfolio) && !new[] { "CCMR", "NCB", "MTB", "Contingency","Other","Collections" }.Contains(portfolio))
                    {
                        errors.Add($"Row {rowNumber}: Portfolio '{portfolio}' is not valid.");
                        continue;
                    }

                    int lawFirmId;
                    using (var firmCmd = new SqlCommand(@"
SELECT LawFirmId
FROM dbo.LawFirm
WHERE FirmName = @FirmName;", conn))
                    {
                        firmCmd.Parameters.AddWithValue("@FirmName", firmName);

                        var firmIdObj = firmCmd.ExecuteScalar();
                        if (firmIdObj == null || firmIdObj == DBNull.Value)
                        {
                            errors.Add($"Row {rowNumber}: Firm '{firmName}' was not found.");
                            continue;
                        }

                        lawFirmId = Convert.ToInt32(firmIdObj);
                    }

                    using (var dupCmd = new SqlCommand(@"
SELECT COUNT(*)
FROM dbo.LawFirmStateMap
WHERE LawFirmId = @LawFirmId
  AND StateCode = @StateCode
  AND ISNULL(Portfolio, '') = ISNULL(@Portfolio, '');", conn))
                    {
                        dupCmd.Parameters.AddWithValue("@LawFirmId", lawFirmId);
                        dupCmd.Parameters.AddWithValue("@StateCode", stateCode);
                        dupCmd.Parameters.AddWithValue("@Portfolio", (object?)portfolio ?? DBNull.Value);

                        int exists = Convert.ToInt32(dupCmd.ExecuteScalar());
                        if (exists > 0)
                        {
                            errors.Add($"Row {rowNumber}: State '{stateCode}' already exists for '{firmName}'.");
                            continue;
                        }
                    }

                    using var cmd = new SqlCommand(@"
INSERT INTO dbo.LawFirmStateMap
(
    LawFirmId,
    StateCode,
    COCO_NO,
    ExportId,
    PlacementPercent,
    SftpFolderPath,
    IsActive,
    Notes,
    CreatedDate,
    UpdatedDate,
    Portfolio
)
VALUES
(
    @LawFirmId,
    @StateCode,
    @COCO_NO,
    @ExportId,
    @PlacementPercent,
    @SftpFolderPath,
    @IsActive,
    @Notes,
    GETDATE(),
    GETDATE(),
    @Portfolio
);", conn);

                    cmd.Parameters.AddWithValue("@LawFirmId", lawFirmId);
                    cmd.Parameters.AddWithValue("@StateCode", stateCode);
                    cmd.Parameters.AddWithValue("@COCO_NO", (object?)cocoNo ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("@ExportId", (object?)exportId ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("@PlacementPercent", (object?)placementPercent ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("@SftpFolderPath", (object?)sftpFolderPath ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("@IsActive", isActive);
                    cmd.Parameters.AddWithValue("@Notes", (object?)notes ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("@Portfolio", (object?)portfolio ?? DBNull.Value);

                    cmd.ExecuteNonQuery();
                    inserted++;
                }
                catch (Exception ex)
                {
                    errors.Add($"Row {rowNumber}: {ex.Message}");
                }
            }

            TempData["Message"] = $"{inserted} state mapping(s) imported.";
            if (errors.Any())
                TempData["Error"] = string.Join("<br/>", errors);

            return RedirectToAction(nameof(Index));
        }

        [HttpGet]
        public IActionResult DownloadStateMappingTemplate()
        {
            var csv = new System.Text.StringBuilder();
            csv.AppendLine("FirmName,StateCode,COCO_NO,ExportId,PlacementPercent,Portfolio,SftpFolderPath,IsActive,Notes");
            csv.AppendLine("Test Firm,LA,23234,rere,100.00,Contingency,LS9873,0,Sample Louisiana mapping");

            var bytes = System.Text.Encoding.UTF8.GetBytes(csv.ToString());
            return File(bytes, "text/csv", "LawFirmStateMappingImportTemplate.csv");
        }

        [HttpGet]
        public IActionResult PortfolioReport(string portfolio = "")
        {
            var vm = new LawFirmPortfolioReportVm
            {
                SelectedPortfolio = portfolio,
                PortfolioOptions = new List<string>(),
                Rows = new List<LawFirmPortfolioReportRowVm>()
            };

            using var conn = new SqlConnection(_connectionString);
            conn.Open();

            // =========================
            // LOAD PORTFOLIO DROPDOWN
            // =========================
            using (var cmd = new SqlCommand(@"
        SELECT DISTINCT Portfolio
        FROM dbo.LawFirmStateMap
        WHERE Portfolio IS NOT NULL
        ORDER BY Portfolio;", conn))
            {
                using var rdr = cmd.ExecuteReader();
                while (rdr.Read())
                {
                    vm.PortfolioOptions.Add(rdr["Portfolio"].ToString());
                }
            }

            // =========================
            // LOAD REPORT DATA
            // =========================
            using (var cmd = new SqlCommand(@"
        SELECT 
            b.StateCode,
            a.FirmName,
            b.Portfolio,
            b.COCO_NO,
            b.ExportId,
            b.SftpFolderPath,
            b.PlacementPercent
        FROM dbo.LawFirm a
        INNER JOIN dbo.LawFirmStateMap b 
            ON b.LawFirmId = a.LawFirmId
        WHERE (@Portfolio = '' OR b.Portfolio = @Portfolio) and b.IsActive = 1
        ORDER BY b.StateCode, a.FirmName;", conn))
            {
                cmd.Parameters.AddWithValue("@Portfolio", portfolio ?? "");

                using var rdr = cmd.ExecuteReader();

                while (rdr.Read())
                {
                    vm.Rows.Add(new LawFirmPortfolioReportRowVm
                    {
                        StateCode = rdr["StateCode"]?.ToString(),
                        FirmName = rdr["FirmName"]?.ToString(),
                        Portfolio = rdr["Portfolio"]?.ToString(),
                        COCO_NO = rdr["COCO_NO"]?.ToString(),
                        ExportId = rdr["ExportId"]?.ToString(),
                        SftpFolderPath = rdr["SftpFolderPath"]?.ToString(),
                        PlacementPercent = rdr["PlacementPercent"] == DBNull.Value
                            ? null
                            : Convert.ToDecimal(rdr["PlacementPercent"])
                    });
                }
            }

            return View(vm);
        }

        [HttpGet]
        public IActionResult ExportPortfolioReport(string portfolio = "")
        {
            var rows = new List<LawFirmPortfolioReportRowVm>();

            using var conn = new SqlConnection(_connectionString);
            using var cmd = new SqlCommand(@"
        SELECT 
            b.StateCode,
            a.FirmName,
            b.Portfolio,
            b.COCO_NO,
            b.ExportId,
            b.SftpFolderPath,
            b.PlacementPercent
        FROM dbo.LawFirm a
        INNER JOIN dbo.LawFirmStateMap b 
            ON b.LawFirmId = a.LawFirmId
        WHERE (@Portfolio = '' OR b.Portfolio = @Portfolio) and b.IsActive = 1
        ORDER BY b.StateCode, a.FirmName;", conn);

            cmd.Parameters.AddWithValue("@Portfolio", portfolio ?? "");

            conn.Open();

            using var rdr = cmd.ExecuteReader();

            while (rdr.Read())
            {
                rows.Add(new LawFirmPortfolioReportRowVm
                {
                    StateCode = rdr["StateCode"]?.ToString(),
                    FirmName = rdr["FirmName"]?.ToString(),
                    Portfolio = rdr["Portfolio"]?.ToString(),
                    COCO_NO = rdr["COCO_NO"]?.ToString(),
                    ExportId = rdr["ExportId"]?.ToString(),
                    SftpFolderPath = rdr["SftpFolderPath"]?.ToString(),
                    PlacementPercent = rdr["PlacementPercent"] == DBNull.Value
                        ? null
                        : Convert.ToDecimal(rdr["PlacementPercent"])
                });
            }

            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Portfolio Report");

            ws.Cell(1, 1).Value = "State";
            ws.Cell(1, 2).Value = "Firm Name";
            ws.Cell(1, 3).Value = "Portfolio";
            ws.Cell(1, 4).Value = "COCO_NO";
            ws.Cell(1, 5).Value = "Export ID";
            ws.Cell(1, 6).Value = "SFTP Folder";
            ws.Cell(1, 7).Value = "Placement %";

            int row = 2;

            foreach (var item in rows)
            {
                ws.Cell(row, 1).Value = item.StateCode;
                ws.Cell(row, 2).Value = item.FirmName;
                ws.Cell(row, 3).Value = item.Portfolio;
                ws.Cell(row, 4).Value = item.COCO_NO;
                ws.Cell(row, 5).Value = item.ExportId;
                ws.Cell(row, 6).Value = item.SftpFolderPath;
                ws.Cell(row, 7).Value = item.PlacementPercent;
                row++;
            }

            ws.Row(1).Style.Font.Bold = true;
            ws.Columns().AdjustToContents();
            ws.RangeUsed().SetAutoFilter();

            using var stream = new MemoryStream();
            wb.SaveAs(stream);

            var safePortfolio = string.IsNullOrWhiteSpace(portfolio) ? "All" : portfolio;

            return File(
                stream.ToArray(),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                $"PortfolioReport_{safePortfolio}_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");
        }

        private List<LawFirmAccountingMapVm> GetAccountingMaps(int lawFirmId)
        {
            var list = new List<LawFirmAccountingMapVm>();

            using var conn = new SqlConnection(_connectionString);
            using var cmd = new SqlCommand(@"
        SELECT 
            LawFirmAccountingMapId,
            AccountingFirmName,
            IsActive
        FROM dbo.LawFirmAccountingMap
        WHERE LawFirmId = @LawFirmId
        ORDER BY AccountingFirmName;", conn);

            cmd.Parameters.AddWithValue("@LawFirmId", lawFirmId);

            conn.Open();

            using var rdr = cmd.ExecuteReader();

            while (rdr.Read())
            {
                list.Add(new LawFirmAccountingMapVm
                {
                    LawFirmAccountingMapId = Convert.ToInt32(rdr["LawFirmAccountingMapId"]),
                    AccountingFirmName = rdr["AccountingFirmName"]?.ToString(),
                    IsActive = Convert.ToBoolean(rdr["IsActive"])
                });
            }

            return list;
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public IActionResult DeleteAccountingMap(int lawFirmAccountingMapId, int lawFirmId)
        {
            if (lawFirmAccountingMapId <= 0 || lawFirmId <= 0)
            {
                TempData["Error"] = "Invalid accounting mapping.";
                return RedirectToAction(nameof(Edit), new { id = lawFirmId });
            }

            using var conn = new SqlConnection(_connectionString);
            using var cmd = new SqlCommand(@"
DELETE FROM dbo.LawFirmAccountingMap
WHERE LawFirmAccountingMapId = @LawFirmAccountingMapId;", conn);

            cmd.Parameters.AddWithValue("@LawFirmAccountingMapId", lawFirmAccountingMapId);

            conn.Open();
            cmd.ExecuteNonQuery();

            TempData["Message"] = "Accounting mapping deleted.";
            return RedirectToAction(nameof(Edit), new { id = lawFirmId });
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public IActionResult UpdateAccountingMap(int lawFirmAccountingMapId, int lawFirmId, string accountingFirmName, bool confirmReassign = false)
        {
            bool isActive = Request.Form["isActive"].Any(x => x == "true");

            if (lawFirmAccountingMapId <= 0 || lawFirmId <= 0 || string.IsNullOrWhiteSpace(accountingFirmName))
            {
                TempData["Error"] = "Invalid accounting mapping.";
                return RedirectToAction(nameof(Edit), new { id = lawFirmId });
            }

            if (isActive)
            {
                var conflicts = FindConflictingLawFirmMapRows(accountingFirmName, excludeMapId: lawFirmAccountingMapId);
                if (conflicts.Count > 0)
                {
                    if (!confirmReassign)
                    {
                        return PromptAccountingMapReassign(
                            nameof(UpdateAccountingMap), lawFirmId, lawFirmAccountingMapId, accountingFirmName, isActive, conflicts);
                    }

                    return ReassignAccountingMap(lawFirmId, accountingFirmName, isActive, lawFirmAccountingMapId, conflicts,
                        successMessage: "Accounting mapping updated.");
                }
            }

            using var conn = new SqlConnection(_connectionString);
            using var cmd = new SqlCommand(@"
UPDATE dbo.LawFirmAccountingMap
SET
    AccountingFirmName = @AccountingFirmName,
    IsActive = @IsActive
WHERE LawFirmAccountingMapId = @LawFirmAccountingMapId;", conn);

            cmd.Parameters.AddWithValue("@LawFirmAccountingMapId", lawFirmAccountingMapId);
            cmd.Parameters.AddWithValue("@AccountingFirmName", accountingFirmName);
            cmd.Parameters.AddWithValue("@IsActive", isActive);

            conn.Open();
            cmd.ExecuteNonQuery();

            TempData["Message"] = "Accounting mapping updated.";
            return RedirectToAction(nameof(Edit), new { id = lawFirmId });
        }

        /// <summary>
        /// Sep 08 - whether accountingFirmName is already actively mapped to a DIFFERENT firm. New
        /// here: this side had no conflict check at all, so such a mapping was invisible from any one
        /// firm's Edit page. excludeMapId lets Update check every OTHER row, not itself.
        ///
        /// Every match, not TOP 1 - this table predates any duplicate guard, so two firms genuinely
        /// can both claim a name, and clearing only the first would leave it ambiguous.
        /// </summary>
        private List<ConflictingAccountingMapRow> FindConflictingLawFirmMapRows(string accountingFirmName, int excludeMapId = 0)
        {
            var rows = new List<ConflictingAccountingMapRow>();

            using var conn = new SqlConnection(_connectionString);
            using var cmd = new SqlCommand(@"
        SELECT
            m.LawFirmAccountingMapId,
            m.LawFirmId,
            lf.FirmName,
            lf.Status
        FROM dbo.LawFirmAccountingMap m
        JOIN dbo.LawFirm lf ON lf.LawFirmId = m.LawFirmId
        WHERE m.AccountingFirmName = @AccountingFirmName
          AND m.IsActive = 1
          AND m.LawFirmAccountingMapId <> @ExcludeMapId
        ORDER BY m.LawFirmAccountingMapId;", conn);

            cmd.Parameters.AddWithValue("@AccountingFirmName", accountingFirmName);
            cmd.Parameters.AddWithValue("@ExcludeMapId", excludeMapId);

            conn.Open();

            using var rdr = cmd.ExecuteReader();

            while (rdr.Read())
            {
                rows.Add(new ConflictingAccountingMapRow
                {
                    LawFirmAccountingMapId = Convert.ToInt32(rdr["LawFirmAccountingMapId"]),
                    LawFirmId = Convert.ToInt32(rdr["LawFirmId"]),
                    FirmName = rdr["FirmName"]?.ToString(),
                    Status = rdr["Status"]?.ToString()
                });
            }

            return rows;
        }

        /// <summary>
        /// Sep 08 - the one place a conflicting mapping's owner is named: warning, TempData messages,
        /// audit text. Inactive firms are labelled, since "already associated with X" otherwise reads
        /// as a live relationship; never filtered out though. Mirrors ClientController's version.
        /// </summary>
        private static string DescribeConflictOwners(List<ConflictingAccountingMapRow> rows)
        {
            return string.Join(", ", rows
                .Select(r => string.Equals(r.Status, "Active", StringComparison.OrdinalIgnoreCase)
                    ? r.FirmName
                    : $"{r.FirmName} (Inactive)")
                .Distinct());
        }

        /// <summary>Sep 08 - a LawFirmAccountingMap row on some OTHER firm already actively claiming
        /// the name being mapped. Controller-private; nothing outside this flow needs it.</summary>
        private sealed class ConflictingAccountingMapRow
        {
            public int LawFirmAccountingMapId { get; set; }
            public int LawFirmId { get; set; }
            public string? FirmName { get; set; }

            /// <summary>The owning firm's status - labelled in the warning, never filtered on.</summary>
            public string? Status { get; set; }
        }

        /// <summary>
        /// Sep 08 - first half of the reassignment: stash the attempt and redirect to Edit, where the
        /// modal re-offers it with confirmReassign set. Writes nothing, so cancelling leaves both
        /// firms untouched - the reason this is a round trip and not a flag on the original write.
        /// </summary>
        private IActionResult PromptAccountingMapReassign(
            string postAction,
            int lawFirmId,
            int lawFirmAccountingMapId,
            string accountingFirmName,
            bool isActive,
            List<ConflictingAccountingMapRow> conflicts)
        {
            var pending = new LawFirmAccountingMapConflictVm
            {
                LawFirmId = lawFirmId,
                LawFirmAccountingMapId = lawFirmAccountingMapId,
                AccountingFirmName = accountingFirmName,
                IsActive = isActive,
                PostAction = postAction,
                ExistingFirmName = DescribeConflictOwners(conflicts)
            };

            TempData[AccountingMapConflictKey] = JsonSerializer.Serialize(pending);
            return RedirectToAction(nameof(Edit), new { id = lawFirmId });
        }

        /// <summary>
        /// Sep 08 - the confirmed half: delete the old holder's row, then write this firm's, in ONE
        /// transaction - old-first so the states never overlap, transactional so a failure between
        /// them can't strand the mapping deleted-but-never-recreated. Hard DELETE rather than
        /// IsActive = 0, matching the Client side. Audits run after the commit (own connection).
        /// </summary>
        private IActionResult ReassignAccountingMap(
            int lawFirmId,
            string accountingFirmName,
            bool isActive,
            int lawFirmAccountingMapId,
            List<ConflictingAccountingMapRow> conflicts,
            string successMessage)
        {
            using var conn = new SqlConnection(_connectionString);
            conn.Open();

            using var tx = conn.BeginTransaction();

            try
            {
                foreach (var conflict in conflicts)
                {
                    using var deleteCmd = new SqlCommand(@"
DELETE FROM dbo.LawFirmAccountingMap
WHERE LawFirmAccountingMapId = @LawFirmAccountingMapId;", conn, tx);

                    deleteCmd.Parameters.AddWithValue("@LawFirmAccountingMapId", conflict.LawFirmAccountingMapId);
                    deleteCmd.ExecuteNonQuery();
                }

                // lawFirmAccountingMapId > 0 means the user was editing an existing row of this
                // firm's own, so the mapping moves onto that row rather than adding a second one.
                using var writeCmd = lawFirmAccountingMapId > 0
                    ? new SqlCommand(@"
UPDATE dbo.LawFirmAccountingMap
SET
    AccountingFirmName = @AccountingFirmName,
    IsActive = @IsActive
WHERE LawFirmAccountingMapId = @LawFirmAccountingMapId;", conn, tx)
                    : new SqlCommand(@"
INSERT INTO dbo.LawFirmAccountingMap
(
    LawFirmId,
    AccountingFirmName,
    IsActive,
    CreatedDate
)
VALUES
(
    @LawFirmId,
    @AccountingFirmName,
    @IsActive,
    GETDATE()
);", conn, tx);

                writeCmd.Parameters.AddWithValue("@AccountingFirmName", accountingFirmName);
                writeCmd.Parameters.AddWithValue("@IsActive", isActive);

                if (lawFirmAccountingMapId > 0)
                    writeCmd.Parameters.AddWithValue("@LawFirmAccountingMapId", lawFirmAccountingMapId);
                else
                    writeCmd.Parameters.AddWithValue("@LawFirmId", lawFirmId);

                writeCmd.ExecuteNonQuery();

                tx.Commit();
            }
            catch (Exception ex)
            {
                tx.Rollback();

                TempData["Error"] = $"Could not reassign \"{accountingFirmName}\": {ex.Message} " +
                    "No mappings were changed.";
                return RedirectToAction(nameof(Edit), new { id = lawFirmId });
            }

            var previousOwners = DescribeConflictOwners(conflicts);
            var currentFirmName = GetFirmNameById(lawFirmId) ?? $"law firm {lawFirmId}";

            foreach (var conflict in conflicts)
            {
                AddAudit(conflict.LawFirmId, "AccountingMap", conflict.LawFirmAccountingMapId, "Delete",
                    $"Accounting mapping \"{accountingFirmName}\" reassigned to {currentFirmName}");
            }

            AddAudit(lawFirmId, "AccountingMap", lawFirmAccountingMapId > 0 ? (int?)lawFirmAccountingMapId : null, "Update",
                $"Accounting mapping \"{accountingFirmName}\" reassigned from {previousOwners}");

            TempData["Message"] = $"{successMessage} \"{accountingFirmName}\" was reassigned from {previousOwners}.";
            return RedirectToAction(nameof(Edit), new { id = lawFirmId });
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



    private LawFirmAccountingActivityVm GetAccountingActivity(int lawFirmId)
        {
            var vm = new LawFirmAccountingActivityVm();

            using var conn = new SqlConnection(_connectionString);
            conn.Open();

            using (var cmd = new SqlCommand(@"
        SELECT MAX(a.TranDate)
        FROM Accounting_Data.dbo.FirmInvoices a
        INNER JOIN BR_App.dbo.LawFirmAccountingMap b 
            ON b.AccountingFirmName = a.FIRM
        WHERE b.LawFirmId = @LawFirmId
          AND b.IsActive = 1;", conn))
            {
                cmd.Parameters.AddWithValue("@LawFirmId", lawFirmId);
                var result = cmd.ExecuteScalar();
                vm.LastInvoice = result == DBNull.Value ? null : (DateTime?)Convert.ToDateTime(result);
            }

            using (var cmd = new SqlCommand(@"
        SELECT MAX(a.TranDate)
        FROM Accounting_Data.dbo.FirmCost a
        INNER JOIN BR_App.dbo.LawFirmAccountingMap b 
            ON b.AccountingFirmName = a.FIRM
        WHERE b.LawFirmId = @LawFirmId
          AND b.IsActive = 1;", conn))
            {
                cmd.Parameters.AddWithValue("@LawFirmId", lawFirmId);
                var result = cmd.ExecuteScalar();
                vm.LastCostFile = result == DBNull.Value ? null : (DateTime?)Convert.ToDateTime(result);
            }

            using (var cmd = new SqlCommand(@"
        SELECT MAX(a.TranDate)
        FROM Accounting_Data.dbo.FirmRemits a
        INNER JOIN BR_App.dbo.LawFirmAccountingMap b 
            ON b.AccountingFirmName = a.FIRM
        WHERE b.LawFirmId = @LawFirmId
          AND b.IsActive = 1;", conn))
            {
                cmd.Parameters.AddWithValue("@LawFirmId", lawFirmId);
                var result = cmd.ExecuteScalar();
                vm.LastRemit = result == DBNull.Value ? null : (DateTime?)Convert.ToDateTime(result);
            }

            return vm;
        }



    }



}
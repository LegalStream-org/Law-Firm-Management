using ClosedXML.Excel;
using Law_Firm_Management.Models;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.Data.SqlClient;
using Microsoft.VisualBasic.FileIO;
using System.Data;
using System.Text;
using System.Text.Json;
using Microsoft.AspNetCore.Mvc.Rendering;

namespace Law_Firm_Management.Controllers
{
    public class ClientController : Controller
    {
        private readonly string _connectionString;

        /// <summary>Sep 08 - TempData key the reassignment prompt travels under, named once so the
        /// writing actions and the Edit GET that reads it can't drift.</summary>
        private const string AccountingMapConflictKey = "AccountingMapConflict";

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

        public ClientController(IConfiguration configuration)
        {
            _connectionString = configuration.GetConnectionString("Connection")
                ?? throw new InvalidOperationException("Missing SQL Connection connection string.");
        }

        public IActionResult Index(string search, string status, string clientType, string sort = "CLIENT_NAME", string dir = "ASC")
        {
            var rows = GetClients();

            if (!string.IsNullOrWhiteSpace(search))
            {
                rows = rows
                    .Where(x =>
                        (x.ClientName ?? "").Contains(search, StringComparison.OrdinalIgnoreCase) ||
                        (x.ShortName ?? "").Contains(search, StringComparison.OrdinalIgnoreCase))
                    .ToList();
            }

            if (!string.IsNullOrWhiteSpace(status))
            {
                rows = rows
                    .Where(x => string.Equals(x.Status, status, StringComparison.OrdinalIgnoreCase))
                    .ToList();
            }

            if (!string.IsNullOrWhiteSpace(clientType))
            {
                rows = rows
                    .Where(x => string.Equals(x.ClientTypeName, clientType, StringComparison.OrdinalIgnoreCase))
                    .ToList();
            }

            rows = SortRows(rows, sort, dir);

            ViewBag.Search = search ?? "";
            ViewBag.Status = status ?? "";
            ViewBag.ClientType = clientType ?? "";
            ViewBag.ClientTypeOptions = GetClientTypeSelectOptions();
            ViewBag.Sort = sort;
            ViewBag.Dir = dir;

            return View(rows);
        }

        public IActionResult Export(string search, string status, string clientType, string sort = "CLIENT_NAME", string dir = "ASC")
        {
            var rows = GetClients();

            if (!string.IsNullOrWhiteSpace(search))
            {
                rows = rows
                    .Where(x =>
                        (x.ClientName ?? "").Contains(search, StringComparison.OrdinalIgnoreCase) ||
                        (x.ShortName ?? "").Contains(search, StringComparison.OrdinalIgnoreCase))
                    .ToList();
            }

            if (!string.IsNullOrWhiteSpace(status))
            {
                rows = rows
                    .Where(x => string.Equals(x.Status, status, StringComparison.OrdinalIgnoreCase))
                    .ToList();
            }

            if (!string.IsNullOrWhiteSpace(clientType))
            {
                rows = rows
                    .Where(x => string.Equals(x.ClientTypeName, clientType, StringComparison.OrdinalIgnoreCase))
                    .ToList();
            }

            rows = SortRows(rows, sort, dir);

            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Clients");

            ws.Cell(1, 1).Value = "Client Name";
            ws.Cell(1, 2).Value = "Short Name";
            ws.Cell(1, 3).Value = "Status";
            ws.Cell(1, 4).Value = "Client Type";
            ws.Cell(1, 5).Value = "Contact Count";
            ws.Cell(1, 6).Value = "Updated Date";

            int row = 2;
            foreach (var item in rows)
            {
                ws.Cell(row, 1).Value = item.ClientName;
                ws.Cell(row, 2).Value = item.ShortName;
                ws.Cell(row, 3).Value = item.Status;
                ws.Cell(row, 4).Value = item.ClientTypeName;
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
                $"ClientList_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");
        }

        [HttpGet]
        public IActionResult Create()
        {
            return View(new ClientEditVm
            {
                Status = "Active",
                ContactTypeOptions = ContactTypeOptions,
                ClientTypeOptions = GetClientTypeSelectOptions()
            });
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public IActionResult Create(ClientEditVm vm)
        {
            if (!ModelState.IsValid)
            {
                vm.ContactTypeOptions = ContactTypeOptions;
                vm.ClientTypeOptions = GetClientTypeSelectOptions();
                return View(vm);
            }

            using var conn = new SqlConnection(_connectionString);
            using var cmd = new SqlCommand(@"
INSERT INTO dbo.Client
(
    ClientName,
    ShortName,
    Status,
    ClientTypeId,
    MainPhone,
    MainEmail,
    Website,
    Address1,
    Address2,
    City,
    [State],
    Zip,
    ProfileNotes,
    CreatedBy,
    CreatedDate,
    UpdatedBy,
    UpdatedDate
)
VALUES
(
    @ClientName,
    @ShortName,
    @Status,
    @ClientTypeId,
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

            cmd.Parameters.AddWithValue("@ClientName", (object?)vm.ClientName ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@ShortName", (object?)vm.ShortName ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@Status", (object?)vm.Status ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@ClientTypeId", (object?)vm.ClientTypeId ?? DBNull.Value);
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

            AddAudit(newId, "Client", newId, "Insert", $"Created client {vm.ClientName}");

            // #1/#2 (Aug 21) - auto-create the canonical self-map row so this client's own name is
            // always resolvable through ClientAccountingMap from day one, not just via Client.ClientName
            // directly. See EnsureCanonicalAccountingMap's own comment for the full reasoning.
            EnsureCanonicalAccountingMap(newId, vm.ClientName);

            TempData["Message"] = "Client created.";
            return RedirectToAction(nameof(Edit), new { id = newId });
        }

        [HttpGet]
        public IActionResult Edit(int id)
        {
            var vm = GetClientEditVm(id);
            if (vm == null)
                return NotFound();

            // Sep 08 - set by Add/UpdateAccountingMap when the chosen name is already mapped elsewhere.
            // Read here rather than in GetClientEditVm (shared with Details and the POST re-render) so
            // the prompt only appears on the page that raised it; the ClientId guard blocks stale ones.
            var pendingConflict = TempData[AccountingMapConflictKey] as string;
            if (!string.IsNullOrWhiteSpace(pendingConflict))
            {
                var conflictVm = JsonSerializer.Deserialize<ClientAccountingMapConflictVm>(pendingConflict);
                if (conflictVm != null && conflictVm.ClientId == id)
                    vm.PendingAccountingMapConflict = conflictVm;
            }

            return View(vm);
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public IActionResult Edit(ClientEditVm vm)
        {
            if (!ModelState.IsValid)
            {
                vm.ContactTypeOptions = ContactTypeOptions;
                vm.ClientTypeOptions = GetClientTypeSelectOptions();
                vm.Contacts = GetContacts(vm.ClientId);
                vm.Notes = GetNotes(vm.ClientId);
                vm.AuditHistory = GetAuditHistory(vm.ClientId);
                vm.Documents = GetDocuments(vm.ClientId);
                return View(vm);
            }

            // #3 (Aug 21) - fetched before the UPDATE below runs, so this is a real before/after
            // comparison, not just the posted value on its own.
            var oldClientName = GetClientNameById(vm.ClientId);

            using var conn = new SqlConnection(_connectionString);
            using var cmd = new SqlCommand(@"
UPDATE dbo.Client
SET
    ClientName = @ClientName,
    ShortName = @ShortName,
    Status = @Status,
    ClientTypeId = @ClientTypeId,
    MainPhone = @MainPhone,
    MainEmail = @MainEmail,
    Website = @Website,
    Address1 = @Address1,
    Address2 = @Address2,
    City = @City,
    [State] = @State,
    Zip = @Zip,
    ProfileNotes = @ProfileNotes,
    UpdatedBy = @UpdatedBy,
    UpdatedDate = GETDATE()
WHERE ClientId = @ClientId;", conn);

            cmd.Parameters.AddWithValue("@ClientId", vm.ClientId);
            cmd.Parameters.AddWithValue("@ClientName", (object?)vm.ClientName ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@ShortName", (object?)vm.ShortName ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@Status", (object?)vm.Status ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@ClientTypeId", (object?)vm.ClientTypeId ?? DBNull.Value);
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

            AddAudit(vm.ClientId, "Client", vm.ClientId, "Update", $"Updated client {vm.ClientName}");

            // #3 (Aug 21) - only auto-maps the new name when ClientName actually changed - not on
            // every Edit save regardless. Old name's row is left alone; see
            // EnsureCanonicalAccountingMap's own comment for why.
            if (!string.Equals(oldClientName, vm.ClientName, StringComparison.Ordinal))
            {
                EnsureCanonicalAccountingMap(vm.ClientId, vm.ClientName);
            }

            TempData["Message"] = "Client updated.";
            return RedirectToAction(nameof(Edit), new { id = vm.ClientId });
        }

        [HttpGet]
        public IActionResult Details(int id)
        {
            var vm = GetClientEditVm(id);
            if (vm == null)
                return NotFound();

            return View(vm);
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public IActionResult DeleteClient(int clientId)
        {
            using var conn = new SqlConnection(_connectionString);
            conn.Open();

            using (var cmd = new SqlCommand("DELETE FROM dbo.ClientDocument WHERE ClientId = @ClientId;", conn))
            {
                cmd.Parameters.AddWithValue("@ClientId", clientId);
                cmd.ExecuteNonQuery();
            }

            using (var cmd = new SqlCommand("DELETE FROM dbo.ClientNote WHERE ClientId = @ClientId;", conn))
            {
                cmd.Parameters.AddWithValue("@ClientId", clientId);
                cmd.ExecuteNonQuery();
            }

            using (var cmd = new SqlCommand("DELETE FROM dbo.ClientContact WHERE ClientId = @ClientId;", conn))
            {
                cmd.Parameters.AddWithValue("@ClientId", clientId);
                cmd.ExecuteNonQuery();
            }

            using (var cmd = new SqlCommand("DELETE FROM dbo.ClientAudit WHERE ClientId = @ClientId;", conn))
            {
                cmd.Parameters.AddWithValue("@ClientId", clientId);
                cmd.ExecuteNonQuery();
            }

            using (var cmd = new SqlCommand("DELETE FROM dbo.ClientPortfolio WHERE ClientId = @ClientId;", conn))
            {
                cmd.Parameters.AddWithValue("@ClientId", clientId);
                cmd.ExecuteNonQuery();
            }

            using (var cmd = new SqlCommand("DELETE FROM dbo.ClientTask WHERE ClientId = @ClientId;", conn))
            {
                cmd.Parameters.AddWithValue("@ClientId", clientId);
                cmd.ExecuteNonQuery();
            }

            using (var cmd = new SqlCommand("DELETE FROM dbo.ClientSftpConfig WHERE ClientId = @ClientId;", conn))
            {
                cmd.Parameters.AddWithValue("@ClientId", clientId);
                cmd.ExecuteNonQuery();
            }


            using (var cmd = new SqlCommand("DELETE FROM dbo.Client WHERE ClientId = @ClientId;", conn))
            {
                cmd.Parameters.AddWithValue("@ClientId", clientId);
                cmd.ExecuteNonQuery();
            }

            TempData["Message"] = "Client deleted.";
            return RedirectToAction(nameof(Index));
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public IActionResult AddContact(ClientContactVm vm)
        {
            using var conn = new SqlConnection(_connectionString);
            using var cmd = new SqlCommand(@"
INSERT INTO dbo.ClientContact
(
    ClientId,
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
    @ClientId,
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

            cmd.Parameters.AddWithValue("@ClientId", vm.ClientId);
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

            AddAudit(vm.ClientId, "Contact", null, "Insert", $"Added client contact {vm.ContactName}");

            TempData["Message"] = "Contact added.";
            return RedirectToAction(nameof(Edit), new { id = vm.ClientId });
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public IActionResult UpdateContact(ClientContactVm vm)
        {
            using var conn = new SqlConnection(_connectionString);
            using var cmd = new SqlCommand(@"
UPDATE dbo.ClientContact
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
WHERE ClientContactId = @ClientContactId;", conn);

            cmd.Parameters.AddWithValue("@ClientContactId", vm.ClientContactId);
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

            AddAudit(vm.ClientId, "Contact", vm.ClientContactId, "Update", $"Updated client contact {vm.ContactName}");

            TempData["Message"] = "Contact updated.";
            return RedirectToAction(nameof(Edit), new { id = vm.ClientId });
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public IActionResult DeleteContact(int clientContactId, int clientId)
        {
            using var conn = new SqlConnection(_connectionString);
            using var cmd = new SqlCommand("DELETE FROM dbo.ClientContact WHERE ClientContactId = @ClientContactId;", conn);

            cmd.Parameters.AddWithValue("@ClientContactId", clientContactId);

            conn.Open();
            cmd.ExecuteNonQuery();

            AddAudit(clientId, "Contact", clientContactId, "Delete", "Deleted client contact");

            TempData["Message"] = "Contact deleted.";
            return RedirectToAction(nameof(Edit), new { id = clientId });
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public IActionResult AddNote(ClientNoteVm vm, string returnAction = "Edit")
        {
            if (vm.ClientId <= 0)
            {
                TempData["Error"] = "Invalid client.";
                return RedirectToAction(returnAction, new { id = vm.ClientId });
            }

            if (string.IsNullOrWhiteSpace(vm.NoteText))
            {
                TempData["Error"] = "Note text is required.";
                return RedirectToAction(returnAction, new { id = vm.ClientId });
            }

            using var conn = new SqlConnection(_connectionString);
            using var cmd = new SqlCommand(@"
INSERT INTO dbo.ClientNote
(
    ClientId,
    NoteType,
    NoteText,
    EnteredBy,
    EnteredDate
)
VALUES
(
    @ClientId,
    @NoteType,
    @NoteText,
    @EnteredBy,
    GETDATE()
);", conn);

            cmd.Parameters.AddWithValue("@ClientId", vm.ClientId);
            cmd.Parameters.AddWithValue("@NoteType", (object?)vm.NoteType ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@NoteText", vm.NoteText.Trim());
            cmd.Parameters.AddWithValue("@EnteredBy", User?.Identity?.Name ?? "System");

            conn.Open();
            cmd.ExecuteNonQuery();

            AddAudit(vm.ClientId, "Note", null, "Insert", $"Added client note {vm.NoteType}");

            TempData["Message"] = "Note added.";
            return RedirectToAction(returnAction, new { id = vm.ClientId });
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public IActionResult UpdateNote(ClientNoteVm vm)
        {
            using var conn = new SqlConnection(_connectionString);
            using var cmd = new SqlCommand(@"
UPDATE dbo.ClientNote
SET
    NoteType = @NoteType,
    NoteText = @NoteText
WHERE ClientNoteId = @ClientNoteId;", conn);

            cmd.Parameters.AddWithValue("@ClientNoteId", vm.ClientNoteId);
            cmd.Parameters.AddWithValue("@NoteType", (object?)vm.NoteType ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@NoteText", (object?)vm.NoteText ?? DBNull.Value);

            conn.Open();
            cmd.ExecuteNonQuery();

            AddAudit(vm.ClientId, "Note", vm.ClientNoteId, "Update", $"Updated client note {vm.NoteType}");

            TempData["Message"] = "Note updated.";
            return RedirectToAction(nameof(Edit), new { id = vm.ClientId });
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public IActionResult DeleteNote(int clientNoteId, int clientId)
        {
            using var conn = new SqlConnection(_connectionString);
            using var cmd = new SqlCommand("DELETE FROM dbo.ClientNote WHERE ClientNoteId = @ClientNoteId;", conn);

            cmd.Parameters.AddWithValue("@ClientNoteId", clientNoteId);

            conn.Open();
            cmd.ExecuteNonQuery();

            AddAudit(clientId, "Note", clientNoteId, "Delete", "Deleted client note");

            TempData["Message"] = "Note deleted.";
            return RedirectToAction(nameof(Edit), new { id = clientId });
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

                    var clientName = fields[0]?.Trim();
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

                    if (string.IsNullOrWhiteSpace(clientName))
                    {
                        errors.Add($"Row {rowNumber}: ClientName is required.");
                        continue;
                    }

                    using (var dupCmd = new SqlCommand("SELECT COUNT(*) FROM dbo.Client WHERE ClientName = @ClientName;", conn))
                    {
                        dupCmd.Parameters.AddWithValue("@ClientName", clientName);
                        int exists = Convert.ToInt32(dupCmd.ExecuteScalar());
                        if (exists > 0)
                        {
                            errors.Add($"Row {rowNumber}: Client '{clientName}' already exists.");
                            continue;
                        }
                    }

                    using var cmd = new SqlCommand(@"
INSERT INTO dbo.Client
(
    ClientName, ShortName, Status, MainPhone, MainEmail, Website,
    Address1, Address2, City, [State], Zip, ProfileNotes,
    CreatedBy, CreatedDate, UpdatedBy, UpdatedDate
)
VALUES
(
    @ClientName, @ShortName, @Status, @MainPhone, @MainEmail, @Website,
    @Address1, @Address2, @City, @State, @Zip, @ProfileNotes,
    @CreatedBy, GETDATE(), @UpdatedBy, GETDATE()
);

SELECT CAST(SCOPE_IDENTITY() AS INT);", conn);

                    cmd.Parameters.AddWithValue("@ClientName", (object?)clientName ?? DBNull.Value);
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
                    AddAudit(newId, "Client", newId, "Insert", $"Imported client {clientName}");
                    inserted++;
                }
                catch (Exception ex)
                {
                    errors.Add($"Row {rowNumber}: {ex.Message}");
                }
            }

            TempData["Message"] = $"{inserted} client(s) imported.";
            if (errors.Any())
                TempData["Error"] = string.Join("<br/>", errors);

            return RedirectToAction(nameof(Index));
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
                    if (fields == null || fields.Length < 9)
                    {
                        errors.Add($"Row {rowNumber}: Not enough columns.");
                        continue;
                    }

                    var clientName = fields[0]?.Trim();
                    var contactType = fields[1]?.Trim();
                    var contactName = fields[2]?.Trim();
                    var title = fields[3]?.Trim();
                    var email = fields[4]?.Trim();
                    var phone = fields[5]?.Trim();
                    var preferredContactMethod = fields[6]?.Trim();
                    var isActiveText = fields[7]?.Trim();
                    var notes = fields[8]?.Trim();

                    if (string.IsNullOrWhiteSpace(clientName))
                    {
                        errors.Add($"Row {rowNumber}: ClientName is required.");
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

                    int clientId;
                    using (var clientCmd = new SqlCommand("SELECT ClientId FROM dbo.Client WHERE ClientName = @ClientName;", conn))
                    {
                        clientCmd.Parameters.AddWithValue("@ClientName", clientName);
                        var clientIdObj = clientCmd.ExecuteScalar();
                        if (clientIdObj == null || clientIdObj == DBNull.Value)
                        {
                            errors.Add($"Row {rowNumber}: Client '{clientName}' was not found.");
                            continue;
                        }

                        clientId = Convert.ToInt32(clientIdObj);
                    }

                    using (var dupCmd = new SqlCommand(@"
SELECT COUNT(*)
FROM dbo.ClientContact
WHERE ClientId = @ClientId
  AND ISNULL(ContactName, '') = @ContactName
  AND ISNULL(Email, '') = @Email;", conn))
                    {
                        dupCmd.Parameters.AddWithValue("@ClientId", clientId);
                        dupCmd.Parameters.AddWithValue("@ContactName", contactName ?? "");
                        dupCmd.Parameters.AddWithValue("@Email", email ?? "");

                        int exists = Convert.ToInt32(dupCmd.ExecuteScalar());
                        if (exists > 0)
                        {
                            errors.Add($"Row {rowNumber}: Contact '{contactName}' already exists for '{clientName}'.");
                            continue;
                        }
                    }

                    using var cmd = new SqlCommand(@"
INSERT INTO dbo.ClientContact
(
    ClientId, ContactType, ContactName, Title, Email, Phone,
    PreferredContactMethod, IsActive, Notes, CreatedDate, UpdatedDate
)
VALUES
(
    @ClientId, @ContactType, @ContactName, @Title, @Email, @Phone,
    @PreferredContactMethod, @IsActive, @Notes, GETDATE(), GETDATE()
);", conn);

                    cmd.Parameters.AddWithValue("@ClientId", clientId);
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
        public IActionResult DownloadClientTemplate()
        {
            var csv = new StringBuilder();
            csv.AppendLine("ClientName,ShortName,Status,MainPhone,MainEmail,Website,Address1,Address2,City,State,Zip,ProfileNotes");
            csv.AppendLine("ABC Corp,ABC,Active,555-111-2222,info@abc.com,https://www.abc.com,123 Main St,,Kansas City,MO,64101,Main client");

            return File(Encoding.UTF8.GetBytes(csv.ToString()), "text/csv", "ClientImportTemplate.csv");
        }

        [HttpGet]
        public IActionResult DownloadClientContactTemplate()
        {
            var csv = new StringBuilder();
            csv.AppendLine("ClientName,ContactType,ContactName,Title,Email,Phone,PreferredContactMethod,IsActive,Notes");
            csv.AppendLine("ABC Corp,IT,John Smith,Manager,john@abc.com,555-111-2222,Email,true,Main IT contact");

            return File(Encoding.UTF8.GetBytes(csv.ToString()), "text/csv", "ClientContactImportTemplate.csv");
        }

        [HttpGet]
        public IActionResult TaskCalendarEvents(int clientId, DateTime? start, DateTime? end)
        {
            var rangeStart = (start ?? DateTime.Today.AddMonths(-1)).Date;
            var rangeEnd = (end ?? DateTime.Today.AddMonths(3)).Date;

            var events = new List<object>();

            using var conn = new SqlConnection(_connectionString);
            using var cmd = new SqlCommand(@"
SELECT
    ClientTaskId,
    ClientId,
    TaskType,
    TaskName,
    Frequency,
    DueDay,
    AssignedTo,
    Status,
    DeliveryMethod,
    NextDueDate,
    LastCompletedDate,
    Notes,
    SftpConfigId ClientSftpConfigId
FROM dbo.ClientTask
WHERE ClientId = @ClientId
  AND NextDueDate IS NOT NULL
ORDER BY TaskName;", conn);

            cmd.Parameters.AddWithValue("@ClientId", clientId);

            conn.Open();
            using var rdr = cmd.ExecuteReader();

            while (rdr.Read())
            {
                int clientTaskId = Convert.ToInt32(rdr["ClientTaskId"]);
                string taskType = rdr["TaskType"]?.ToString() ?? "";
                string taskName = rdr["TaskName"]?.ToString() ?? "Task";
                string frequency = rdr["Frequency"]?.ToString() ?? "";
                string dueDay = rdr["DueDay"]?.ToString() ?? "";
                string assignedTo = rdr["AssignedTo"]?.ToString() ?? "";
                string status = rdr["Status"]?.ToString() ?? "";
                string deliveryMethod = rdr["DeliveryMethod"]?.ToString() ?? "";
                string notes = rdr["Notes"]?.ToString() ?? "";
                int? clientSftpConfigId = rdr["ClientSftpConfigId"] == DBNull.Value
                    ? null
                    : Convert.ToInt32(rdr["ClientSftpConfigId"]);

                DateTime anchorDate = Convert.ToDateTime(rdr["NextDueDate"]).Date;

                string title = string.IsNullOrWhiteSpace(taskType)
                    ? taskName
                    : $"{taskType} - {taskName}";

                string color = status.Trim().ToLower() switch
                {
                    "completed" => "#198754",
                    "in progress" => "#0d6efd",
                    "pending" => "#ffc107",
                    "overdue" => "#dc3545",
                    _ => "#6c757d"
                };

                foreach (var occurrence in ExpandTaskOccurrences(anchorDate, frequency, rangeStart, rangeEnd))
                {
                    events.Add(new
                    {
                        id = $"{clientTaskId}_{occurrence:yyyyMMdd}",
                        title = title,
                        start = occurrence.ToString("yyyy-MM-dd"),
                        allDay = true,
                        backgroundColor = color,
                        borderColor = color,
                        extendedProps = new
                        {
                            clientTaskId,
                            taskType,
                            taskName,
                            frequency,
                            dueDay,
                            assignedTo,
                            status,
                            deliveryMethod,
                            notes,
                            clientSftpConfigId
                        }
                    });
                }
            }

            return Json(events);
        }

        private static List<DateTime> ExpandTaskOccurrences(
    DateTime anchorDate,
    string? frequency,
    DateTime rangeStart,
    DateTime rangeEnd)
        {
            var results = new List<DateTime>();
            var freq = (frequency ?? "").Trim().ToLowerInvariant();

            if (anchorDate > rangeEnd)
                return results;

            // Non-recurring / blank frequency
            if (string.IsNullOrWhiteSpace(freq) ||
                freq == "one time" ||
                freq == "once" ||
                freq == "ad hoc" ||
                freq == "adhoc")
            {
                if (anchorDate >= rangeStart && anchorDate <= rangeEnd)
                    results.Add(anchorDate);

                return results;
            }

            if (freq == "daily")
            {
                var current = anchorDate < rangeStart ? rangeStart : anchorDate;
                while (current <= rangeEnd)
                {
                    results.Add(current);
                    current = current.AddDays(1);
                }
                return results;
            }

            if (freq == "weekly")
            {
                var current = anchorDate;
                while (current < rangeStart)
                    current = current.AddDays(7);

                while (current <= rangeEnd)
                {
                    results.Add(current);
                    current = current.AddDays(7);
                }
                return results;
            }

            if (freq == "bi-weekly" || freq == "biweekly" || freq == "every 2 weeks")
            {
                var current = anchorDate;
                while (current < rangeStart)
                    current = current.AddDays(14);

                while (current <= rangeEnd)
                {
                    results.Add(current);
                    current = current.AddDays(14);
                }
                return results;
            }

            if (freq == "monthly")
            {
                var current = anchorDate;

                while (current < rangeStart)
                    current = current.AddMonths(1);

                while (current <= rangeEnd)
                {
                    results.Add(current);
                    current = current.AddMonths(1);
                }
                return results;
            }

            if (freq == "quarterly")
            {
                var current = anchorDate;

                while (current < rangeStart)
                    current = current.AddMonths(3);

                while (current <= rangeEnd)
                {
                    results.Add(current);
                    current = current.AddMonths(3);
                }
                return results;
            }

            if (freq == "annual" || freq == "annually" || freq == "yearly")
            {
                var current = anchorDate;

                while (current < rangeStart)
                    current = current.AddYears(1);

                while (current <= rangeEnd)
                {
                    results.Add(current);
                    current = current.AddYears(1);
                }
                return results;
            }

            // Fallback: treat unknown frequency as one-time
            if (anchorDate >= rangeStart && anchorDate <= rangeEnd)
                results.Add(anchorDate);

            return results;
        }

        private List<ClientListRowVm> GetClients()
        {
            var rows = new List<ClientListRowVm>();

            using var conn = new SqlConnection(_connectionString);
            using var cmd = new SqlCommand(@"
SELECT
    c.ClientId,
    c.ClientName,
    c.ShortName,
    c.Status,
    ct.TypeName AS ClientTypeName,
    ISNULL(c.UpdatedDate, c.CreatedDate) AS UpdatedDate,
    COUNT(DISTINCT cc.ClientContactId) AS ContactCount
FROM dbo.Client c
LEFT JOIN dbo.ClientContact cc ON c.ClientId = cc.ClientId
LEFT JOIN dbo.ClientType ct ON ct.ClientTypeId = c.ClientTypeId
GROUP BY
    c.ClientId,
    c.ClientName,
    c.ShortName,
    c.Status,
    ct.TypeName,
    c.UpdatedDate,
    c.CreatedDate;", conn);

            conn.Open();
            using var rdr = cmd.ExecuteReader();
            while (rdr.Read())
            {
                rows.Add(new ClientListRowVm
                {
                    ClientId = Convert.ToInt32(rdr["ClientId"]),
                    ClientName = rdr["ClientName"]?.ToString(),
                    ShortName = rdr["ShortName"]?.ToString(),
                    Status = rdr["Status"]?.ToString(),
                    ClientTypeName = rdr["ClientTypeName"]?.ToString(),
                    ContactCount = Convert.ToInt32(rdr["ContactCount"]),
                    UpdatedDate = rdr["UpdatedDate"] == DBNull.Value ? null : (DateTime?)Convert.ToDateTime(rdr["UpdatedDate"])
                });
            }

            return rows;
        }

        // Powers the ClientType dropdown on Client Create/Edit. The list itself is owned and
        // managed by SettingsController (Settings > Client Types) - this just reads it.
        private List<SelectListItem> GetClientTypeSelectOptions()
        {
            var options = SettingsController.GetActiveClientTypeOptions(_connectionString)
                .Select(t => new SelectListItem
                {
                    Value = t.ClientTypeId.ToString(),
                    Text = t.TypeName
                })
                .ToList();

            options.Insert(0, new SelectListItem { Value = "", Text = "-- Select Client Type --" });
            return options;
        }

        private ClientEditVm? GetClientEditVm(int id)
        {
            ClientEditVm? vm = null;

            using var conn = new SqlConnection(_connectionString);
            using var cmd = new SqlCommand("SELECT * FROM dbo.Client WHERE ClientId = @ClientId;", conn);
            cmd.Parameters.AddWithValue("@ClientId", id);

            conn.Open();
            using (var rdr = cmd.ExecuteReader())
            {
                if (rdr.Read())
                {
                    vm = new ClientEditVm
                    {
                        ClientId = Convert.ToInt32(rdr["ClientId"]),
                        ClientName = rdr["ClientName"]?.ToString(),
                        ShortName = rdr["ShortName"]?.ToString(),
                        Status = rdr["Status"]?.ToString() ?? "Active",
                        ClientTypeId = rdr["ClientTypeId"] == DBNull.Value ? null : Convert.ToInt32(rdr["ClientTypeId"]),
                        MainPhone = rdr["MainPhone"]?.ToString(),
                        MainEmail = rdr["MainEmail"]?.ToString(),
                        Website = rdr["Website"]?.ToString(),
                        Address1 = rdr["Address1"]?.ToString(),
                        Address2 = rdr["Address2"]?.ToString(),
                        City = rdr["City"]?.ToString(),
                        State = rdr["State"]?.ToString(),
                        Zip = rdr["Zip"]?.ToString(),
                        ProfileNotes = rdr["ProfileNotes"]?.ToString()
                    };
                }
            }

            if (vm == null)
                return null;

            vm.Contacts = GetContacts(id);
            vm.Notes = GetNotes(id);
            vm.AuditHistory = GetAuditHistory(id);
            vm.Documents = GetDocuments(id);
            vm.ContactTypeOptions = ContactTypeOptions;
            vm.ClientTypeOptions = GetClientTypeSelectOptions();
            vm.SftpConfigs = GetSftpConfigs(id);
            vm.Tasks = GetTasks(id);
            vm.AuthTypeOptions = AuthTypeOptions;
            vm.TaskTypeOptions = TaskTypeOptions;
            vm.TaskFrequencyOptions = TaskFrequencyOptions;
            vm.TaskStatusOptions = TaskStatusOptions;
            vm.DeliveryMethodOptions = DeliveryMethodOptions;

            // Sep 08 - was every Client name ever seen in the Accounting transaction tables; now that
            // app's own canonical list, ported from ClientLookupService.GetActiveClientNamesAsync with
            // only the DB prefixes swapped (it runs on Accounting_Data reaching into BR_App, we're the
            // reverse, same server).
            //
            // Ours, not the source's: c.Status = 'Active' ("active" there means an active MAPPING), and
            // the third branch re-admitting mapped aliases - needed because this dropdown also renders
            // each existing mapping's stored value, and a value matching no option makes the browser
            // silently show its first one (the Aug 21 bug).
            vm.AccountingMaps = GetAccountingMaps(id);
            vm.AccountingClientOptions = new List<string>();

            using (var conn2 = new SqlConnection(_connectionString))
            {
                conn2.Open();
                using var acctCmd = new SqlCommand(@"
SELECT MIN(ClientName) AS ClientName FROM (
    SELECT c.ClientName
    FROM dbo.ClientAccountingMap m
    JOIN dbo.Client c ON c.ClientId = m.ClientId
    WHERE m.IsActive = 1
      AND c.Status = 'Active'

    UNION

    SELECT ac.Forwarder AS ClientName
    FROM Accounting_Data.dbo.AccountingClients ac
    WHERE NOT EXISTS (
        SELECT 1 FROM dbo.ClientAccountingMap m2
        WHERE m2.AccountingClientName = ac.Forwarder AND m2.IsActive = 1
    )

    UNION

    SELECT frc.ClientName
    FROM Accounting_Data.dbo.FirmRemitClients frc
    WHERE NOT EXISTS (
        SELECT 1 FROM dbo.ClientAccountingMap m4
        WHERE m4.AccountingClientName = frc.ClientName AND m4.IsActive = 1
    )

    UNION

    SELECT m3.AccountingClientName AS ClientName
    FROM dbo.ClientAccountingMap m3
    WHERE m3.IsActive = 1
) b
WHERE ClientName IS NOT NULL AND LTRIM(RTRIM(ClientName)) <> ''
GROUP BY UPPER(LTRIM(RTRIM(ClientName)))
ORDER BY MIN(ClientName);", conn2);

                using var acctRdr = acctCmd.ExecuteReader();
                while (acctRdr.Read())
                {
                    vm.AccountingClientOptions.Add(acctRdr["ClientName"].ToString());
                }
            }

            vm.Contacts = GetContacts(id);
            vm.Notes = GetNotes(id);
            vm.AuditHistory = GetAuditHistory(id);
            vm.Documents = GetDocuments(id);
            vm.ContactTypeOptions = ContactTypeOptions;
            vm.ClientTypeOptions = GetClientTypeSelectOptions();
            vm.SftpConfigs = GetSftpConfigs(id);
            vm.Tasks = GetTasks(id);
            vm.Portfolios = GetPortfolios(id);
            vm.AuthTypeOptions = AuthTypeOptions;
            vm.TaskTypeOptions = TaskTypeOptions;
            vm.TaskFrequencyOptions = TaskFrequencyOptions;
            vm.TaskStatusOptions = TaskStatusOptions;
            vm.DeliveryMethodOptions = DeliveryMethodOptions;

            return vm;
        }

        // Aug 19 - Client-side equivalent of LawFirmController's GetAccountingMaps, same shape.
        private List<ClientAccountingMapVm> GetAccountingMaps(int clientId)
        {
            var list = new List<ClientAccountingMapVm>();

            using var conn = new SqlConnection(_connectionString);
            using var cmd = new SqlCommand(@"
        SELECT
            ClientAccountingMapId,
            AccountingClientName,
            IsActive
        FROM dbo.ClientAccountingMap
        WHERE ClientId = @ClientId
        ORDER BY AccountingClientName;", conn);

            cmd.Parameters.AddWithValue("@ClientId", clientId);

            conn.Open();

            using var rdr = cmd.ExecuteReader();

            while (rdr.Read())
            {
                list.Add(new ClientAccountingMapVm
                {
                    ClientAccountingMapId = Convert.ToInt32(rdr["ClientAccountingMapId"]),
                    AccountingClientName = rdr["AccountingClientName"]?.ToString(),
                    IsActive = Convert.ToBoolean(rdr["IsActive"])
                });
            }

            return list;
        }

        /// <summary>
        /// Aug 19 - checks whether accountingClientName is already actively mapped to a DIFFERENT
        /// client. Real gap found in LawFirmAccountingMap's own equivalent screen while building this
        /// (AddAccountingMap/UpdateAccountingMap there do a plain INSERT/UPDATE with no such check,
        /// and GetAccountingMaps only ever queries by the current LawFirmId - so a conflicting mapping
        /// on a DIFFERENT firm is invisible from any single firm's own Edit page). Left as-is there
        /// per direct instruction at the time; built in here from the start. Sep 08 - that side is now
        /// in line, see LawFirmController.FindConflictingLawFirmMapRows. excludeMapId lets Update check
        /// every OTHER row without flagging one against itself.
        /// </summary>
        private string? FindConflictingClientMap(string accountingClientName, int excludeMapId = 0)
        {
            var rows = FindConflictingClientMapRows(accountingClientName, excludeMapId);
            return rows.Count == 0 ? null : DescribeConflictOwners(rows);
        }

        /// <summary>
        /// Sep 08 - the one place a conflicting mapping's owner is named: warning, TempData messages,
        /// audit text, EnsureCanonicalAccountingMap's refusal. Inactive clients are labelled, since
        /// "already associated with X" otherwise reads as a live relationship; never filtered out
        /// though - a mapping an Inactive client holds is still a real conflict.
        /// </summary>
        private static string DescribeConflictOwners(List<ConflictingAccountingMapRow> rows)
        {
            return string.Join(", ", rows
                .Select(r => string.Equals(r.Status, "Active", StringComparison.OrdinalIgnoreCase)
                    ? r.ClientName
                    : $"{r.ClientName} (Inactive)")
                .Distinct());
        }

        /// <summary>
        /// Sep 08 - the conflict lookup FindConflictingClientMap has always done, returning the rows
        /// themselves: reassigning needs each row's id to delete it and its ClientId to audit it.
        /// One query with the older method layered on top, so "what counts as a conflict" stays in
        /// one place. Every match, not TOP 1 - clearing only the first would leave it ambiguous.
        /// </summary>
        private List<ConflictingAccountingMapRow> FindConflictingClientMapRows(string accountingClientName, int excludeMapId = 0)
        {
            var rows = new List<ConflictingAccountingMapRow>();

            using var conn = new SqlConnection(_connectionString);
            using var cmd = new SqlCommand(@"
        SELECT
            m.ClientAccountingMapId,
            m.ClientId,
            c.ClientName,
            c.Status
        FROM dbo.ClientAccountingMap m
        JOIN dbo.Client c ON c.ClientId = m.ClientId
        WHERE m.AccountingClientName = @AccountingClientName
          AND m.IsActive = 1
          AND m.ClientAccountingMapId <> @ExcludeMapId
        ORDER BY m.ClientAccountingMapId;", conn);

            cmd.Parameters.AddWithValue("@AccountingClientName", accountingClientName);
            cmd.Parameters.AddWithValue("@ExcludeMapId", excludeMapId);

            conn.Open();

            using var rdr = cmd.ExecuteReader();

            while (rdr.Read())
            {
                rows.Add(new ConflictingAccountingMapRow
                {
                    ClientAccountingMapId = Convert.ToInt32(rdr["ClientAccountingMapId"]),
                    ClientId = Convert.ToInt32(rdr["ClientId"]),
                    ClientName = rdr["ClientName"]?.ToString(),
                    Status = rdr["Status"]?.ToString()
                });
            }

            return rows;
        }

        /// <summary>Sep 08 - a ClientAccountingMap row on some OTHER client already actively claiming
        /// the name being mapped. Controller-private; nothing outside this flow needs it.</summary>
        private sealed class ConflictingAccountingMapRow
        {
            public int ClientAccountingMapId { get; set; }
            public int ClientId { get; set; }
            public string? ClientName { get; set; }

            /// <summary>The owning client's status - labelled in the warning, never filtered on.</summary>
            public string? Status { get; set; }
        }

        /// <summary>
        /// #1/#2/#3 (Aug 21) - ensures a Client's own canonical name always has a matching,
        /// resolvable row in ClientAccountingMap, so any future lookup ("given this raw
        /// accounting-side ClientName, what's the real Client?") is always a single join through
        /// this table - never a two-step "check Client.ClientName first, then fall back to the
        /// map" special case someone has to remember to implement correctly every time. Deliberately
        /// diverges from LawFirmAccountingMap's shape (alias-only) - that table already has real data
        /// and possibly existing consumers assuming alias-only, so changing it now risks breaking
        /// something unseen; ClientAccountingMap is brand new, so this is the one moment to get the
        /// shape right before habits form around it.
        ///
        /// Called from Create (new client) and from Edit only when ClientName actually changed (a
        /// rename) - never unconditionally on every Edit save. On rename, the OLD name's row is
        /// deliberately left exactly as-is, not updated or deactivated - historical Accounting
        /// transactions created before the rename still carry the old name string and need to keep
        /// resolving through it forever. The new name just gets its own additional row alongside it.
        /// Same add-don't-destructively-edit philosophy as reverse-and-repost elsewhere in this
        /// project, applied here to a lookup table instead of a JE.
        ///
        /// Idempotent two ways: (a) if this exact ClientId+name is already an active row, no-op - no
        /// duplicate row for the same client; (b) if the name is already actively claimed by a
        /// DIFFERENT client, refuses rather than silently creating a conflicting mapping - reuses
        /// FindConflictingClientMap, the same detection Add/UpdateAccountingMap use.
        ///
        /// Sep 08 - still REFUSES here, where those two actions now offer to reassign. This is a side
        /// effect of an already-saved rename, so there's no chosen mapping to hang a confirmation off,
        /// and deleting another client's mapping mid-save is what that confirmation exists to prevent.
        /// </summary>
        private void EnsureCanonicalAccountingMap(int clientId, string? clientName)
        {
            if (clientId <= 0 || string.IsNullOrWhiteSpace(clientName)) return;

            using (var conn = new SqlConnection(_connectionString))
            using (var checkCmd = new SqlCommand(@"
SELECT COUNT(*) FROM dbo.ClientAccountingMap
WHERE ClientId = @ClientId AND AccountingClientName = @AccountingClientName AND IsActive = 1;", conn))
            {
                checkCmd.Parameters.AddWithValue("@ClientId", clientId);
                checkCmd.Parameters.AddWithValue("@AccountingClientName", clientName);
                conn.Open();
                var alreadyMapped = (int)checkCmd.ExecuteScalar();
                if (alreadyMapped > 0) return;
            }

            var conflict = FindConflictingClientMap(clientName);
            if (conflict != null)
            {
                // A different client already actively claims this exact name - refuse rather than
                // create a second active mapping for it. This runs from Create/Edit (the client save
                // has already succeeded by this point), not a dedicated mapping action, so this
                // surfaces as a non-blocking warning rather than failing the whole save.
                TempData["AccountingMapWarning"] =
                    $"Note: could not auto-map \"{clientName}\" as this client's own accounting name - " +
                    $"it's already actively mapped to {conflict}.";
                return;
            }

            using var insertConn = new SqlConnection(_connectionString);
            using var insertCmd = new SqlCommand(@"
INSERT INTO dbo.ClientAccountingMap
(
    ClientId,
    AccountingClientName,
    IsActive,
    CreatedDate
)
VALUES
(
    @ClientId,
    @AccountingClientName,
    1,
    GETDATE()
);", insertConn);

            insertCmd.Parameters.AddWithValue("@ClientId", clientId);
            insertCmd.Parameters.AddWithValue("@AccountingClientName", clientName);

            insertConn.Open();
            insertCmd.ExecuteNonQuery();
        }

        /// <summary>
        /// #3 (Aug 21) - fetches a Client's current name before an Edit's UPDATE runs, so the rename
        /// check in Edit() has a real before/after to compare, not just the posted (new) value alone.
        /// </summary>
        private string? GetClientNameById(int clientId)
        {
            using var conn = new SqlConnection(_connectionString);
            using var cmd = new SqlCommand("SELECT ClientName FROM dbo.Client WHERE ClientId = @ClientId;", conn);
            cmd.Parameters.AddWithValue("@ClientId", clientId);
            conn.Open();
            return cmd.ExecuteScalar() as string;
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public IActionResult AddAccountingMap(int clientId, string accountingClientName, bool isActive = true, bool confirmReassign = false)
        {
            if (clientId <= 0 || string.IsNullOrWhiteSpace(accountingClientName))
            {
                TempData["Error"] = "Please select an accounting client.";
                return RedirectToAction(nameof(Edit), new { id = clientId });
            }

            if (isActive)
            {
                // Sep 08 - re-runs on the CONFIRMED post too: the confirmation went out to the browser
                // and back, so the rows it described are a claim about the past. These are the real ones.
                var conflicts = FindConflictingClientMapRows(accountingClientName);
                if (conflicts.Count > 0)
                {
                    if (!confirmReassign)
                    {
                        return PromptAccountingMapReassign(
                            nameof(AddAccountingMap), clientId, 0, accountingClientName, isActive, conflicts);
                    }

                    return ReassignAccountingMap(clientId, accountingClientName, isActive, 0, conflicts,
                        successMessage: "Accounting mapping added.");
                }
            }

            using var conn = new SqlConnection(_connectionString);
            using var cmd = new SqlCommand(@"
INSERT INTO dbo.ClientAccountingMap
(
    ClientId,
    AccountingClientName,
    IsActive,
    CreatedDate
)
VALUES
(
    @ClientId,
    @AccountingClientName,
    @IsActive,
    GETDATE()
);", conn);

            cmd.Parameters.AddWithValue("@ClientId", clientId);
            cmd.Parameters.AddWithValue("@AccountingClientName", accountingClientName);
            cmd.Parameters.AddWithValue("@IsActive", isActive);

            conn.Open();
            cmd.ExecuteNonQuery();

            TempData["Message"] = "Accounting mapping added.";
            return RedirectToAction(nameof(Edit), new { id = clientId });
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public IActionResult DeleteAccountingMap(int clientAccountingMapId, int clientId)
        {
            if (clientAccountingMapId <= 0 || clientId <= 0)
            {
                TempData["Error"] = "Invalid accounting mapping.";
                return RedirectToAction(nameof(Edit), new { id = clientId });
            }

            using var conn = new SqlConnection(_connectionString);
            using var cmd = new SqlCommand(@"
DELETE FROM dbo.ClientAccountingMap
WHERE ClientAccountingMapId = @ClientAccountingMapId;", conn);

            cmd.Parameters.AddWithValue("@ClientAccountingMapId", clientAccountingMapId);

            conn.Open();
            cmd.ExecuteNonQuery();

            TempData["Message"] = "Accounting mapping deleted.";
            return RedirectToAction(nameof(Edit), new { id = clientId });
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public IActionResult UpdateAccountingMap(int clientAccountingMapId, int clientId, string accountingClientName, bool confirmReassign = false)
        {
            bool isActive = Request.Form["isActive"].Any(x => x == "true");

            if (clientAccountingMapId <= 0 || clientId <= 0 || string.IsNullOrWhiteSpace(accountingClientName))
            {
                TempData["Error"] = "Invalid accounting mapping.";
                return RedirectToAction(nameof(Edit), new { id = clientId });
            }

            if (isActive)
            {
                var conflicts = FindConflictingClientMapRows(accountingClientName, excludeMapId: clientAccountingMapId);
                if (conflicts.Count > 0)
                {
                    if (!confirmReassign)
                    {
                        return PromptAccountingMapReassign(
                            nameof(UpdateAccountingMap), clientId, clientAccountingMapId, accountingClientName, isActive, conflicts);
                    }

                    return ReassignAccountingMap(clientId, accountingClientName, isActive, clientAccountingMapId, conflicts,
                        successMessage: "Accounting mapping updated.");
                }
            }

            using var conn = new SqlConnection(_connectionString);
            using var cmd = new SqlCommand(@"
UPDATE dbo.ClientAccountingMap
SET
    AccountingClientName = @AccountingClientName,
    IsActive = @IsActive
WHERE ClientAccountingMapId = @ClientAccountingMapId;", conn);

            cmd.Parameters.AddWithValue("@ClientAccountingMapId", clientAccountingMapId);
            cmd.Parameters.AddWithValue("@AccountingClientName", accountingClientName);
            cmd.Parameters.AddWithValue("@IsActive", isActive);

            conn.Open();
            cmd.ExecuteNonQuery();

            TempData["Message"] = "Accounting mapping updated.";
            return RedirectToAction(nameof(Edit), new { id = clientId });
        }

        /// <summary>
        /// Sep 08 - first half of the reassignment: stash the attempt and redirect to Edit, where the
        /// modal re-offers it with confirmReassign set. Writes nothing, so cancelling leaves both
        /// clients untouched - the reason this is a round trip and not a flag on the original write.
        /// </summary>
        private IActionResult PromptAccountingMapReassign(
            string postAction,
            int clientId,
            int clientAccountingMapId,
            string accountingClientName,
            bool isActive,
            List<ConflictingAccountingMapRow> conflicts)
        {
            var pending = new ClientAccountingMapConflictVm
            {
                ClientId = clientId,
                ClientAccountingMapId = clientAccountingMapId,
                AccountingClientName = accountingClientName,
                IsActive = isActive,
                PostAction = postAction,
                ExistingClientName = DescribeConflictOwners(conflicts)
            };

            TempData[AccountingMapConflictKey] = JsonSerializer.Serialize(pending);
            return RedirectToAction(nameof(Edit), new { id = clientId });
        }

        /// <summary>
        /// Sep 08 - the confirmed half: delete the old holder's row, then write this client's, in ONE
        /// transaction - old-first so the states never overlap, transactional so a failure between
        /// them can't strand the mapping deleted-but-never-recreated.
        ///
        /// Hard DELETE rather than IsActive = 0, per instruction: if the row was the other client's
        /// canonical self-map, only a rename recreates it. Audits run after the commit because
        /// AddAudit opens its own connection.
        /// </summary>
        private IActionResult ReassignAccountingMap(
            int clientId,
            string accountingClientName,
            bool isActive,
            int clientAccountingMapId,
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
DELETE FROM dbo.ClientAccountingMap
WHERE ClientAccountingMapId = @ClientAccountingMapId;", conn, tx);

                    deleteCmd.Parameters.AddWithValue("@ClientAccountingMapId", conflict.ClientAccountingMapId);
                    deleteCmd.ExecuteNonQuery();
                }

                // clientAccountingMapId > 0 means the user was editing an existing row of this
                // client's own, so the mapping moves onto that row rather than adding a second one.
                using var writeCmd = clientAccountingMapId > 0
                    ? new SqlCommand(@"
UPDATE dbo.ClientAccountingMap
SET
    AccountingClientName = @AccountingClientName,
    IsActive = @IsActive
WHERE ClientAccountingMapId = @ClientAccountingMapId;", conn, tx)
                    : new SqlCommand(@"
INSERT INTO dbo.ClientAccountingMap
(
    ClientId,
    AccountingClientName,
    IsActive,
    CreatedDate
)
VALUES
(
    @ClientId,
    @AccountingClientName,
    @IsActive,
    GETDATE()
);", conn, tx);

                writeCmd.Parameters.AddWithValue("@AccountingClientName", accountingClientName);
                writeCmd.Parameters.AddWithValue("@IsActive", isActive);

                if (clientAccountingMapId > 0)
                    writeCmd.Parameters.AddWithValue("@ClientAccountingMapId", clientAccountingMapId);
                else
                    writeCmd.Parameters.AddWithValue("@ClientId", clientId);

                writeCmd.ExecuteNonQuery();

                tx.Commit();
            }
            catch (Exception ex)
            {
                tx.Rollback();

                TempData["Error"] = $"Could not reassign \"{accountingClientName}\": {ex.Message} " +
                    "No mappings were changed.";
                return RedirectToAction(nameof(Edit), new { id = clientId });
            }

            var previousOwners = DescribeConflictOwners(conflicts);
            var currentClientName = GetClientNameById(clientId) ?? $"client {clientId}";

            foreach (var conflict in conflicts)
            {
                AddAudit(conflict.ClientId, "AccountingMap", conflict.ClientAccountingMapId, "Delete",
                    $"Accounting mapping \"{accountingClientName}\" reassigned to {currentClientName}");
            }

            AddAudit(clientId, "AccountingMap", clientAccountingMapId > 0 ? (int?)clientAccountingMapId : null, "Update",
                $"Accounting mapping \"{accountingClientName}\" reassigned from {previousOwners}");

            TempData["Message"] = $"{successMessage} \"{accountingClientName}\" was reassigned from {previousOwners}.";
            return RedirectToAction(nameof(Edit), new { id = clientId });
        }

        private List<ClientPortfolioVm> GetPortfolios(int clientId)
        {
            var rows = new List<ClientPortfolioVm>();

            using var conn = new SqlConnection(_connectionString);
            using var cmd = new SqlCommand(@"
SELECT
    ClientPortfolioId,
    ClientId,
    PortfolioName,
    PortfolioCode,
    ForwNo,
    IsActive,
    CreatedDate,
    UpdatedDate
FROM dbo.ClientPortfolio
WHERE ClientId = @ClientId
ORDER BY PortfolioName;", conn);

            cmd.Parameters.AddWithValue("@ClientId", clientId);

            conn.Open();
            using var rdr = cmd.ExecuteReader();

            while (rdr.Read())
            {
                rows.Add(new ClientPortfolioVm
                {
                    ClientPortfolioId = Convert.ToInt32(rdr["ClientPortfolioId"]),
                    ClientId = Convert.ToInt32(rdr["ClientId"]),
                    PortfolioName = rdr["PortfolioName"]?.ToString(),
                    PortfolioCode = rdr["PortfolioCode"] == DBNull.Value ? null : rdr["PortfolioCode"]?.ToString(),
                    ForwNo = rdr["ForwNo"]?.ToString(),
                    IsActive = Convert.ToBoolean(rdr["IsActive"]),
                    CreatedDate = rdr["CreatedDate"] == DBNull.Value ? null : Convert.ToDateTime(rdr["CreatedDate"]),
                    UpdatedDate = rdr["UpdatedDate"] == DBNull.Value ? null : Convert.ToDateTime(rdr["UpdatedDate"])
                });
            }

            return rows;
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public IActionResult AddPortfolio(ClientPortfolioVm vm)
        {
            if (vm.ClientId <= 0)
            {
                TempData["Error"] = "Invalid client.";
                return RedirectToAction(nameof(Edit), new { id = vm.ClientId });
            }

            if (string.IsNullOrWhiteSpace(vm.PortfolioName))
            {
                TempData["Error"] = "Portfolio Name is required.";
                return RedirectToAction(nameof(Edit), new { id = vm.ClientId });
            }

            if (string.IsNullOrWhiteSpace(vm.ForwNo))
            {
                TempData["Error"] = "CLS FORW_NO is required.";
                return RedirectToAction(nameof(Edit), new { id = vm.ClientId });
            }

            using var conn = new SqlConnection(_connectionString);
            conn.Open();

            using (var dupCmd = new SqlCommand(@"
SELECT COUNT(*)
FROM dbo.ClientPortfolio
WHERE ClientId = @ClientId
  AND CAST(ForwNo AS varchar(50)) = @ForwNo", conn))
            {
                dupCmd.Parameters.AddWithValue("@ClientId", vm.ClientId);
                dupCmd.Parameters.AddWithValue("@ForwNo", vm.ForwNo.Trim());
                dupCmd.Parameters.AddWithValue("@ClientPortfolioId", vm.ClientPortfolioId);

                int exists = Convert.ToInt32(dupCmd.ExecuteScalar());
                if (exists > 0)
                {
                    TempData["Error"] = "That CLS FORW_NO already exists for this client.";
                    return RedirectToAction(nameof(Edit), new { id = vm.ClientId });
                }
            }

            using (var cmd = new SqlCommand(@"
INSERT INTO dbo.ClientPortfolio
(
    ClientId,
    PortfolioName,
    PortfolioCode,
    ForwNo,
    IsActive,
    CreatedDate,
    UpdatedDate
)
VALUES
(
    @ClientId,
    @PortfolioName,
    @PortfolioCode,
    @ForwNo,
    @IsActive,
    GETDATE(),
    GETDATE()
);", conn))
            {
                cmd.Parameters.AddWithValue("@ClientId", vm.ClientId);
                cmd.Parameters.AddWithValue("@PortfolioName", vm.PortfolioName.Trim());
                cmd.Parameters.AddWithValue("@PortfolioCode", (object?)vm.PortfolioCode ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@ForwNo", vm.ForwNo.Trim());
                cmd.Parameters.AddWithValue("@IsActive", vm.IsActive);

                cmd.ExecuteNonQuery();
            }

            AddAudit(vm.ClientId, "Portfolio", null, "Insert", $"Added portfolio {vm.PortfolioName}");

            TempData["Message"] = "Portfolio mapping added.";
            return RedirectToAction(nameof(Edit), new { id = vm.ClientId });
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public IActionResult UpdatePortfolio(ClientPortfolioVm vm)
        {
            if (vm.ClientPortfolioId <= 0 || vm.ClientId <= 0)
            {
                TempData["Error"] = "Invalid portfolio.";
                return RedirectToAction(nameof(Edit), new { id = vm.ClientId });
            }

            if (string.IsNullOrWhiteSpace(vm.PortfolioName))
            {
                TempData["Error"] = "Portfolio Name is required.";
                return RedirectToAction(nameof(Edit), new { id = vm.ClientId });
            }

            if (string.IsNullOrWhiteSpace(vm.ForwNo))
            {
                TempData["Error"] = "CLS FORW_NO is required.";
                return RedirectToAction(nameof(Edit), new { id = vm.ClientId });
            }

            using var conn = new SqlConnection(_connectionString);
            conn.Open();

            using (var dupCmd = new SqlCommand(@"
SELECT COUNT(*)
FROM dbo.ClientPortfolio
WHERE ClientId = @ClientId
  AND CAST(ForwNo AS varchar(50)) = @ForwNo
  AND ClientPortfolioId <> @ClientPortfolioId;", conn))
            {
                dupCmd.Parameters.AddWithValue("@ClientId", vm.ClientId);
                dupCmd.Parameters.AddWithValue("@ForwNo", vm.ForwNo.Trim());
                dupCmd.Parameters.AddWithValue("@ClientPortfolioId", vm.ClientPortfolioId);

                int exists = Convert.ToInt32(dupCmd.ExecuteScalar());
                if (exists > 0)
                {
                    TempData["Error"] = "That CLS FORW_NO already exists for this client.";
                    return RedirectToAction(nameof(Edit), new { id = vm.ClientId });
                }
            }

            using (var cmd = new SqlCommand(@"
UPDATE dbo.ClientPortfolio
SET
    PortfolioName = @PortfolioName,
    PortfolioCode = @PortfolioCode,
    ForwNo = @ForwNo,
    IsActive = @IsActive,
    UpdatedDate = GETDATE()
WHERE ClientPortfolioId = @ClientPortfolioId;", conn))
            {
                cmd.Parameters.AddWithValue("@ClientPortfolioId", vm.ClientPortfolioId);
                cmd.Parameters.AddWithValue("@PortfolioName", vm.PortfolioName.Trim());
                cmd.Parameters.AddWithValue("@PortfolioCode", (object?)vm.PortfolioCode ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@ForwNo", vm.ForwNo.Trim());
                cmd.Parameters.AddWithValue("@IsActive", vm.IsActive);

                cmd.ExecuteNonQuery();
            }

            AddAudit(vm.ClientId, "Portfolio", vm.ClientPortfolioId, "Update", $"Updated portfolio {vm.PortfolioName}");

            TempData["Message"] = "Portfolio mapping updated.";
            return RedirectToAction(nameof(Edit), new { id = vm.ClientId });
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public IActionResult DeletePortfolio(int clientPortfolioId, int clientId)
        {
            using var conn = new SqlConnection(_connectionString);
            using var cmd = new SqlCommand(@"
DELETE FROM dbo.ClientPortfolio
WHERE ClientPortfolioId = @ClientPortfolioId;", conn);

            cmd.Parameters.AddWithValue("@ClientPortfolioId", clientPortfolioId);

            conn.Open();
            cmd.ExecuteNonQuery();

            AddAudit(clientId, "Portfolio", clientPortfolioId, "Delete", "Deleted portfolio mapping");

            TempData["Message"] = "Portfolio mapping deleted.";
            return RedirectToAction(nameof(Edit), new { id = clientId });
        }

        private List<ClientSftpConfigVm> GetSftpConfigs(int clientId)
        {
            var list = new List<ClientSftpConfigVm>();

            using var conn = new SqlConnection(_connectionString);
            using var cmd = new SqlCommand(@"
SELECT *
FROM dbo.ClientSftpConfig
WHERE ClientId = @ClientId
ORDER BY ConfigName;", conn);

            cmd.Parameters.AddWithValue("@ClientId", clientId);

            conn.Open();
            using var rdr = cmd.ExecuteReader();
            while (rdr.Read())
            {
                list.Add(new ClientSftpConfigVm
                {
                    ClientSftpConfigId = Convert.ToInt32(rdr["ClientSftpConfigId"]),
                    ClientId = Convert.ToInt32(rdr["ClientId"]),
                    ConfigName = rdr["ConfigName"]?.ToString(),
                    Host = rdr["Host"]?.ToString(),
                    Port = rdr["Port"] == DBNull.Value ? 22 : Convert.ToInt32(rdr["Port"]),
                    Username = rdr["Username"]?.ToString(),
                    AuthType = rdr["AuthType"]?.ToString(),
                    PasswordValue = rdr["PasswordValue"]?.ToString(),
                    KeyFileName = rdr["KeyFileName"]?.ToString(),
                    RemoteInboundPath = rdr["RemoteInboundPath"]?.ToString(),
                    RemoteOutboundPath = rdr["RemoteOutboundPath"]?.ToString(),
                    ArchivePath = rdr["ArchivePath"]?.ToString(),
                    ErrorPath = rdr["ErrorPath"]?.ToString(),
                    IsActive = rdr["IsActive"] != DBNull.Value && Convert.ToBoolean(rdr["IsActive"]),
                    Notes = rdr["Notes"]?.ToString(),
                    LastConnectionTestDate = rdr["LastConnectionTestDate"] == DBNull.Value ? null : (DateTime?)Convert.ToDateTime(rdr["LastConnectionTestDate"]),
                    LastConnectionTestStatus = rdr["LastConnectionTestStatus"]?.ToString()
                });
            }

            return list;
        }

        private List<ClientTaskVm> GetTasks(int clientId)
        {
            var list = new List<ClientTaskVm>();

            using var conn = new SqlConnection(_connectionString);
            using var cmd = new SqlCommand(@"
SELECT
    t.*,
    s.ConfigName AS SftpConfigName
FROM dbo.ClientTask t
LEFT JOIN dbo.ClientSftpConfig s ON t.SftpConfigId = s.ClientSftpConfigId
WHERE t.ClientId = @ClientId
ORDER BY t.TaskType, t.TaskName;", conn);

            cmd.Parameters.AddWithValue("@ClientId", clientId);

            conn.Open();
            using var rdr = cmd.ExecuteReader();
            while (rdr.Read())
            {
                list.Add(new ClientTaskVm
                {
                    ClientTaskId = Convert.ToInt32(rdr["ClientTaskId"]),
                    ClientId = Convert.ToInt32(rdr["ClientId"]),
                    TaskType = rdr["TaskType"]?.ToString(),
                    TaskName = rdr["TaskName"]?.ToString(),
                    Frequency = rdr["Frequency"]?.ToString(),
                    DueDay = rdr["DueDay"]?.ToString(),
                    AssignedTo = rdr["AssignedTo"]?.ToString(),
                    Status = rdr["Status"]?.ToString(),
                    DeliveryMethod = rdr["DeliveryMethod"]?.ToString(),
                    ClientSftpConfigId = rdr["SftpConfigId"] == DBNull.Value ? null : (int?)Convert.ToInt32(rdr["SftpConfigId"]),
                    SftpConfigName = rdr["SftpConfigName"]?.ToString(),
                    Notes = rdr["Notes"]?.ToString(),
                    LastCompletedDate = rdr["LastCompletedDate"] == DBNull.Value ? null : (DateTime?)Convert.ToDateTime(rdr["LastCompletedDate"]),
                    NextDueDate = rdr["NextDueDate"] == DBNull.Value ? null : (DateTime?)Convert.ToDateTime(rdr["NextDueDate"])
                });
            }

            return list;
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public IActionResult AddSftpConfig(ClientSftpConfigVm vm, string returnAction = "Edit")
        {
            using var conn = new SqlConnection(_connectionString);
            using var cmd = new SqlCommand(@"
INSERT INTO dbo.ClientSftpConfig
(
    ClientId, ConfigName, Host, Port, Username, AuthType, PasswordValue, KeyFileName,
    RemoteInboundPath, RemoteOutboundPath, ArchivePath, ErrorPath, IsActive, Notes,
    CreatedDate, UpdatedDate
)
VALUES
(
    @ClientId, @ConfigName, @Host, @Port, @Username, @AuthType, @PasswordValue, @KeyFileName,
    @RemoteInboundPath, @RemoteOutboundPath, @ArchivePath, @ErrorPath, @IsActive, @Notes,
    GETDATE(), GETDATE()
);", conn);

            cmd.Parameters.AddWithValue("@ClientId", vm.ClientId);
            cmd.Parameters.AddWithValue("@ConfigName", (object?)vm.ConfigName ?? DBNull.Value);
            if(vm.Host == null)
            {
                cmd.Parameters.AddWithValue("@Host", "LegalStream");
                cmd.Parameters.AddWithValue("@Port", "22");
                cmd.Parameters.AddWithValue("@Username", "LegalStream");
                cmd.Parameters.AddWithValue("@AuthType", "Password");
                cmd.Parameters.AddWithValue("@PasswordValue", "LegalStream");
                cmd.Parameters.AddWithValue("@KeyFileName", "LegalStream");
                cmd.Parameters.AddWithValue("@RemoteInboundPath", (object?)vm.RemoteInboundPath ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@RemoteOutboundPath", (object?)vm.RemoteOutboundPath ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@ArchivePath", (object?)vm.ArchivePath ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@ErrorPath", (object?)vm.ErrorPath ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@IsActive", vm.IsActive);
                cmd.Parameters.AddWithValue("@Notes", (object?)vm.Notes ?? DBNull.Value);
            
            }
            else
            {
                cmd.Parameters.AddWithValue("@Host", (object?)vm.Host ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@Port", vm.Port);
                cmd.Parameters.AddWithValue("@Username", (object?)vm.Username ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@AuthType", (object?)vm.AuthType ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@PasswordValue", (object?)vm.PasswordValue ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@KeyFileName", (object?)vm.KeyFileName ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@RemoteInboundPath", (object?)vm.RemoteInboundPath ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@RemoteOutboundPath", (object?)vm.RemoteOutboundPath ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@ArchivePath", (object?)vm.ArchivePath ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@ErrorPath", (object?)vm.ErrorPath ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@IsActive", vm.IsActive);
                cmd.Parameters.AddWithValue("@Notes", (object?)vm.Notes ?? DBNull.Value);

            }
              
           

            conn.Open();
            cmd.ExecuteNonQuery();

            AddAudit(vm.ClientId, "SFTP", null, "Insert", $"Added SFTP config {vm.ConfigName}");
            TempData["Message"] = "SFTP configuration added.";
            return RedirectToAction(returnAction, new { id = vm.ClientId });
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public IActionResult DeleteSftpConfig(int clientSftpConfigId, int clientId, string returnAction = "Edit")
        {
            using var conn = new SqlConnection(_connectionString);
            using var cmd = new SqlCommand("DELETE FROM dbo.ClientSftpConfig WHERE ClientSftpConfigId = @ClientSftpConfigId;", conn);

            cmd.Parameters.AddWithValue("@ClientSftpConfigId", clientSftpConfigId);

            conn.Open();
            cmd.ExecuteNonQuery();

            AddAudit(clientId, "SFTP", clientSftpConfigId, "Delete", "Deleted SFTP configuration");
            TempData["Message"] = "SFTP configuration deleted.";
            return RedirectToAction(returnAction, new { id = clientId });
        }

        private List<ClientContactVm> GetContacts(int clientId)
        {
            var list = new List<ClientContactVm>();

            using var conn = new SqlConnection(_connectionString);
            using var cmd = new SqlCommand("SELECT * FROM dbo.ClientContact WHERE ClientId = @ClientId ORDER BY ContactType, ContactName;", conn);
            cmd.Parameters.AddWithValue("@ClientId", clientId);

            conn.Open();
            using var rdr = cmd.ExecuteReader();
            while (rdr.Read())
            {
                list.Add(new ClientContactVm
                {
                    ClientContactId = Convert.ToInt32(rdr["ClientContactId"]),
                    ClientId = Convert.ToInt32(rdr["ClientId"]),
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

        [HttpPost]
        [ValidateAntiForgeryToken]
        public IActionResult AddTask(ClientTaskVm vm, string returnAction = "Edit")
        {
            using var conn = new SqlConnection(_connectionString);
            using var cmd = new SqlCommand(@"
INSERT INTO dbo.ClientTask
(
    ClientId, TaskType, TaskName, Frequency, DueDay, AssignedTo, Status,
    DeliveryMethod, SftpConfigId, Notes, LastCompletedDate, NextDueDate,
    CreatedDate, UpdatedDate
)
VALUES
(
    @ClientId, @TaskType, @TaskName, @Frequency, @DueDay, @AssignedTo, @Status,
    @DeliveryMethod, @ClientSftpConfigId, @Notes, @LastCompletedDate, @NextDueDate,
    GETDATE(), GETDATE()
);", conn);

            cmd.Parameters.AddWithValue("@ClientId", vm.ClientId);
            cmd.Parameters.AddWithValue("@TaskType", (object?)vm.TaskType ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@TaskName", (object?)vm.TaskName ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@Frequency", (object?)vm.Frequency ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@DueDay", (object?)vm.DueDay ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@AssignedTo", (object?)vm.AssignedTo ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@Status", (object?)vm.Status ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@DeliveryMethod", (object?)vm.DeliveryMethod ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@ClientSftpConfigId", (object?)vm.ClientSftpConfigId ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@Notes", (object?)vm.Notes ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@LastCompletedDate", (object?)vm.LastCompletedDate ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@NextDueDate", (object?)vm.NextDueDate ?? DBNull.Value);

            conn.Open();
            cmd.ExecuteNonQuery();

            AddAudit(vm.ClientId, "Task", null, "Insert", $"Added client task {vm.TaskName}");
            TempData["Message"] = "Task added.";
            return RedirectToAction(returnAction, new { id = vm.ClientId });
        }

        [HttpGet]
        public IActionResult ImportPortfolio()
        {
            return View();
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public IActionResult ImportPortfolio(IFormFile file)
        {
            if (file == null || file.Length == 0)
            {
                TempData["Error"] = "Please upload a CSV file.";
                return View();
            }

            var errors = new List<string>();
            int inserted = 0;
            int updated = 0;
            int rowNumber = 0;

            using var stream = file.OpenReadStream();
            using var parser = new TextFieldParser(stream);
            parser.TextFieldType = FieldType.Delimited;
            parser.SetDelimiters(",");
            parser.HasFieldsEnclosedInQuotes = true;

            if (!parser.EndOfData)
            {
                parser.ReadFields(); // header row
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
                    if (fields == null || fields.Length < 5)
                    {
                        errors.Add($"Row {rowNumber}: Not enough columns.");
                        continue;
                    }

                    var clientName = fields[0]?.Trim();
                    var portfolioName = fields[1]?.Trim();
                    var portfolioCode = fields[2]?.Trim();
                    var forwNo = fields[3]?.Trim();
                    var isActiveText = fields[4]?.Trim();

                    if (string.IsNullOrWhiteSpace(clientName))
                    {
                        errors.Add($"Row {rowNumber}: ClientName is required.");
                        continue;
                    }

                    if (string.IsNullOrWhiteSpace(portfolioName))
                    {
                        errors.Add($"Row {rowNumber}: PortfolioName is required.");
                        continue;
                    }

                    if (string.IsNullOrWhiteSpace(forwNo))
                    {
                        errors.Add($"Row {rowNumber}: ForwNo is required.");
                        continue;
                    }

                    bool isActive = true;
                    if (!string.IsNullOrWhiteSpace(isActiveText))
                    {
                        if (string.Equals(isActiveText, "true", StringComparison.OrdinalIgnoreCase) ||
                            string.Equals(isActiveText, "yes", StringComparison.OrdinalIgnoreCase) ||
                            string.Equals(isActiveText, "1", StringComparison.OrdinalIgnoreCase) ||
                            string.Equals(isActiveText, "active", StringComparison.OrdinalIgnoreCase))
                        {
                            isActive = true;
                        }
                        else if (string.Equals(isActiveText, "false", StringComparison.OrdinalIgnoreCase) ||
                                 string.Equals(isActiveText, "no", StringComparison.OrdinalIgnoreCase) ||
                                 string.Equals(isActiveText, "0", StringComparison.OrdinalIgnoreCase) ||
                                 string.Equals(isActiveText, "inactive", StringComparison.OrdinalIgnoreCase))
                        {
                            isActive = false;
                        }
                        else
                        {
                            errors.Add($"Row {rowNumber}: IsActive must be true/false, yes/no, 1/0, or active/inactive.");
                            continue;
                        }
                    }

                    int clientId;
                    using (var clientCmd = new SqlCommand(@"
SELECT ClientId
FROM dbo.Client
WHERE ClientName = @ClientName;", conn))
                    {
                        clientCmd.Parameters.AddWithValue("@ClientName", clientName);
                        var clientResult = clientCmd.ExecuteScalar();

                        if (clientResult == null || clientResult == DBNull.Value)
                        {
                            errors.Add($"Row {rowNumber}: Client '{clientName}' not found.");
                            continue;
                        }

                        clientId = Convert.ToInt32(clientResult);
                    }

                    int? existingId = null;
                    using (var existingCmd = new SqlCommand(@"
SELECT ClientPortfolioId
FROM dbo.ClientPortfolio
WHERE ClientId = @ClientId
  AND PortfolioName = @PortfolioName;", conn))
                    {
                        existingCmd.Parameters.AddWithValue("@ClientId", clientId);
                        existingCmd.Parameters.AddWithValue("@PortfolioName", portfolioName);

                        var existingResult = existingCmd.ExecuteScalar();
                        if (existingResult != null && existingResult != DBNull.Value)
                        {
                            existingId = Convert.ToInt32(existingResult);
                        }
                    }

                    if (existingId.HasValue)
                    {
                        using var updateCmd = new SqlCommand(@"
UPDATE dbo.ClientPortfolio
SET
    PortfolioCode = @PortfolioCode,
    ForwNo = @ForwNo,
    IsActive = @IsActive,
    UpdatedDate = GETDATE()
WHERE ClientPortfolioId = @ClientPortfolioId;", conn);

                        updateCmd.Parameters.AddWithValue("@ClientPortfolioId", existingId.Value);
                        updateCmd.Parameters.AddWithValue("@PortfolioCode", (object?)portfolioCode ?? DBNull.Value);
                        updateCmd.Parameters.AddWithValue("@ForwNo", forwNo);
                        updateCmd.Parameters.AddWithValue("@IsActive", isActive);

                        updateCmd.ExecuteNonQuery();

                        AddAudit(clientId, "Portfolio", existingId.Value, "Update", $"Imported update for portfolio {portfolioName}");
                        updated++;
                    }
                    else
                    {
                        using var insertCmd = new SqlCommand(@"
INSERT INTO dbo.ClientPortfolio
(
    ClientId,
    PortfolioName,
    PortfolioCode,
    ForwNo,
    IsActive,
    CreatedDate,
    UpdatedDate
)
VALUES
(
    @ClientId,
    @PortfolioName,
    @PortfolioCode,
    @ForwNo,
    @IsActive,
    GETDATE(),
    GETDATE()
);", conn);

                        insertCmd.Parameters.AddWithValue("@ClientId", clientId);
                        insertCmd.Parameters.AddWithValue("@PortfolioName", portfolioName);
                        insertCmd.Parameters.AddWithValue("@PortfolioCode", (object?)portfolioCode ?? DBNull.Value);
                        insertCmd.Parameters.AddWithValue("@ForwNo", forwNo);
                        insertCmd.Parameters.AddWithValue("@IsActive", isActive);

                        insertCmd.ExecuteNonQuery();

                        AddAudit(clientId, "Portfolio", null, "Insert", $"Imported portfolio {portfolioName}");
                        inserted++;
                    }
                }
                catch (Exception ex)
                {
                    errors.Add($"Row {rowNumber}: {ex.Message}");
                }
            }

            TempData["Message"] = $"{inserted} portfolio mapping(s) inserted, {updated} updated.";
            if (errors.Any())
                TempData["Error"] = string.Join("<br/>", errors);

            return RedirectToAction(nameof(ImportPortfolio));
        }

        [HttpGet]
        public IActionResult DownloadClientPortfolioTemplate()
        {
            var csv = new StringBuilder();
            csv.AppendLine("ClientName,PortfolioName,PortfolioCode,ForwNo,IsActive");
            csv.AppendLine("Channel,Main,CHAN,4578,1");

            return File(Encoding.UTF8.GetBytes(csv.ToString()), "text/csv", "ClientPortfolioTemplate.csv");
        }

    
        [HttpPost]
        [ValidateAntiForgeryToken]
        public IActionResult DeleteTask(int clientTaskId, int clientId, string returnAction = "Edit")
        {
            using var conn = new SqlConnection(_connectionString);
            using var cmd = new SqlCommand("DELETE FROM dbo.ClientTask WHERE ClientTaskId = @ClientTaskId;", conn);

            cmd.Parameters.AddWithValue("@ClientTaskId", clientTaskId);

            conn.Open();
            cmd.ExecuteNonQuery();

            AddAudit(clientId, "Task", clientTaskId, "Delete", "Deleted client task");
            TempData["Message"] = "Task deleted.";
            return RedirectToAction(returnAction, new { id = clientId });
        }

        private List<ClientNoteVm> GetNotes(int clientId)
        {
            var list = new List<ClientNoteVm>();

            using var conn = new SqlConnection(_connectionString);
            using var cmd = new SqlCommand("SELECT * FROM dbo.ClientNote WHERE ClientId = @ClientId ORDER BY EnteredDate DESC, ClientNoteId DESC;", conn);
            cmd.Parameters.AddWithValue("@ClientId", clientId);

            conn.Open();
            using var rdr = cmd.ExecuteReader();
            while (rdr.Read())
            {
                list.Add(new ClientNoteVm
                {
                    ClientNoteId = Convert.ToInt32(rdr["ClientNoteId"]),
                    ClientId = Convert.ToInt32(rdr["ClientId"]),
                    NoteType = rdr["NoteType"]?.ToString(),
                    NoteText = rdr["NoteText"]?.ToString(),
                    EnteredBy = rdr["EnteredBy"]?.ToString(),
                    EnteredDate = rdr["EnteredDate"] == DBNull.Value ? DateTime.MinValue : Convert.ToDateTime(rdr["EnteredDate"])
                });
            }

            return list;
        }

        private List<ClientAuditVm> GetAuditHistory(int clientId)
        {
            var list = new List<ClientAuditVm>();

            using var conn = new SqlConnection(_connectionString);
            using var cmd = new SqlCommand("SELECT * FROM dbo.ClientAudit WHERE ClientId = @ClientId ORDER BY EnteredDate DESC, ClientAuditId DESC;", conn);
            cmd.Parameters.AddWithValue("@ClientId", clientId);

            conn.Open();
            using var rdr = cmd.ExecuteReader();
            while (rdr.Read())
            {
                list.Add(new ClientAuditVm
                {
                    ClientAuditId = Convert.ToInt32(rdr["ClientAuditId"]),
                    ClientId = Convert.ToInt32(rdr["ClientId"]),
                    EntityType = rdr["EntityType"]?.ToString(),
                    EntityId = rdr["EntityId"] == DBNull.Value ? null : (int?)Convert.ToInt32(rdr["EntityId"]),
                    ActionType = rdr["ActionType"]?.ToString(),
                    AuditText = rdr["AuditText"]?.ToString(),
                    EnteredBy = rdr["EnteredBy"]?.ToString(),
                    EnteredDate = rdr["EnteredDate"] == DBNull.Value ? DateTime.MinValue : Convert.ToDateTime(rdr["EnteredDate"])
                });
            }

            return list;
        }

        private List<ClientDocumentVm> GetDocuments(int clientId)
        {
            var list = new List<ClientDocumentVm>();

            using var conn = new SqlConnection(_connectionString);
            using var cmd = new SqlCommand(@"
SELECT
    ClientDocumentId,
    ClientId,
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
FROM dbo.ClientDocument
WHERE ClientId = @ClientId
ORDER BY UploadedDate DESC, ClientDocumentId DESC;", conn);

            cmd.Parameters.AddWithValue("@ClientId", clientId);

            conn.Open();
            using var rdr = cmd.ExecuteReader();
            while (rdr.Read())
            {
                list.Add(new ClientDocumentVm
                {
                    ClientDocumentId = Convert.ToInt32(rdr["ClientDocumentId"]),
                    ClientId = Convert.ToInt32(rdr["ClientId"]),
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

        private void AddAudit(int clientId, string entityType, int? entityId, string actionType, string auditText)
        {
            using var conn = new SqlConnection(_connectionString);
            using var cmd = new SqlCommand(@"
INSERT INTO dbo.ClientAudit
(
    ClientId,
    EntityType,
    EntityId,
    ActionType,
    AuditText,
    EnteredBy,
    EnteredDate
)
VALUES
(
    @ClientId,
    @EntityType,
    @EntityId,
    @ActionType,
    @AuditText,
    @EnteredBy,
    GETDATE()
);", conn);

            cmd.Parameters.AddWithValue("@ClientId", clientId);
            cmd.Parameters.AddWithValue("@EntityType", entityType);
            cmd.Parameters.AddWithValue("@EntityId", (object?)entityId ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@ActionType", actionType);
            cmd.Parameters.AddWithValue("@AuditText", (object?)auditText ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@EnteredBy", User?.Identity?.Name ?? "System");

            conn.Open();
            cmd.ExecuteNonQuery();
        }

        private List<ClientListRowVm> SortRows(List<ClientListRowVm> rows, string sort, string dir)
        {
            bool desc = string.Equals(dir, "DESC", StringComparison.OrdinalIgnoreCase);

            return (sort?.ToUpperInvariant()) switch
            {
                "SHORT_NAME" => desc ? rows.OrderByDescending(x => x.ShortName).ToList() : rows.OrderBy(x => x.ShortName).ToList(),
                "STATUS" => desc ? rows.OrderByDescending(x => x.Status).ToList() : rows.OrderBy(x => x.Status).ToList(),
                "CLIENT_TYPE" => desc ? rows.OrderByDescending(x => x.ClientTypeName).ToList() : rows.OrderBy(x => x.ClientTypeName).ToList(),
                "CONTACT_COUNT" => desc ? rows.OrderByDescending(x => x.ContactCount).ToList() : rows.OrderBy(x => x.ContactCount).ToList(),
                "UPDATED_DATE" => desc ? rows.OrderByDescending(x => x.UpdatedDate).ToList() : rows.OrderBy(x => x.UpdatedDate).ToList(),
                _ => desc ? rows.OrderByDescending(x => x.ClientName).ToList() : rows.OrderBy(x => x.ClientName).ToList()
            };
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public IActionResult AddDocument(
    int clientId,
    string? documentType,
    string? documentName,
    IFormFile? uploadFile,
    DateTime? effectiveDate,
    DateTime? expirationDate,
    string? notes,
    bool isActive = true,
    string returnAction = "Edit")
        {
            if (clientId <= 0)
            {
                TempData["Error"] = "Invalid client.";
                return RedirectToAction(returnAction, new { id = clientId });
            }

            if (uploadFile == null || uploadFile.Length == 0)
            {
                TempData["Error"] = "Please choose a file.";
                return RedirectToAction(returnAction, new { id = clientId });
            }

            byte[] fileData;
            using (var ms = new MemoryStream())
            {
                uploadFile.CopyTo(ms);
                fileData = ms.ToArray();
            }

            using var conn = new SqlConnection(_connectionString);
            using var cmd = new SqlCommand(@"
INSERT INTO dbo.ClientDocument
(
    ClientId,
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
    @ClientId,
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

            cmd.Parameters.AddWithValue("@ClientId", clientId);
            cmd.Parameters.AddWithValue("@DocumentType", (object?)documentType ?? DBNull.Value);
            cmd.Parameters.AddWithValue(
    "@DocumentName",
    (object?)(documentName ?? Path.GetFileNameWithoutExtension(uploadFile.FileName)) ?? DBNull.Value
);
            cmd.Parameters.AddWithValue("@OriginalFileName", (object?)uploadFile.FileName ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@ContentType", (object?)uploadFile.ContentType ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@FileData", fileData);
            cmd.Parameters.AddWithValue("@FileSizeBytes", uploadFile.Length);
            cmd.Parameters.AddWithValue("@Notes", (object?)notes ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@EffectiveDate", (object?)effectiveDate ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@ExpirationDate", (object?)expirationDate ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@IsActive", isActive);
            cmd.Parameters.AddWithValue("@UploadedBy", User?.Identity?.Name ?? "System");

            conn.Open();
            cmd.ExecuteNonQuery();

            AddAudit(clientId, "Document", null, "Insert", $"Added client document {documentName ?? uploadFile.FileName}");

            TempData["Message"] = "Document uploaded.";
            return RedirectToAction(returnAction, new { id = clientId });
        }

        [HttpGet]
        public IActionResult DownloadDocument(int id)
        {
            using var conn = new SqlConnection(_connectionString);
            using var cmd = new SqlCommand(@"
SELECT OriginalFileName, ContentType, FileData
FROM dbo.ClientDocument
WHERE ClientDocumentId = @ClientDocumentId;", conn);

            cmd.Parameters.AddWithValue("@ClientDocumentId", id);

            conn.Open();
            using var rdr = cmd.ExecuteReader();

            if (!rdr.Read())
                return NotFound();

            var fileName = rdr["OriginalFileName"]?.ToString() ?? "document.bin";
            var contentType = rdr["ContentType"]?.ToString() ?? "application/octet-stream";
            var fileData = (byte[])rdr["FileData"];

            return File(fileData, contentType, fileName);
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public IActionResult DeleteDocument(int clientDocumentId, int clientId, string returnAction = "Edit")
        {
            using var conn = new SqlConnection(_connectionString);
            using var cmd = new SqlCommand(@"
DELETE FROM dbo.ClientDocument
WHERE ClientDocumentId = @ClientDocumentId;", conn);

            cmd.Parameters.AddWithValue("@ClientDocumentId", clientDocumentId);

            conn.Open();
            cmd.ExecuteNonQuery();

            AddAudit(clientId, "Document", clientDocumentId, "Delete", "Deleted client document");

            TempData["Message"] = "Document deleted.";
            return RedirectToAction(returnAction, new { id = clientId });
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public IActionResult ReplaceDocument(
            int clientDocumentId,
            int clientId,
            string? documentType,
            string? documentName,
            IFormFile? uploadFile,
            DateTime? effectiveDate,
            DateTime? expirationDate,
            string? notes,
            bool isActive = true,
            string returnAction = "Edit")
        {
            if (uploadFile == null || uploadFile.Length == 0)
            {
                TempData["Error"] = "Please choose a replacement file.";
                return RedirectToAction(returnAction, new { id = clientId });
            }

            byte[] fileData;
            using (var ms = new MemoryStream())
            {
                uploadFile.CopyTo(ms);
                fileData = ms.ToArray();
            }

            using var conn = new SqlConnection(_connectionString);
            using var cmd = new SqlCommand(@"
UPDATE dbo.ClientDocument
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
WHERE ClientDocumentId = @ClientDocumentId;", conn);

            cmd.Parameters.AddWithValue("@ClientDocumentId", clientDocumentId);
            cmd.Parameters.AddWithValue("@DocumentType", (object?)documentType ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@DocumentName", (object?)(documentName ?? Path.GetFileNameWithoutExtension(uploadFile.FileName)) ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@OriginalFileName", (object?)uploadFile.FileName ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@ContentType", (object?)uploadFile.ContentType ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@FileData", fileData);
            cmd.Parameters.AddWithValue("@FileSizeBytes", uploadFile.Length);
            cmd.Parameters.AddWithValue("@Notes", (object?)notes ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@EffectiveDate", (object?)effectiveDate ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@ExpirationDate", (object?)expirationDate ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@IsActive", isActive);
            cmd.Parameters.AddWithValue("@UploadedBy", User?.Identity?.Name ?? "System");

            conn.Open();
            cmd.ExecuteNonQuery();

            AddAudit(clientId, "Document", clientDocumentId, "Update", $"Replaced client document {documentName ?? uploadFile.FileName}");

            TempData["Message"] = "Document replaced.";
            return RedirectToAction(returnAction, new { id = clientId });
        }

        private static readonly List<string> AuthTypeOptions = new()
{
    "Password",
    "Key"
};

        private static readonly List<string> TaskTypeOptions = new()
{
    "Export",
    "Report",
    "Remit",
    "Other"
};

        private static readonly List<string> TaskFrequencyOptions = new()
{
    "Daily",
    "Weekly",
    "Monthly",
    "Ad Hoc"
};

        private static readonly List<string> TaskStatusOptions = new()
{
    "Active",
    "Inactive",
    "Paused"
};

        private static readonly List<string> DeliveryMethodOptions = new()
{
    "SFTP",
    "Email",
    "Manual Upload"
};

        [HttpGet]
        public IActionResult Calendar()
        {
            ViewBag.ClientOptions = GetClientSelectOptions();
            ViewBag.TaskTypeOptions = TaskTypeOptions;
            ViewBag.TaskStatusOptions = TaskStatusOptions;
            ViewBag.TaskFrequencyOptions = TaskFrequencyOptions;

            return View();
        }

        private List<SelectListItem> GetClientSelectOptions()
        {
            var list = new List<SelectListItem>();

            using var conn = new SqlConnection(_connectionString);
            using var cmd = new SqlCommand(@"
SELECT ClientId, ClientName
FROM dbo.Client
ORDER BY ClientName;", conn);

            conn.Open();
            using var rdr = cmd.ExecuteReader();

            while (rdr.Read())
            {
                list.Add(new SelectListItem
                {
                    Value = rdr["ClientId"].ToString(),
                    Text = rdr["ClientName"]?.ToString() ?? ""
                });
            }

            return list;
        }

        [HttpGet]
        public IActionResult CalendarEvents(
    DateTime? start,
    DateTime? end,
    int? clientId,
    string? assignedTo,
    string? taskType,
    string? status,
    string? frequency)
        {
            var rangeStart = (start ?? DateTime.Today.AddMonths(-1)).Date;
            var rangeEnd = (end ?? DateTime.Today.AddMonths(3)).Date;

            var events = new List<object>();

            using var conn = new SqlConnection(_connectionString);

            var sql = @"
SELECT
    t.ClientTaskId,
    t.ClientId,
    c.ClientName,
    c.ShortName,
    t.TaskType,
    t.TaskName,
    t.Frequency,
    t.DueDay,
    t.AssignedTo,
    t.Status,
    t.DeliveryMethod,
    t.Notes,
    t.NextDueDate,
    t.LastCompletedDate,
    t.SftpConfigId
FROM dbo.ClientTask t
INNER JOIN dbo.Client c ON c.ClientId = t.ClientId
WHERE t.NextDueDate IS NOT NULL";

            if (clientId.HasValue)
                sql += " AND t.ClientId = @ClientId";

            if (!string.IsNullOrWhiteSpace(assignedTo))
                sql += " AND t.AssignedTo = @AssignedTo";

            if (!string.IsNullOrWhiteSpace(taskType))
                sql += " AND t.TaskType = @TaskType";

            if (!string.IsNullOrWhiteSpace(status))
                sql += " AND t.Status = @Status";

            if (!string.IsNullOrWhiteSpace(frequency))
                sql += " AND t.Frequency = @Frequency";

            sql += " ORDER BY c.ClientName, t.TaskName;";

            using var cmd = new SqlCommand(sql, conn);

            if (clientId.HasValue)
                cmd.Parameters.AddWithValue("@ClientId", clientId.Value);

            if (!string.IsNullOrWhiteSpace(assignedTo))
                cmd.Parameters.AddWithValue("@AssignedTo", assignedTo);

            if (!string.IsNullOrWhiteSpace(taskType))
                cmd.Parameters.AddWithValue("@TaskType", taskType);

            if (!string.IsNullOrWhiteSpace(status))
                cmd.Parameters.AddWithValue("@Status", status);

            if (!string.IsNullOrWhiteSpace(frequency))
                cmd.Parameters.AddWithValue("@Frequency", frequency);

            conn.Open();
            using var rdr = cmd.ExecuteReader();

            while (rdr.Read())
            {
                int clientTaskId = Convert.ToInt32(rdr["ClientTaskId"]);
                int rowClientId = Convert.ToInt32(rdr["ClientId"]);

                string clientName = rdr["ClientName"]?.ToString() ?? "";
                string shortName = rdr["ShortName"]?.ToString() ?? "";
                string taskTypeValue = rdr["TaskType"]?.ToString() ?? "";
                string taskName = rdr["TaskName"]?.ToString() ?? "Task";
                string frequencyValue = rdr["Frequency"]?.ToString() ?? "";
                string dueDay = rdr["DueDay"]?.ToString() ?? "";
                string assignedToValue = rdr["AssignedTo"]?.ToString() ?? "";
                string statusValue = rdr["Status"]?.ToString() ?? "";
                string deliveryMethod = rdr["DeliveryMethod"]?.ToString() ?? "";
                string notes = rdr["Notes"]?.ToString() ?? "";

                int? sftpConfigId = rdr["SftpConfigId"] == DBNull.Value
                    ? null
                    : Convert.ToInt32(rdr["SftpConfigId"]);

                DateTime anchorDate = Convert.ToDateTime(rdr["NextDueDate"]).Date;
                DateTime? lastCompletedDate = rdr["LastCompletedDate"] == DBNull.Value
                    ? null
                    : Convert.ToDateTime(rdr["LastCompletedDate"]);

                string clientLabel = !string.IsNullOrWhiteSpace(shortName) ? shortName : clientName;
                string title = $"{clientLabel} - {taskTypeValue} - {taskName}".Trim(' ', '-');

                string color = GetTaskCalendarColor(taskTypeValue, statusValue);

                foreach (var occurrence in ExpandTaskOccurrences(anchorDate, frequencyValue, rangeStart, rangeEnd))
                {
                    events.Add(new
                    {
                        id = $"{clientTaskId}_{occurrence:yyyyMMdd}",
                        title = title,
                        start = occurrence.ToString("yyyy-MM-dd"),
                        allDay = true,
                        backgroundColor = color,
                        borderColor = color,
                        extendedProps = new
                        {
                            clientTaskId,
                            clientId = rowClientId,
                            clientName,
                            shortName,
                            taskType = taskTypeValue,
                            taskName,
                            frequency = frequencyValue,
                            dueDay,
                            assignedTo = assignedToValue,
                            status = statusValue,
                            deliveryMethod,
                            notes,
                            sftpConfigId,
                            nextDueDate = anchorDate.ToString("MM/dd/yyyy"),
                            lastCompletedDate = lastCompletedDate?.ToString("MM/dd/yyyy") ?? ""
                        }
                    });
                }
            }

            return Json(events);
        }

        private static string GetTaskCalendarColor(string? taskType, string? status)
        {
            var type = (taskType ?? "").Trim().ToLowerInvariant();
            var stat = (status ?? "").Trim().ToLowerInvariant();

            if (stat == "overdue")
                return "#dc3545";

            if (stat == "completed")
                return "#198754";

            return type switch
            {
                "remit" => "#198754",
                "report" => "#0d6efd",
                "export" => "#fd7e14",
                "import" => "#6f42c1",
                "operations" => "#20c997",
                _ => "#6c757d"
            };
        }

    }
}
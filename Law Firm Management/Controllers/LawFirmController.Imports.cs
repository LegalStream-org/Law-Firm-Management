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
    public partial class LawFirmController
    {


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

                    // Names are unique across firms and agencies, so name-based imports stay unambiguous.
                    using (var dupCmd = new SqlCommand(@"
SELECT TOP 1 EntityType
FROM dbo.LawFirm
WHERE FirmName = @FirmName;", conn))
                    {
                        dupCmd.Parameters.AddWithValue("@FirmName", firmName);

                        var existingType = dupCmd.ExecuteScalar() as string;
                        if (existingType != null)
                        {
                            errors.Add(string.Equals(existingType, RecordType, StringComparison.OrdinalIgnoreCase)
                                ? $"Row {rowNumber}: '{firmName}' already exists."
                                : $"Row {rowNumber}: '{firmName}' already exists as {DescribeType(existingType)}.");
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
    @ProfileNotes,
    @RecordType,
    @CreatedBy,
    GETDATE(),
    @UpdatedBy,
    GETDATE()
);

SELECT CAST(SCOPE_IDENTITY() AS INT);", conn);

                    AddRecordTypeParam(cmd);

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

                    // Only attach rows to a record of this controller's type.
                    var (foundId, lookupError) = FindRecordIdByName(conn, firmName);
                    if (foundId == null)
                    {
                        errors.Add($"Row {rowNumber}: {lookupError}");
                        continue;
                    }
                    int lawFirmId = foundId.Value;

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

                    // Only attach rows to a record of this controller's type.
                    var (foundId, lookupError) = FindRecordIdByName(conn, firmName);
                    if (foundId == null)
                    {
                        errors.Add($"Row {rowNumber}: {lookupError}");
                        continue;
                    }
                    int lawFirmId = foundId.Value;

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
    }
}

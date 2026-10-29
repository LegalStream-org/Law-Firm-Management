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
    }
}

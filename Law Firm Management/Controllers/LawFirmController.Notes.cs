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
    }
}

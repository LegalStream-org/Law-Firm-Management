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
    }
}

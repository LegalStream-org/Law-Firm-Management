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
                    TempData["Error"] = $"That state and portfolio are already mapped for this {Labels.Short.ToLower()}.";
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
                    TempData["Error"] = $"That state and portfolio are already mapped for this {Labels.Short.ToLower()}.";
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
    }
}

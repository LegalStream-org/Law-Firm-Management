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

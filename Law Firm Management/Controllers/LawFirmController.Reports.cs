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
            // StateMap has no type of its own, so join to LawFirm to filter.
            using (var cmd = new SqlCommand(@"
        SELECT DISTINCT b.Portfolio
        FROM dbo.LawFirmStateMap b
        INNER JOIN dbo.LawFirm a
            ON a.LawFirmId = b.LawFirmId
        WHERE b.Portfolio IS NOT NULL
          AND a.EntityType = @RecordType
        ORDER BY b.Portfolio;", conn))
            {
                AddRecordTypeParam(cmd);

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
          AND a.EntityType = @RecordType
        ORDER BY b.StateCode, a.FirmName;", conn))
            {
                cmd.Parameters.AddWithValue("@Portfolio", portfolio ?? "");
                AddRecordTypeParam(cmd);

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
          AND a.EntityType = @RecordType
        ORDER BY b.StateCode, a.FirmName;", conn);

            cmd.Parameters.AddWithValue("@Portfolio", portfolio ?? "");
            AddRecordTypeParam(cmd);

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
    }
}

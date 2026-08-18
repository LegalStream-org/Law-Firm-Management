using Law_Firm_Management.Models;
using Law_Firm_Management.Service;
using Microsoft.AspNetCore.Mvc;
using ClosedXML.Excel;
using System.Data;

namespace Law_Firm_Management.Controllers
{
    public class FirmManagementController : Controller
    {
        private readonly IFirmManagementService _service;

        public FirmManagementController(IFirmManagementService service)
        {
            _service = service;
        }

        [HttpGet]
        public async Task<IActionResult> Index(string? state, string? status, string? search)
        {
            var model = await _service.GetFirmsAsync(state, status, search);

            ViewBag.State = state;
            ViewBag.Status = status;
            ViewBag.Search = search;

            return View(model);
        }

        [HttpGet]
        public IActionResult Create()
        {
            return View(new FirmNcbVm
            {
                Status = "Active"
            });
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Create(FirmNcbVm model)
        {
            if (!ModelState.IsValid)
                return View(model);

            await _service.CreateFirmAsync(model);
            TempData["Success"] = "Firm added successfully.";
            return RedirectToAction(nameof(Index));
        }

        [HttpGet]
        public async Task<IActionResult> Edit(int id)
        {
            var model = await _service.GetFirmByIdAsync(id);
            if (model == null)
                return NotFound();

            return View(model);
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Edit(FirmNcbVm model)
        {
            if (!ModelState.IsValid)
                return View(model);

            await _service.UpdateFirmAsync(model);
            TempData["Success"] = "Firm updated successfully.";
            return RedirectToAction(nameof(Index));
        }

        [HttpGet]
        public async Task<IActionResult> Export(string? state, string? status, string? search)
        {
            var dt = await _service.GetFirmsDataTableAsync(state, status, search);

            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Firms");

            ws.Cell(1, 1).InsertTable(dt);

            ws.Columns().AdjustToContents();

            var fileName = $"FirmManagement_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";

            using var stream = new MemoryStream();
            wb.SaveAs(stream);
            stream.Position = 0;

            return File(
                stream.ToArray(),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                fileName);
        }
    }
}
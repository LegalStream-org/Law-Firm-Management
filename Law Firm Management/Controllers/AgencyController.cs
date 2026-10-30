using Law_Firm_Management.Models;
using Microsoft.AspNetCore.Mvc;

namespace Law_Firm_Management.Controllers
{
    /// <summary>Agency pages: every LawFirm action and view, limited to EntityType = 'Agency'.</summary>
    public class AgencyController : LawFirmController
    {
        public AgencyController(IConfiguration configuration) : base(configuration) { }

        protected override string RecordType => LawFirmRecordType.Agency;

        /// <summary>Reuses Views/LawFirm, so there is one set of views for both types.</summary>
        [NonAction]
        public override ViewResult View(string? viewName, object? model)
        {
            viewName ??= ControllerContext.ActionDescriptor.ActionName;
            return base.View(viewName.StartsWith("~") ? viewName : $"~/Views/LawFirm/{viewName}.cshtml", model);
        }
    }
}

using Law_Firm_Management.Models;
using System.Data;

namespace Law_Firm_Management.Service
{
    public interface IFirmManagementService
    {
        Task<List<FirmNcbVm>> GetFirmsAsync(string? state = null, string? status = null, string? search = null);
        Task<FirmNcbVm?> GetFirmByIdAsync(int id);
        Task CreateFirmAsync(FirmNcbVm model);
        Task UpdateFirmAsync(FirmNcbVm model);

        Task<DataTable> GetFirmsDataTableAsync(string? state = null, string? status = null, string? search = null);
    }
}

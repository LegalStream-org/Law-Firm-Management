using System.ComponentModel.DataAnnotations;

namespace Law_Firm_Management.Models
{
    // Row shown on Settings > Client Types list
    public class ClientTypeListRowVm
    {
        public int ClientTypeId { get; set; }
        public string? TypeName { get; set; }
        public string? Description { get; set; }
        public bool SuppressCommissionByDefault { get; set; }
        public bool IsActive { get; set; }
        public int ClientCount { get; set; }
        public DateTime? UpdatedDate { get; set; }
    }

    // Create / Edit form for a single Client Type
    public class ClientTypeEditVm
    {
        public int ClientTypeId { get; set; }

        [Required(ErrorMessage = "Type name is required.")]
        [StringLength(100)]
        [Display(Name = "Type Name")]
        public string? TypeName { get; set; }

        [StringLength(255)]
        public string? Description { get; set; }

        [Display(Name = "Auto-deselect commissions for this type")]
        public bool SuppressCommissionByDefault { get; set; }

        public bool IsActive { get; set; } = true;
    }
}

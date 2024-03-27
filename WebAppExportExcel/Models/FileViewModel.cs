using System.ComponentModel;
using System.ComponentModel.DataAnnotations;

namespace WebAppExportExcel.Models
{
    public class FileViewModel
    {
        
        public FileViewModel() { }

        [DisplayName("Chọn file Ticktay")]
        [Required]
        public IFormFile filetick { get; set; }
        [Required]
        [DisplayName("Chọn file công")]
        public IFormFile filecong { get; set; }
    }
}

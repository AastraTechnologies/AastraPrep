using System.ComponentModel.DataAnnotations;

namespace AastraTimeSheet.Models
{
    public class Files
    {
        [Display(Name = "Employee Daily Updates")]
        public IFormFile EmployeeUpdatesFile { get; set; }

        [Display(Name = "Leave Records")]
        public IFormFile LeaveRecordsFile { get; set; }
    }
}

using AastraTimeSheet.Models;
using AastraTimeSheet.Services;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using System.ComponentModel;
using System.Diagnostics;

namespace AastraTimeSheet.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }

        public IActionResult Index()
        {
            return View();
        }

        public async Task<IActionResult> ProcessFile(Files path)
        {
            try
            {

                if (path.EmployeeUpdatesFile.Length > 0)
                {

                    ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                    var _employeeUpdatesFile = path.EmployeeUpdatesFile;
                    var _leaveRecordsFile = path.LeaveRecordsFile;
                    FileRead fileRead = new FileRead();
                    List<PresentData> empReadData = await fileRead.ReadEmpExcelData(_employeeUpdatesFile);
                    List<LeaveData> leaveReadData = (_leaveRecordsFile != null) ? await fileRead.ReadLeavExcelData(_leaveRecordsFile) : new List<LeaveData>();
                    FileWrite fileWrite = new FileWrite();
                    byte[] fileBytes = fileWrite.WriteExcelData(empReadData, leaveReadData);
                    return File(fileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "NewFile.xlsx");
                }
                else
                {
                    ModelState.AddModelError("", "Employee Updates File cannot be empty.");
                    return View("ErrorView", path);
                }
            }
            catch (Exception ex)
            {
                return View("ErrorView", ex.Message);
            }
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}

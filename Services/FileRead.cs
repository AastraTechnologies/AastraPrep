using AastraTimeSheet.Models;
using OfficeOpenXml;

namespace AastraTimeSheet.Services
{
    public class FileRead
    {
        public async Task<List<PresentData>> ReadEmpExcelData(IFormFile inputEmpFile)
        {
            var lstEmpData = new List<PresentData>();
            try
            {
                using (var memoryStream = new MemoryStream())
                {
                    await inputEmpFile.CopyToAsync(memoryStream);
                    using (var package = new ExcelPackage(memoryStream))
                    {
                        foreach (var worksheet in package.Workbook.Worksheets)
                        {
                            for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                            {
                                var date = (worksheet.Cells[row, 1] != null && !string.IsNullOrWhiteSpace(Convert.ToString(worksheet.Cells[row, 1]))) ? Convert.ToString(worksheet.Cells[row, 1].Text) : string.Empty;

                                if (!string.IsNullOrEmpty(date))
                                {
                                    if (!string.IsNullOrEmpty(date))
                                    {
                                        string StartTime = (worksheet.Cells[row, 10] != null && !string.IsNullOrWhiteSpace(Convert.ToString(worksheet.Cells[row, 10]))) ? Convert.ToString(worksheet.Cells[row, 10].Text) : string.Empty;
                                        string EndTime = (worksheet.Cells[row, 11] != null && !string.IsNullOrWhiteSpace(Convert.ToString(worksheet.Cells[row, 11]))) ? Convert.ToString(worksheet.Cells[row, 11].Text) : string.Empty;
                                        if (string.IsNullOrEmpty(StartTime) && string.IsNullOrEmpty(EndTime))
                                        {
                                            string timeRange = (worksheet.Cells[row, 12] != null && !string.IsNullOrWhiteSpace(Convert.ToString(worksheet.Cells[row, 12]))) ? Convert.ToString(worksheet.Cells[row, 12].Text) : string.Empty;
                                            if (!string.IsNullOrEmpty(timeRange))
                                            {
                                                string[] times = timeRange.Split('-');
                                                StartTime = times[0].Trim();
                                                EndTime = times[1].Trim();
                                            }
                                        }
                                        var record = new PresentData()
                                        {
                                            Date = date,
                                            Name = (worksheet.Cells[row, 7] != null && !string.IsNullOrWhiteSpace(Convert.ToString(worksheet.Cells[row, 7]))) ? Convert.ToString(worksheet.Cells[row, 7].Text) : string.Empty,
                                            StartTime = StartTime,
                                            EndTime = EndTime,
                                            Hours = (worksheet.Cells[row, 13] != null && !string.IsNullOrWhiteSpace(Convert.ToString(worksheet.Cells[row, 13]))) ? Convert.ToString(worksheet.Cells[row, 13].Text) : string.Empty,
                                            Project = (worksheet.Cells[row, 9] != null && !string.IsNullOrWhiteSpace(Convert.ToString(worksheet.Cells[row, 9]))) ? Convert.ToString(worksheet.Cells[row, 9].Text) : string.Empty,
                                            Remarks = (worksheet.Cells[row, 17] != null && !string.IsNullOrWhiteSpace(Convert.ToString(worksheet.Cells[row, 17]))) ? Convert.ToString(worksheet.Cells[row, 17].Text) : string.Empty
                                        };

                                        lstEmpData.Add(record);
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
                return new List<PresentData>();
            }
            return lstEmpData;
        }

        public async Task<List<LeaveData>> ReadLeavExcelData(IFormFile inputLeaveFile)
        {
            var lstLeaveData = new List<LeaveData>();
            try
            {
                using (var memoryStream = new MemoryStream())
                {
                    await inputLeaveFile.CopyToAsync(memoryStream);
                    using (var package = new ExcelPackage(memoryStream))
                    {
                        foreach (var worksheet in package.Workbook.Worksheets)
                        {
                            Console.WriteLine($"Reading data from worksheet: {worksheet.Name}");

                            for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                            {
                                var holliday = (worksheet.Cells[row, 1] != null && !string.IsNullOrWhiteSpace(Convert.ToString(worksheet.Cells[row, 1]))) ? Convert.ToString(worksheet.Cells[row, 1].Text) : string.Empty;
                                var leaveDate = (worksheet.Cells[row, 2] != null && !string.IsNullOrWhiteSpace(Convert.ToString(worksheet.Cells[row, 2]))) ? Convert.ToString(worksheet.Cells[row, 2].Text) : string.Empty;

                                if (!string.IsNullOrEmpty(holliday) && !string.IsNullOrEmpty(leaveDate))
                                {
                                    var record = new LeaveData()
                                    {
                                        Holliday = holliday,
                                        LeaveDate = leaveDate
                                    };

                                    lstLeaveData.Add(record);
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
                return new List<LeaveData>();
            }
            return lstLeaveData;
        }
    }
}

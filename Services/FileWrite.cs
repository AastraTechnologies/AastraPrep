using AastraTimeSheet.Models;
using OfficeOpenXml.Style;
using System.Globalization;

namespace AastraTimeSheet.Services
{
    public class FileWrite
    {
        static bool isFridayPrinted = false;
        public byte[] WriteExcelData(List<PresentData> empData, List<LeaveData> leaveData)
        {
            try
            {
                string[] formats = { "dd-MMM-yyyy", "MM/dd/yyyy", "yyyyMMdd", "yyyy-MM-dd", "dd/MM/yyyy", "dd/MMM/yyyy", "dd-MMM-yy", "MMM-dd-yyyy", "yyyy/MM/dd", "dd-MM-yyyy", "MM-dd-yyyy", "yyyy.MM.dd" };
                ExcelTemplate template = new ExcelTemplate();
                byte[] fileContents;
                var validEmpDict = empData
                    .Where(d => !string.IsNullOrEmpty(d.Date) && d.Date != "")
                    .ToDictionary(d => DateTime.ParseExact(d.Date, formats, CultureInfo.InvariantCulture));
                var validLeaveDict = leaveData
                    .Where(d => !string.IsNullOrEmpty(d.LeaveDate) && d.LeaveDate != "")
                    .ToDictionary(d => DateTime.ParseExact(d.LeaveDate, formats, CultureInfo.InvariantCulture));

                DateTime oneDate = validEmpDict.Keys.First();

                int year = oneDate.Year;
                int month = oneDate.Month;

                DateTime startDate = new DateTime(year, month, 1);
                DateTime endDate = new DateTime(year, month, DateTime.DaysInMonth(year, month));


                using (var package = template.CreateReportTemplate())
                {
                    var worksheet = package.Workbook.Worksheets[0];

                    int recordIndex = 6;

                    for (DateTime date = startDate; date <= endDate; date = date.AddDays(1))
                    {

                        worksheet.Cells[recordIndex, 1].Value = date.ToString("dd-MMM-yyyy"); // Date
                        worksheet.Cells[recordIndex, 2].Value = date.DayOfWeek.ToString(); // Day
                                                                                           // 
                        if (validEmpDict.ContainsKey(date))
                        {
                            var record = validEmpDict[date];
                            if (recordIndex == 6)
                            {
                                worksheet.Cells[4, 1].Value = $" Consultant's Name : {record.Name}";
                            }
                            worksheet.Cells[recordIndex, 3].Value = record.StartTime;
                            worksheet.Cells[recordIndex, 4].Value = record.EndTime;
                            worksheet.Cells[recordIndex, 5].Value = record.Hours;
                            worksheet.Cells[recordIndex, 6].Value = string.Empty;
                            worksheet.Cells[recordIndex, 7].Value = record.Project;
                            worksheet.Cells[recordIndex, 8].Value = record.Remarks;
                            isFridayPrinted = false;
                        }
                        else if (validLeaveDict.ContainsKey(date))
                        {
                            worksheet.Cells[recordIndex, 3, recordIndex, 8].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            worksheet.Cells[recordIndex, 3, recordIndex, 8].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                            worksheet.Cells[recordIndex, 3, recordIndex, 8].Style.Font.Size = 12;
                            worksheet.Cells[recordIndex, 3, recordIndex, 8].Merge = true;
                            worksheet.Cells[recordIndex, 3].Value = validLeaveDict[date].Holliday;
                        }
                        else if (date.DayOfWeek == DayOfWeek.Friday || date.DayOfWeek == DayOfWeek.Saturday)
                        {
                            if (date.DayOfWeek == DayOfWeek.Friday)
                            {
                                worksheet.Cells[recordIndex, 3, recordIndex, 8].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                worksheet.Cells[recordIndex, 3, recordIndex, 8].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                worksheet.Cells[recordIndex, 3, recordIndex, 8].Style.Font.Size = 12;
                                worksheet.Cells[recordIndex, 3, recordIndex, 8].Merge = true;
                                worksheet.Cells[recordIndex, 3].Value = "Holiday";
                                isFridayPrinted = true;

                            }
                            else if (date.DayOfWeek == DayOfWeek.Saturday)
                            {

                                if (isFridayPrinted)
                                {
                                    worksheet.Cells[recordIndex - 1, 3, recordIndex, 8].Merge = true;
                                    worksheet.Cells[recordIndex - 1, 3].Value = "Holiday";
                                    worksheet.Cells[recordIndex - 1, 3, recordIndex, 8].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                    worksheet.Cells[recordIndex - 1, 3, recordIndex, 8].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                    worksheet.Cells[recordIndex - 1, 3, recordIndex, 8].Style.Font.Size = 12;
                                }

                                else
                                {
                                    worksheet.Cells[recordIndex, 3, recordIndex, 8].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                    worksheet.Cells[recordIndex, 3, recordIndex, 8].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                    worksheet.Cells[recordIndex, 3, recordIndex, 8].Style.Font.Size = 12;
                                    worksheet.Cells[recordIndex, 3, recordIndex, 8].Merge = true;
                                    worksheet.Cells[recordIndex, 3].Value = "Holiday";
                                }

                            }
                        }
                        else
                        {
                            worksheet.Cells[recordIndex, 3].Value = "00:00";

                            worksheet.Cells[recordIndex, 4].Value = "00:00";
                        }

                        recordIndex++;
                    }

                    fileContents = package.GetAsByteArray();
                }
                return fileContents;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
                return new byte[0];
            }
        }
    }
}

using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace AastraTimeSheet.Services
{
    public class ExcelTemplate
    {
        public ExcelPackage CreateReportTemplate()
        {
            var package = new ExcelPackage();
            var worksheet = package.Workbook.Worksheets.Add("Report");

            worksheet.DefaultRowHeight = 15;
            worksheet.Cells.Style.Font.Name = "Cambria";

            worksheet.Cells[1, 1].Value = "AASTRA TIMESHEET";
            worksheet.Cells[1, 1, 2, 8].Merge = true;
            worksheet.Cells[1, 1, 2, 8].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[1, 1, 2, 8].Style.Font.Size = 22;
            using (ExcelRange Rng = worksheet.Cells[1, 1, 2, 8])
            {
                Rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                Rng.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(235, 241, 222));
                Rng.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            }

            string currentMonthAndYear = DateTime.Now.ToString("MMM yyyy");
            worksheet.Cells[3, 1].Value = currentMonthAndYear;
            worksheet.Cells[3, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[3, 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            worksheet.Cells[3, 1, 3, 8].Merge = true;
            worksheet.Row(3).Height = 20;


            worksheet.Cells[4, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[4, 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            worksheet.Cells[4, 1, 4, 8].Merge = true;
            worksheet.Cells[4, 1, 4, 8].Style.Font.Bold = true;
            worksheet.Row(4).Height = 22;
            worksheet.Row(5).Height = 22;


            //Header of table  
            worksheet.Cells[5, 1].Value = "Date";
            worksheet.Cells[5, 2].Value = "Day";
            worksheet.Cells[5, 3].Value = "Start Time";
            worksheet.Cells[5, 4].Value = "End Time";
            worksheet.Cells[5, 5].Value = "Hours";
            worksheet.Cells[5, 6].Value = "Staff Signature";
            worksheet.Cells[5, 7].Value = "Project - Component/BR";
            worksheet.Cells[5, 8].Value = "Remarks";
            worksheet.Cells[5, 1, 5, 8].Style.Font.Bold = true;
            using (ExcelRange Rng = worksheet.Cells[5, 1, 5, 8])
            {
                Rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                Rng.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(235, 241, 222));

            }





            for (int row = 5; row <= worksheet.Dimension.End.Row; row++)
            {
                for (int col = 1; col <= 6; col++)
                {
                    worksheet.Cells[row, col].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells[row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                }
            }
            worksheet.Row(5).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Row(5).Style.VerticalAlignment = ExcelVerticalAlignment.Center;

            // Column widths
            worksheet.Column(1).Width = 15;
            worksheet.Column(2).Width = 12;
            worksheet.Column(3).Width = 13;
            worksheet.Column(4).Width = 10;
            worksheet.Column(5).Width = 10;
            worksheet.Column(6).Width = 20;
            worksheet.Column(7).Width = 45;
            worksheet.Column(8).Width = 100;

            // Border styles
            worksheet.Cells[1, 1, 37, 8].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            worksheet.Cells[1, 1, 37, 8].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            worksheet.Cells[1, 1, 37, 8].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            worksheet.Cells[1, 1, 37, 8].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

            return package;
        }
    }
}

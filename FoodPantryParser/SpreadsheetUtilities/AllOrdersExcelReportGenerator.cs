using FoodPantryParser.SpreadsheetModels;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;


namespace FoodPantryParser.SpreadsheetUtilities
{
    public static class AllOrdersExcelReportGenerator
    {
        public static void CreateReport(string filePath, string sheetName, string reportTitle, List<Ordersheet> orderSheets)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add(sheetName);

                // Add title
                worksheet.Cells["A1"].Value = reportTitle;
                worksheet.Cells["A1:H1"].Merge = true;
                worksheet.Cells["A1"].Style.Font.Size = 20;
                worksheet.Cells["A1"].Style.Font.Bold = true;
                worksheet.Cells["A1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["A1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells["A1"].Style.Fill.BackgroundColor.SetColor(Color.LightGray);

                // Add data
                int row = 3;
                var properties = typeof(Order).GetProperties();
                //// Add data table header

                for (int i = 0; i < properties.Length; i++)
                {
                    worksheet.Cells[row, i + 1].Value = properties[i].Name;
                    worksheet.Cells[row, i + 1].Style.Font.Bold = true;
                    worksheet.Cells[row, i + 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    worksheet.Cells[row, i + 1].Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
                    worksheet.Cells[row, i + 1].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                }
                row++;
                foreach (Ordersheet sheet in orderSheets)
                {
                    foreach (var item in sheet.Orders)
                    {
                        for (int col = 0; col < properties.Length; col++)
                        {
                            var value = properties[col].GetValue(item);
                            worksheet.Cells[row, col + 1].Value = value;
                            worksheet.Cells[row, col + 1].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                            // Format numbers
                            if (value is decimal || value is double || value is float)
                            {
                                worksheet.Cells[row, col + 1].Style.Numberformat.Format = "#,##0.00";
                            }
                            // Format dates
                            else if (value is DateTime)
                            {
                                worksheet.Cells[row, col + 1].Style.Numberformat.Format = "mm/dd/yyyy";
                            }
                        }
                        row++;
                    }

                    // Auto-fit columns
                    worksheet.Cells.AutoFitColumns();
                }

                // Add timestamp
                worksheet.Cells[row + 2, 1].Value = $"Generated: {DateTime.Now}";
                worksheet.Cells[row + 2, 1].Style.Font.Italic = true;

                // Save the file
                var file = new FileInfo(filePath);

                if (file.Exists)
                {
                    throw new Exception($"Duplicate file encountered {file.Name}. Remove all files before proceeding.");
                }

                package.SaveAs(file);
            }
               
        }
    }
}

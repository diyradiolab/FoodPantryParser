using FoodPantryParser.SpreadsheetModels;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;


namespace FoodPantryParser.SpreadsheetUtilities
{
    public static class ByDateExcelReportGenerator
    {
        public static void CreateReport<T>(string filePath, string sheetName, List<T> data, string reportTitle, Ordersheet orderSheet)
        {

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

                // Add data table header
                var currentCol = 0;
                var properties = typeof(T).GetProperties();
                for (int i = 0; i < properties.Length + 2; i++) //+2 for vouchers and new
                {
                    if (i < properties.Length)
                    {
                        worksheet.Cells[row, i + 1].Value = properties[i].Name;
                        worksheet.Cells[row, i + 1].Style.Font.Bold = true;
                        worksheet.Cells[row, i + 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        worksheet.Cells[row, i + 1].Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
                        worksheet.Cells[row, i + 1].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    }            
                }
                row++;
                foreach (var item in data)
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

                // Add totals row if numeric columns exist
                AddTotalsRow(worksheet, row, properties, orderSheet);

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

        private static void AddTotalsRow(ExcelWorksheet worksheet, int row, System.Reflection.PropertyInfo[] properties, Ordersheet orderSheet = null)
        {
            bool hasTotals = false;
            for (int col = 0; col < properties.Length; col++)
            {
                var propertyType = properties[col].PropertyType;
                if (IsNumericType(propertyType))
                {
                    if (properties[col].Name == "AgencyNumber") continue; // Don't sum agency number
                    // Add sum formula
                    worksheet.Cells[row, col + 1].Formula = $"SUM({GetExcelColumnName(col + 1)}4:{GetExcelColumnName(col + 1)}{row - 1})";
                    worksheet.Cells[row, col + 1].Style.Font.Bold = true;
                    worksheet.Cells[row, col + 1].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells[row, col + 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    worksheet.Cells[row, col + 1].Style.Fill.BackgroundColor.SetColor(Color.LightGreen);
                    hasTotals = true;
                }
            }

            if (hasTotals)
            {
                worksheet.Cells[row, 1].Value = "Totals";
                worksheet.Cells[row, 1].Style.Font.Bold = true;
            }
        }

        private static bool IsNumericType(Type type)
        {
            return type == typeof(int) || type == typeof(long) || type == typeof(float) ||
                   type == typeof(double) || type == typeof(decimal) ||
                   type == typeof(int?) || type == typeof(long?) || type == typeof(float?) ||
                   type == typeof(double?) || type == typeof(decimal?);
        }

        private static string GetExcelColumnName(int columnNumber)
        {
            string columnName = "";
            while (columnNumber > 0)
            {
                int remainder = (columnNumber - 1) % 26;
                columnName = Convert.ToChar('A' + remainder) + columnName;
                columnNumber = (columnNumber - 1) / 26;
            }
            return columnName;
        }
    }
}

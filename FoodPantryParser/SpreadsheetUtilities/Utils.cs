using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FoodPantryParser.SpreadsheetUtilities
{
    public static class ExcelRowIterator
    {
        public static IEnumerable<Dictionary<string, object>> ReadRowsFromExcel(ExcelWorksheet worksheet, int startRow = 15, string worksheetName = null)
        {

            // Get the dimensions of the worksheet
            int endRow = worksheet.Dimension.End.Row;
            int endCol = worksheet.Dimension.End.Column;

            // Iterate through each row starting at startRow
            for (int row = startRow; row <= endRow; row++)
            {
                var rowData = new Dictionary<string, object>();

                // Check if row is completely empty
                bool isRowEmpty = true;
                for (int col = 1; col <= endCol; col++)
                {
                    var cellValue = worksheet.Cells[row, col].Value;
                    if (cellValue != null)
                    {
                        isRowEmpty = false;
                        break;
                    }
                }

                // Skip if row is empty
                if (isRowEmpty)
                    continue;

                // Add each cell in the row to dictionary
                for (int col = 1; col <= endCol; col++)
                {
                    var cellValue = worksheet.Cells[row, col].Value;
                    string columnLetter = GetExcelColumnName(col);
                    rowData[$"{columnLetter}"] = cellValue;
                }

                yield return rowData;
            }

        }

        // Helper method to convert column number to letter (1 = A, 2 = B, etc.)
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

        public static object GetCellValue(ExcelWorksheet worksheet, string cellReference)
        {
            ExcelAddress cellLocation = new ExcelAddress(cellReference);
            return worksheet.Cells[cellLocation.Start.Row, cellLocation.Start.Column].Value;
        }


    }
}

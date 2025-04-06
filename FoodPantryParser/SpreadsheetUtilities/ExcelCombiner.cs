using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using OfficeOpenXml;

namespace FoodPantryParser.SpreadsheetUtilities
{
    public class SharedProcessingState
    {
        public Dictionary<DateTime, Dictionary<string, double>> DateToSumOrdersValues { get; set; } =
            new Dictionary<DateTime, Dictionary<string, double>>();

        public Dictionary<int, double> ColumnTotals { get; set; } = new Dictionary<int, double>();

        public int NonZeroSumOrdersCount { get; set; } = 0;

        public HashSet<DateTime> ZeroDatesAcrossAllAgencies { get; set; } = new HashSet<DateTime>();
    }

    public class ExcelCombiner
    {
        private SharedProcessingState sharedState = null;
        private int totalsLabelRow = -1;
        private int totalsLabelColumn = -1;
        private ExcelWorksheet totalsWorksheet = null;

        private Dictionary<int, int> sourceToDestRowMap = new Dictionary<int, int>();
        private Dictionary<int, DateTime> rowDateMap = new Dictionary<int, DateTime>();

        public bool CombineSpreadsheets(
            string sourceFolderPath,
            string outputFilePath,
            string outputWorksheetName = "Combined Data",
            string searchPattern = "*.xlsx",
            bool includeXls = true,
            bool highlightZeroRows = true,
            bool showTotals = true,
            bool includeFileNameHeader = true,
            Action<string> logAction = null)
        {
            logAction = logAction ?? Console.WriteLine;
            this.sharedState = new SharedProcessingState();
            this.sourceToDestRowMap.Clear();
            this.rowDateMap.Clear();

            try
            {
                if (!Directory.Exists(sourceFolderPath))
                {
                    logAction("Folder does not exist. Please check the path and try again.");
                    return false;
                }

                var excelFiles = new List<string>(Directory.GetFiles(sourceFolderPath, searchPattern));
                if (excelFiles.Count == 0)
                {
                    logAction("No Excel files found in the specified folder.");
                    return false;
                }

                logAction($"Found {excelFiles.Count} Excel files.");

                using (var destinationPackage = new ExcelPackage(new FileInfo(outputFilePath)))
                {
                    var destWorksheet = destinationPackage.Workbook.Worksheets.Add(outputWorksheetName);
                    this.totalsWorksheet = destWorksheet;

                    int currentColumn = 1;

                    foreach (string filePath in excelFiles)
                    {
                        string fileName = Path.GetFileNameWithoutExtension(filePath);
                        logAction($"Processing: {fileName}...");

                        try
                        {
                            using (var sourcePackage = new ExcelPackage(new FileInfo(filePath)))
                            {
                                if (sourcePackage.Workbook.Worksheets.Count == 0)
                                {
                                    logAction($"  - No worksheets found in {fileName}, skipping");
                                    continue;
                                }

                                var sourceWorksheet = sourcePackage.Workbook.Worksheets[0];
                                logAction($"  - Reading worksheet: {sourceWorksheet.Name}");

                                if (sourceWorksheet.Dimension == null)
                                {
                                    logAction($"  - Worksheet is empty, skipping");
                                    continue;
                                }

                                int rows = sourceWorksheet.Dimension.Rows;
                                int cols = sourceWorksheet.Dimension.Columns;

                                int startRow = 1;
                                if (includeFileNameHeader)
                                {
                                    destWorksheet.Cells[1, currentColumn].Value = fileName;
                                    using (var range = destWorksheet.Cells[1, currentColumn, 1, currentColumn + cols - 1])
                                    {
                                        range.Merge = true;
                                        range.Style.Font.Bold = true;
                                        range.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                                    }
                                    startRow = 2;
                                }

                                int sourceStartRow = 4;

                                int dateColIndex = -1;
                                int sumOrdersColIndex = -1;
                                List<int> sumColumnIndices = new List<int>();

                                for (int c = 1; c <= cols; c++)
                                {
                                    string headerText = Convert.ToString(sourceWorksheet.Cells[sourceStartRow - 1, c].Value)?.Trim();
                                    if (headerText != null)
                                    {
                                        if (headerText.Equals("ReportDate", StringComparison.OrdinalIgnoreCase))
                                        {
                                            dateColIndex = c;
                                            logAction($"  - Identified Report Date column at position {c}");
                                        }
                                        else if (headerText.Equals("SumOrders", StringComparison.OrdinalIgnoreCase))
                                        {
                                            sumOrdersColIndex = c;
                                            sumColumnIndices.Add(c);
                                            logAction($"  - Identified SumOrders column at position {c}");
                                        }
                                        else if (headerText.StartsWith("Sum", StringComparison.OrdinalIgnoreCase))
                                        {
                                            sumColumnIndices.Add(c);
                                            logAction($"  - Identified Sum column at position {c}: {headerText}");
                                        }
                                    }
                                }

                                if (dateColIndex == -1 || sumOrdersColIndex == -1)
                                {
                                    logAction($"  - Warning: Could not find ReportDate ({dateColIndex}) or SumOrders ({sumOrdersColIndex}) columns");
                                }

                                int lastDataRow = rows;
                                for (int r = rows; r >= sourceStartRow; r--)
                                {
                                    for (int c = 1; c <= cols; c++)
                                    {
                                        string cellText = Convert.ToString(sourceWorksheet.Cells[r, c].Value);
                                        if (cellText != null && cellText.StartsWith("Generated:"))
                                        {
                                            lastDataRow = r - 1;
                                            logAction($"  - Found 'Generated:' text at row {r}, setting last data row to {lastDataRow}");
                                            break;
                                        }
                                    }
                                }

                                if (dateColIndex > 0 && sumOrdersColIndex > 0)
                                {
                                    for (int r = sourceStartRow; r <= lastDataRow; r++)
                                    {
                                        var dateValue = sourceWorksheet.Cells[r, dateColIndex].Value;
                                        var sumOrdersValue = sourceWorksheet.Cells[r, sumOrdersColIndex].Value;

                                        DateTime? date = null;
                                        if (dateValue != null)
                                        {
                                            if (dateValue is DateTime dt)
                                            {
                                                date = dt;
                                            }
                                            else if (double.TryParse(dateValue.ToString(), out double excelDate))
                                            {
                                                try
                                                {
                                                    date = DateTime.FromOADate(excelDate);
                                                    logAction($"  - Converted Excel date {excelDate} to {date.Value.ToShortDateString()} at row {r}");
                                                }
                                                catch
                                                {
                                                    logAction($"  - Warning: Failed to convert Excel date {excelDate} at row {r}");
                                                }
                                            }
                                            else if (DateTime.TryParse(dateValue.ToString(), out DateTime parsedDate))
                                            {
                                                date = parsedDate;
                                            }
                                        }

                                        if (date.HasValue)
                                        {
                                            double sumOrdersDouble = 0;
                                            if (sumOrdersValue != null && double.TryParse(sumOrdersValue.ToString(), out sumOrdersDouble))
                                            {
                                                // Successfully parsed
                                            }
                                            else
                                            {
                                                logAction($"  - Warning: SumOrders at row {r} is not a number: {sumOrdersValue}");
                                            }

                                            if (!this.sharedState.DateToSumOrdersValues.ContainsKey(date.Value))
                                            {
                                                this.sharedState.DateToSumOrdersValues[date.Value] = new Dictionary<string, double>();
                                            }
                                            this.sharedState.DateToSumOrdersValues[date.Value][fileName] = sumOrdersDouble;
                                            logAction($"  - Date: {date.Value.ToShortDateString()}, {fileName} SumOrders: {sumOrdersDouble}");

                                            if (sumOrdersDouble != 0)
                                            {
                                                this.sharedState.NonZeroSumOrdersCount++;
                                            }
                                        }
                                        else
                                        {
                                            logAction($"  - Warning: Invalid date at row {r}: {dateValue}");
                                        }
                                    }
                                }

                                Dictionary<int, double> columnTotals = new Dictionary<int, double>();
                                foreach (var sumColIdx in sumColumnIndices)
                                {
                                    columnTotals[sumColIdx] = 0;
                                }

                                for (int r = sourceStartRow; r <= lastDataRow; r++)
                                {
                                    foreach (var sumColIdx in sumColumnIndices)
                                    {
                                        var sumValue = sourceWorksheet.Cells[r, sumColIdx].Value;
                                        if (sumValue != null && double.TryParse(sumValue.ToString(), out double sumDouble))
                                        {
                                            columnTotals[sumColIdx] += sumDouble;
                                        }
                                    }
                                }

                                foreach (var entry in columnTotals)
                                {
                                    int colIdx = entry.Key;
                                    if (!this.sharedState.ColumnTotals.ContainsKey(colIdx))
                                    {
                                        this.sharedState.ColumnTotals[colIdx] = 0;
                                    }
                                    this.sharedState.ColumnTotals[colIdx] += entry.Value;
                                }

                                int destCol = currentColumn;
                                Dictionary<int, int> sourceToDestColMap = new Dictionary<int, int>();

                                for (int c = 1; c <= cols; c++)
                                {
                                    sourceToDestColMap[c] = destCol;

                                    var headerText = sourceWorksheet.Cells[sourceStartRow - 1, c].Value;
                                    destWorksheet.Cells[startRow, destCol].Value = headerText;
                                    destWorksheet.Cells[startRow, destCol].Style.Font.Bold = true;

                                    if (c == dateColIndex)
                                    {
                                        destWorksheet.Column(destCol).Style.Numberformat.Format = "mm/dd/yyyy";
                                    }

                                    for (int r = sourceStartRow; r <= lastDataRow; r++)
                                    {
                                        var cellValue = sourceWorksheet.Cells[r, c].Value;
                                        int destRow = startRow + (r - sourceStartRow) + 1;

                                        int uniqueRowKey = (currentColumn * 10000) + r;
                                        if (c == dateColIndex && !this.sourceToDestRowMap.ContainsKey(uniqueRowKey))
                                        {
                                            this.sourceToDestRowMap[uniqueRowKey] = destRow;
                                            if (cellValue != null)
                                            {
                                                if (cellValue is DateTime dt)
                                                {
                                                    this.rowDateMap[uniqueRowKey] = dt;
                                                }
                                                else if (double.TryParse(cellValue.ToString(), out double excelDate))
                                                {
                                                    try
                                                    {
                                                        this.rowDateMap[uniqueRowKey] = DateTime.FromOADate(excelDate);
                                                    }
                                                    catch
                                                    {
                                                        logAction($"  - Warning: Failed to map Excel date {excelDate} at row {r}");
                                                    }
                                                }
                                                else if (DateTime.TryParse(cellValue.ToString(), out DateTime dateValue))
                                                {
                                                    this.rowDateMap[uniqueRowKey] = dateValue;
                                                }
                                            }
                                        }

                                        if (c == dateColIndex && cellValue != null)
                                        {
                                            if (cellValue is DateTime dt)
                                            {
                                                destWorksheet.Cells[destRow, destCol].Value = dt;
                                            }
                                            else if (double.TryParse(cellValue.ToString(), out double excelDate))
                                            {
                                                try
                                                {
                                                    destWorksheet.Cells[destRow, destCol].Value = DateTime.FromOADate(excelDate);
                                                }
                                                catch
                                                {
                                                    destWorksheet.Cells[destRow, destCol].Value = cellValue;
                                                }
                                            }
                                            else if (DateTime.TryParse(cellValue.ToString(), out DateTime dateValue))
                                            {
                                                destWorksheet.Cells[destRow, destCol].Value = dateValue;
                                            }
                                            else
                                            {
                                                destWorksheet.Cells[destRow, destCol].Value = cellValue;
                                            }
                                            destWorksheet.Cells[destRow, destCol].Style.Numberformat.Format = "mm/dd/yyyy";
                                        }
                                        else
                                        {
                                            destWorksheet.Cells[destRow, destCol].Value = cellValue;
                                        }
                                    }
                                    destCol++;
                                }

                                int totalsRow = startRow + (lastDataRow - sourceStartRow) + 2;
                                if (currentColumn == 1)
                                {
                                    this.totalsLabelRow = totalsRow;
                                    this.totalsLabelColumn = currentColumn;
                                    destWorksheet.Cells[totalsRow, currentColumn].Value = "TOTALS";
                                    destWorksheet.Cells[totalsRow, currentColumn].Style.Font.Bold = true;
                                }

                                foreach (var entry in columnTotals)
                                {
                                    int sourceCol = entry.Key;
                                    if (sourceToDestColMap.ContainsKey(sourceCol))
                                    {
                                        int destTotalCol = sourceToDestColMap[sourceCol];
                                        destWorksheet.Cells[totalsRow, destTotalCol].Value = entry.Value;
                                        destWorksheet.Cells[totalsRow, destTotalCol].Style.Font.Bold = true;
                                    }
                                }

                                int lastCol = destCol - 1;
                                var borderRange = destWorksheet.Cells[startRow, lastCol, totalsRow, lastCol];
                                borderRange.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;

                                currentColumn += (destCol - currentColumn) + 1;

                                logAction($"  - Added {lastDataRow - sourceStartRow + 1} rows and {destCol - currentColumn} columns from {fileName}");
                            }
                        }
                        catch (Exception ex)
                        {
                            logAction($"Error processing file {fileName}: {ex.Message}");
                        }
                    }

                    // Calculate zero-order days
                    var allZeroDates = new HashSet<DateTime>();
                    var allAgencies = excelFiles.Select(f => Path.GetFileNameWithoutExtension(f)).ToHashSet();
                    logAction($"All agencies: {string.Join(", ", allAgencies)}");

                    foreach (var dateEntry in this.sharedState.DateToSumOrdersValues)
                    {
                        DateTime date = dateEntry.Key.Date;
                        var agencyOrders = dateEntry.Value;

                        logAction($"Checking date {date.ToShortDateString()}:");
                        foreach (var agency in allAgencies)
                        {
                            if (agencyOrders.ContainsKey(agency))
                            {
                                logAction($"  - {agency}: {agencyOrders[agency]}");
                            }
                            else
                            {
                                logAction($"  - {agency}: No data");
                            }
                        }

                        bool allZero = allAgencies.All(agency =>
                            agencyOrders.ContainsKey(agency) && agencyOrders[agency] == 0);

                        if (allZero)
                        {
                            allZeroDates.Add(date);
                            logAction($"  => Found zero date: {date.ToShortDateString()}");
                        }
                    }

                    this.sharedState.ZeroDatesAcrossAllAgencies = allZeroDates;
                    logAction($"Total zero days: {allZeroDates.Count}");

                    // Update totals label
                    if (this.totalsLabelRow > 0 && this.totalsLabelColumn > 0 && this.totalsWorksheet != null)
                    {
                        int zeroDaysCount = this.sharedState.ZeroDatesAcrossAllAgencies.Count;
                        logAction($"Updating totals label with zero days count: {zeroDaysCount}");
                        this.totalsWorksheet.Cells[this.totalsLabelRow, this.totalsLabelColumn].Value =
                            $"TOTALS ({zeroDaysCount})";
                    }
                    else
                    {
                        logAction("Warning: Totals label position not set properly.");
                    }

                    // Highlight zero-order rows
                    if (highlightZeroRows)
                    {
                        logAction("Highlighting zero-order rows:");
                        foreach (var rowEntry in this.rowDateMap)
                        {
                            int uniqueRowKey = rowEntry.Key;
                            DateTime rowDate = rowEntry.Value;
                            int destRow = this.sourceToDestRowMap[uniqueRowKey];

                            logAction($"  - Row key {uniqueRowKey}, Date {rowDate.ToShortDateString()}, DestRow {destRow}");
                            if (this.sharedState.ZeroDatesAcrossAllAgencies.Contains(rowDate))
                            {
                                logAction($"    => Highlighting row {destRow}");
                                for (int c = 1; c <= destWorksheet.Dimension.Columns; c++)
                                {
                                    destWorksheet.Cells[destRow, c].Style.Fill.PatternType =
                                        OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                    destWorksheet.Cells[destRow, c].Style.Fill.BackgroundColor.SetColor(
                                        System.Drawing.Color.LightPink);
                                }
                            }
                        }
                    }

                    destWorksheet.Cells.AutoFitColumns();
                    destinationPackage.Save();
                }

                logAction($"Successfully combined spreadsheets horizontally into: {outputFilePath}");
                return true;
            }
            catch (Exception ex)
            {
                logAction($"An error occurred: {ex.Message}");
                return false;
            }
        }
    }
}
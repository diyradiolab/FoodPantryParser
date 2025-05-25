using FoodPantryParser.ReportModels;
using FoodPantryParser.SpreadsheetModels;
using FoodPantryParser.SpreadsheetUtilities;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace FoodPantryParser
{
    public class ReportGenerator
    {
        public string DataFolder { get; set; }
        public string OutputFolder { get; set; }
        public string ByDateOutputFolder { get; set; }
        public string ByAgencyOutputFolder { get; set; }
        public List<DateTime> CurrentDates { get; set; }
        public string[] Files { get; set; }
        public int CurrentMonth { get; set; }
        public int CurrentYear { get; set; }

        public int InvalidRowsBetweenOrders { get; set; } // After two rows are encountered with no data, we know there are no more entries. 

        public ReportGenerator(string dataFolder, string outputFolder, int invalidRowsBetweenOrders, int currentMonth, int currentYear)
        {
            DataFolder = dataFolder;
            OutputFolder = outputFolder;
            InvalidRowsBetweenOrders = invalidRowsBetweenOrders;
            ByDateOutputFolder = OutputFolder + @"\ByDate";
            ByAgencyOutputFolder = OutputFolder + @"\ByAgency";
            Files = Directory.GetFiles(DataFolder, "*.xls", SearchOption.TopDirectoryOnly);
            CurrentDates = WeekdayCalendar.GetWeekdaysInMonth(currentMonth, currentYear);
            CurrentMonth = currentMonth;
            CurrentYear = currentYear;
        }

        public void Execute()
        {
            ManageDirectories();
            try
            {
                CreateAllReports();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
                throw ex;
            }


        }

        private void GenerateCalendarReport()
        {
            var monthStr = new DateTime(CurrentYear, CurrentMonth, 1, 0, 0, 0, DateTimeKind.Local).ToString("MMMM");

            GenericExcelReportGenerator.CreateReport(Path.Combine(OutputFolder, $"{monthStr}{CurrentYear.ToString()}.xlsx"), "Weekdays", CurrentDates, $"Weekdays Report");
        }

        private void CreateAllReports()
        {
            // Create weekday report
            Console.Clear();
            Console.WriteLine("Creating Weekday report...");
            Console.WriteLine("-----------------------------------");
            GenerateCalendarReport();
            Console.WriteLine();
            GenerateAllAgencyReportData();
            Console.WriteLine("Creating Monthly reports for each agency...");
            Console.WriteLine("-----------------------------------");
            GenerateByAgencyReportData();
            Console.WriteLine("Combining Monthly reports for each agency...");
            Console.WriteLine("-----------------------------------");
            CombineByAgencyReportData();
            Console.WriteLine();
            Console.WriteLine("Creating daily total reports...");
            Console.WriteLine("-----------------------------------");
            GenerateByDateReportData();
            Console.WriteLine();
            Console.WriteLine("FINISHED!");
        }

        private void CombineByAgencyReportData()
        {
            // Call the combination function
            var combiner = new ExcelCombiner();
            combiner.CombineSpreadsheets(ByAgencyOutputFolder, OutputFolder + @"\March2025TabbedAgencyReports.xlsx");
        }

        private void ManageDirectories()
        {
            if (Directory.Exists(ByDateOutputFolder) || Directory.Exists(ByAgencyOutputFolder))
            {
                Console.WriteLine("Remove ByDate and ByAgency folders before proceeding or else error will be encountered during processing");
                Console.WriteLine("If you encounter an error even after removing the files, that means you have two files with the same Date");
                Console.WriteLine("Press any key to continue after removing these folders");
                Console.ReadKey();
            }

            if (!Directory.Exists(ByDateOutputFolder))
            {
                Directory.CreateDirectory(ByDateOutputFolder);
                Console.WriteLine("Created FPReport\\ByDate folder on desktop");
            }

            if (!Directory.Exists(ByAgencyOutputFolder))
            {
                Directory.CreateDirectory(ByAgencyOutputFolder);
                Console.WriteLine("Created FPReport\\ByAgency folder on desktop");
            }
        }

        private bool ShouldStopProcessing(int consecutiveSkips) => consecutiveSkips > InvalidRowsBetweenOrders;

        private bool IsSkippableRow(IDictionary<string, object> row)
            => row["A"] == null ||
               row["A"].ToString().StartsWith("I acknowledge receipt");

        private Order CreateOrderFromRow(IDictionary<string, object> row, ExcelWorksheet worksheet)
        {
            var order = new Order();

            order.OrderDate = (DateTime)ExcelRowIterator.GetCellValue(worksheet, "B12");
            order.AgencyNumber = long.Parse(row["F"].ToString());
            order.AgencyName = row["B"].ToString();
            order.Adults = Convert.ToInt32(row["L"]);
            order.Children = Convert.ToInt32(row["M"]);
            if (row["N"] != null)
            {
                order.HasVoucher = true;
            }
            if (row["O"] != null)
            {
                order.IsNewClient = true;
            }
            if (row["P"] != null)
            {
                order.IsCity = Geography.IsCity(row["P"].ToString());
            }

            return order;
        }

        private void GenerateByAgencyReportData()
        {
            try
            {
                var orders = new List<Order>();
                foreach (var file in Files)
                {
                    using (var package = new ExcelPackage(file))
                    {
                        //get the first worksheet in the workbook
                        ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

                        int consecutiveSkips = 0;
                        foreach (var row in SpreadsheetUtilities.ExcelRowIterator.ReadRowsFromExcel(worksheet))
                        {
                            //if two rows (InvalidRowsBetweenOrders) have been skipped, we're done.
                            if (ShouldStopProcessing(consecutiveSkips))
                                break;

                            if (IsSkippableRow(row))
                            {
                                consecutiveSkips++;
                                continue;
                            }

                            orders.Add(CreateOrderFromRow(row, worksheet));
                            consecutiveSkips = 0; //reset skipped
                        }
                    }
                }
                CreateByAgencyReport(orders);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                throw;
            }
        }

        private void CreateByAgencyReport(List<Order> allOrders)
        {
            try
            {
                // Each Line need to be Date | Orders | SumAdults | SumChildren
                var ordersByAgency = allOrders
                    .GroupBy(x => x.AgencyName)
                    .OrderBy(x => x.Key)  // Key is AgencyName in this case
                    .Select(grp => grp.ToList())
                    .ToList();
                foreach (var orderGroup in ordersByAgency)
                {
                    //Creates a Dictionary where dates are keys and the values are lists of orders for that date. It's like sorting all orders into date-labeled containers first.
                    var ordersByDate = orderGroup.GroupBy(x => x.OrderDate)
                                                .ToDictionary(g => g.Key, g => g.ToList());

                    //For each date:
                    //Declare a variable to hold orders for this date
                    //TryGetValue attempts to get orders for this date from our dictionary
                    //If found, puts them in dateOrders
                    //If not found, dateOrders will be null
                    // Create a new list to store all reports
                    List<DailyReport> reports = new List<DailyReport>();

                    // Iterate through each date
                    foreach (var date in CurrentDates)
                    {
                        // Try to get orders for this date
                        List<Order> dateOrders;
                        ordersByDate.TryGetValue(date, out dateOrders);

                        // Initialize variables for calculations
                        int sumAdults = 0;
                        int sumChildren = 0;
                        int sumOrders = 0;
                        int sumVouchers = 0;
                        int sumNewClients = 0;

                        // Only perform calculations if we have orders for this date
                        if (dateOrders != null)
                        {
                            // Calculate sums by iterating through all orders
                            foreach (Order order in dateOrders)
                            {
                                sumAdults += order.Adults;
                                sumChildren += order.Children;
                                if (order.HasVoucher)
                                {
                                    sumVouchers += 1;
                                }
                                if (order.IsNewClient)
                                {
                                    sumNewClients += 1;
                                }
                            }
                            sumOrders = dateOrders.Count;
                        }

                        // Create the report and add it to our list
                        DailyReport report = new DailyReport(date)
                        {
                            AgencyName = orderGroup[0].AgencyName,
                            AgencyNumber = orderGroup[0].AgencyNumber,
                            SumAdults = sumAdults,
                            SumChildren = sumChildren,
                            SumOrders = sumOrders,
                            SumVouchers = sumVouchers,
                            SumNewClients = sumNewClients
                        };

                        reports.Add(report);
                    }
                    var fileName = reports[0].AgencyName;
                    if(HasIllegalCharacters(fileName))
                    {
                        throw new Exception($"Agency Name: {fileName} has illegal characters");
                    }
                    GenericExcelReportGenerator.CreateReport(Path.Combine(ByAgencyOutputFolder, $"{fileName}" + ".xlsx"), "Report", reports, $"Monthly Agency Report - {reports[0].AgencyName} - {CurrentMonth}/{CurrentYear}");
                    Console.WriteLine(reports[0].AgencyName);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                throw;
            }
        }

        public static bool HasIllegalCharacters(string filename)
        {
            char[] invalidChars = Path.GetInvalidFileNameChars();
            foreach (char c in filename)
            {
                if (Array.Exists(invalidChars, element => element == c))
                {
                    return true;
                }
            }
            return false;
        }

        private void GenerateByDateReportData()
        {
            try
            {
                var totalVouchers = 0;
                var totalNewClients = 0;
                var allOrders = new List<Order>();
                foreach (var file in Files)
                {
                    var orders = new List<Order>();
                    Ordersheet orderSheet;
                    using (var package = new ExcelPackage(file))
                    {
                        //get the first worksheet in the workbook
                        orderSheet = new Ordersheet();
                        ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                        orderSheet.OrderDate = (DateTime)ExcelRowIterator.GetCellValue(worksheet, "B12");
                        orderSheet.NewClients = Convert.ToInt32(ExcelRowIterator.GetCellValue(worksheet, "C10"));
                        totalNewClients += orderSheet.NewClients;
                        orderSheet.Vouchers = Convert.ToInt32(ExcelRowIterator.GetCellValue(worksheet, "E10"));
                        totalVouchers += orderSheet.Vouchers;
                        var consecutiveSkips = 0;
                        foreach (var row in SpreadsheetUtilities.ExcelRowIterator.ReadRowsFromExcel(worksheet))
                        {
                            //if two rows (InvalidRowsBetweenOrders) have been skipped, we're done.
                            if (ShouldStopProcessing(consecutiveSkips))
                                break;

                            if (IsSkippableRow(row))
                            {
                                consecutiveSkips++;
                                continue;
                            }

                            var order = CreateOrderFromRow(row, worksheet);
                            orders.Add(order);
                            allOrders.Add(order);

                            consecutiveSkips = 0; //reset skipped
                        }
                    }
                    CreateByDateReport(orders, orderSheet); //Todo: Include New and Vouchers in Daily Report.
                }

                var summaryInfo = WriteSummaryInfoReport(totalVouchers, totalNewClients, allOrders);
                WriteSummaryInfoToConsole(summaryInfo);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                throw;
            }
        }

        private void CreateByDateReport(List<Order> orders, Ordersheet orderSheet)
        {
            try
            {
                var orderDate = orders[0].OrderDate;
                var orderDateStr = orders[0].OrderDate.ToString("yyyyMMdd");
                var reports = new List<DailyReport>();
                foreach (var order in orders)
                {
                    // look for current report
                    var aReport = reports.FirstOrDefault(x => x.AgencyName == order.AgencyName);

                    // if no report found in list, add a new one
                    if (aReport == null)
                    {
                        var r = new DailyReport(orderDate)
                        {
                            AgencyName = order.AgencyName,
                            AgencyNumber = order.AgencyNumber,
                            ReportDate = order.OrderDate,
                            SumChildren = order.Children,
                            SumAdults = order.Adults,
                            SumOrders = 1,
                            SumVouchers = order.HasVoucher ? 1 : 0,
                            SumNewClients = order.IsNewClient ? 1 : 0

                        };
                        reports.Add(r);
                        continue;
                    }

                    //if existing report is already in the list, add to the sums
                    aReport.SumChildren += order.Children;
                    aReport.SumAdults += order.Adults;
                    aReport.SumOrders++;
                    aReport.SumVouchers += order.HasVoucher ? 1 : 0;
                    aReport.SumNewClients += order.IsNewClient ? 1 : 0;

                }
                //write that report for the day

                ByDateExcelReportGenerator.CreateReport(Path.Combine(ByDateOutputFolder, orderDateStr + ".xlsx"), "Report", reports, $"Daily Report Totals ({orderDate.ToString("dddd, MMMM dd, yyyy")})", orderSheet);


                Console.WriteLine(orderDate.ToString("dddd, MMMM dd, yyyy"));

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                throw;
            }
        }

        private void WriteSummaryInfoToConsole(StringBuilder summaryInfo)
        {
            using (StringReader reader = new StringReader(summaryInfo.ToString()))
            {
                Console.WriteLine();
                string line;
                while ((line = reader.ReadLine()) != null)
                {
                    Console.WriteLine(line);
                }
            }
        }

        private StringBuilder WriteSummaryInfoReport(int totalVouchers, int totalNewClients, List<Order> allOrders)
        {
            var sb = new StringBuilder();
            sb.AppendLine("Number of order forms processed: " + Files.Length);
            sb.AppendLine("Total Vouchers: " + totalVouchers);
            sb.AppendLine("Total New Clients: " + totalNewClients);
            sb.AppendLine("Total Orders: " + allOrders.Count);
            sb.AppendLine("Total Persons: " + allOrders.Sum(x=>x.Adults + x.Children));
            sb.AppendLine($"Total City Orders: {allOrders.Count(x => x.IsCity)}");
            sb.AppendLine($"Total City Persons: {allOrders.Where(x => x.IsCity).Sum(x => x.Adults + x.Children)}");

            sb.AppendLine();

            var missingDates = FindMissingDates(allOrders, CurrentDates);

            sb.AppendLine("Dates without any orders");
            foreach (var missingDate in missingDates)
            {
                sb.AppendLine($"{missingDate.ToString("dddd, MMMM dd, yyyy")}");
            }

            var infoFile = new FileInfo(Path.Combine(OutputFolder, "info.txt"));
            System.IO.File.WriteAllText(infoFile.FullName, sb.ToString());
            return sb;
        }

        private List<DateTime> FindMissingDates(List<Order> items, List<DateTime> allDates)
        {
            // Get all dates from items
            var itemDates = items.Select(item => item.OrderDate.Date).ToList();  // .Date removes time component

            // Find dates that don't exist in itemDates
            var missingDates = allDates
                .Where(date => !itemDates.Contains(date.Date))
                .ToList();

            return missingDates;
        }

        private void GenerateAllAgencyReportData()
        {
            var orderSheets = new List<Ordersheet>();
            foreach (var file in Files)
            {
                Ordersheet orderSheet = null;
                try
                {
                    using (var package = new ExcelPackage(file))
                    {
                        //get the first worksheet in the workbook
                        orderSheet = new Ordersheet();
                        ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                        orderSheet.OrderDate = (DateTime)ExcelRowIterator.GetCellValue(worksheet, "B12");
                        //orderSheet.NewClients = Convert.ToInt32(ExcelRowIterator.GetCellValue(worksheet, "C10"));
                        //orderSheet.Vouchers = Convert.ToInt32(ExcelRowIterator.GetCellValue(worksheet, "E10"));
                        var consecutiveSkips = 0;
                        foreach (var row in SpreadsheetUtilities.ExcelRowIterator.ReadRowsFromExcel(worksheet))
                        {
                            //if two rows (InvalidRowsBetweenOrders) have been skipped, we're done.
                            if (ShouldStopProcessing(consecutiveSkips))
                                break;

                            if (IsSkippableRow(row))
                            {
                                consecutiveSkips++;
                                continue;
                            }

                            var order = CreateOrderFromRow(row, worksheet);
                            orderSheet.Orders.Add(order);
                            consecutiveSkips = 0; //reset skipped
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error in file " + file + ". Possibly missing date. " + "\r\n" + ex.InnerException);
                    Console.ReadLine();
                }

                orderSheets.Add(orderSheet);
            }
            AllOrdersExcelReportGenerator.CreateReport(Path.Combine(OutputFolder, "all" + ".xlsx"), "Report", "All Orders", orderSheets);
        }
    }
}

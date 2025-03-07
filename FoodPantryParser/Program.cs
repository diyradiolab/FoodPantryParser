using FoodPantryParser.ReportModels;
using OfficeOpenXml;
using System;
using System.Diagnostics;
using System.IO;
using System.Text;


namespace FoodPantryParser
{
    internal static class Program
    {
        public static string DataFolder => @"C:\Users\Alan Hess\Desktop\FoodPantryData";
        public static string OutputFolder => $@"{Environment.GetFolderPath(Environment.SpecialFolder.Desktop)}\FPReport";
        internal const int InvalidRowsBetweenOrders = 10; // After two rows are encountered with no data, we know there are no more entries. 1 OR 2 is common. 10 is generous.

        static void Main(string[] args)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            bool showMenu = true;
            while (showMenu)
            {
                showMenu = MainMenu();
            }
            Shutdown(4, 350);
        }

        public static void Shutdown(int ticks, int interval)
        {
            StringBuilder sbDots = new StringBuilder();
            for (int i = 0; i < ticks; i++)
            {
                sbDots.Append(".");
                Console.Clear();
                Console.WriteLine("Shutting down" + sbDots.ToString());
                System.Threading.Thread.Sleep(interval);
            }

            Console.Clear();
            Console.WriteLine("Bye!");
            System.Threading.Thread.Sleep(800);
        }

        private static bool MainMenu()
        {
            Console.Clear();
            Console.WriteLine("Choose an option:");
            Console.WriteLine("1) Show README");
            Console.WriteLine("2) Create Calendar, By Agency and By Date Reports");
            Console.WriteLine("3) Exit");
            Console.Write("\r\nSelect an option: ");

            switch (Console.ReadLine())
            {
                case "1":
                    ShowReadme();
                    return true;
                case "2":
                    try
                    {
                        (int month, int year) = GetCurrentPeriodFromUser();
                        CreateReports(month, year);
                        return true;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Error: " + ex.Message);
                        Console.ReadLine();
                        MainMenu();
                    }
                    return false;
                    
                case "3":
                    return false;
                default:
                    Console.Clear();
                    Console.WriteLine("Invalid entry. Press any key to return to menu");
                    Console.ReadKey();
                    return true;
            }
        }
        private static (int, int) GetCurrentPeriodFromUser()
        {
            Console.WriteLine("Enter reporting month (1,2,3, etc) ");
            var month = int.Parse(Console.ReadLine());
            Console.WriteLine("Enter reporting year (2024, 2025, etc)");
            var year = int.Parse(Console.ReadLine());
            Console.WriteLine($"You entered: {new DateTime(year, month, 1, 0, 0, 0, DateTimeKind.Local).ToString("MMMM, yyyy")}");
            Console.WriteLine("Is this correct: y/n");
            ConsoleKeyInfo keyInfo = Console.ReadKey();
            char keyPressed = char.ToLower(keyInfo.KeyChar);
            
            switch (keyPressed)
            {
               
                case 'y':
                    return (month, year);
                case 'n':
                    Console.WriteLine("bad");
                    break;
                default:
                    break;
            }
            throw new Exception("Problem with values entered.");
        }
        private static void ShowReadme()
        {
            Console.Clear();
            Console.WriteLine("=== README ===");
            Console.WriteLine("Alan Hess, Franklin County Emergency Food Pantry, 2024");
            Console.WriteLine();
            Console.WriteLine($"To begin, create a folder on your desktop: {DataFolder} and");
            Console.WriteLine("copy all daily Excel reports here.");
            Console.WriteLine();
            Console.WriteLine($"Next, create a subfolder {DataFolder}\\input. Create a text file");
            Console.WriteLine("\"currentDates.txt\" and put a list of all dates for the month, separating dates by a carriage return:");
            Console.WriteLine("11/1/24");
            Console.WriteLine("11/2/24");
            Console.WriteLine("11/5/24");
            Console.WriteLine("...");
            Console.WriteLine();
            Console.WriteLine($"After running the program, reports will be placed here: {OutputFolder}");
            Console.WriteLine();
            Console.WriteLine("A spreadsheet will be availble in this folder named November2024 or similary");
            Console.WriteLine();
            Console.WriteLine("Reports are overwritten, but not deleted. It is advisable to delete the output folders each time the program executed.");
            Console.WriteLine();
            Console.WriteLine("Check the console output for duplicate names (East Frankfort and East Frankfort Baptist, and correct the");
            Console.WriteLine("data files to have the same Agency Name throughout. The program doesn't use the Agency Number for anything,");
            Console.WriteLine("but the Agency Name must match.");
            Console.WriteLine();
            Console.WriteLine("To add or edit an agency, modify the agency list in Initializer.cs and recompile using Visual Studio.");
            Console.WriteLine("\r\nPress any key to return to menu");
            Console.ReadKey();
        }

        private static void CreateReports(int month, int year)
        {
            Console.Clear();
            var reportGenerator = new ReportGenerator(DataFolder, OutputFolder, InvalidRowsBetweenOrders, month, year);
            reportGenerator.Execute();
            Console.WriteLine("\r\nPress any key to return to menu");
            Console.ReadKey();
        }
    }

}

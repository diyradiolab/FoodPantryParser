using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FoodPantryParser.ReportModels
{
    public class WeekdayCalendar
    {
        public static List<DateTime> GetWeekdaysInMonth(int month, int year )
        {


            // Get first day of the month
            var startDate = new DateTime(year, month, 1);

            // Get last day of the month
            var endDate = startDate.AddMonths(1).AddDays(-1);

            // Generate all dates in the month and filter for weekdays
            var weekdays = Enumerable.Range(0, endDate.Day)
                .Select(day => startDate.AddDays(day))
                .Where(date => date.DayOfWeek != DayOfWeek.Saturday &&
                              date.DayOfWeek != DayOfWeek.Sunday)
                .ToList();

            return weekdays;
        }

        public static List<string> FormatWeekdays(List<DateTime> dates, string format = "dddd, MMMM dd, yyyy")
        {
            return dates.Select(date => date.ToString(format)).ToList();
        }

    }
  }
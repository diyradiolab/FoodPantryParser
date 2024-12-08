using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FoodPantryParser.ReportModels
{
    internal class DailyReport
    {
        public long AgencyNumber { get; set; }
        public string AgencyName { get; set; }
        public DateTime ReportDate { get; set; }
        public int SumOrders { get; set; } = 0;
        public int SumAdults { get; set; } = 0;
        public int SumChildren { get; set; } = 0;
        public int SumAdultsChildren => SumAdults + SumChildren;
        public int SumVouchers { get; set; } = 0;
        public int SumNewClients { get; set; } = 0;


        public DailyReport(DateTime date) {
            ReportDate = date;
        }
    }
}

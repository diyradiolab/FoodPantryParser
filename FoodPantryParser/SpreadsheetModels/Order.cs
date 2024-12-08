using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FoodPantryParser.SpreadsheetModels
{
    public class Order
    {
        public DateTime OrderDate { get; set; }
        public long AgencyNumber { get; set; }
        public string AgencyName { get; set; }
        public int Adults { get; set; }
        public int Children { get; set; }
        public bool HasVoucher { get; set; }
        public bool IsNewClient { get; set; }
    }
}

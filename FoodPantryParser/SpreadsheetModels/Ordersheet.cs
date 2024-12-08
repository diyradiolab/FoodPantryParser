using System;
using System.Collections.Generic;

namespace FoodPantryParser.SpreadsheetModels
{
    public class Ordersheet
    {
        public DateTime OrderDate { get; set; }
        public int NewClients { get; set; }
        public int Vouchers { get; set; }
        public List<Order> Orders { get; set; } = new List<Order>();
    }
}

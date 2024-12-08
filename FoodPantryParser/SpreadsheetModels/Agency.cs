using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FoodPantryParser.SpreadsheetModels
{
    public class Agency
    {
        public long AgencyNumber { get; set; }
        public string AgencyName { get; set; }
        public DateTime DateAdded { get; set; }
        public bool IsActive { get; set; }
    }
}

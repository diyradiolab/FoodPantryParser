using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FoodPantryParser.SpreadsheetModels
{
    public static class Geography
    {

        public static bool IsCity(string location) 
        {
            if(string.IsNullOrWhiteSpace(location)) throw new ArgumentNullException("location");

            if (location.ToLower() != "city" && location.ToLower() != "county") throw new ArgumentException("location");

            return (location.ToLower() == "city");
          
        }

    }
}

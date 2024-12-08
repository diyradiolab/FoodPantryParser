using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FoodPantryParser.ReportModels
{
    public  interface IReportInfo
    {
        string DataFolder { get; set; }
        string OutputFolder { get; set; }
    }
}

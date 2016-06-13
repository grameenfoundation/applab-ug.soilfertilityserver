using System;
using System.Collections.Generic;
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;
using Optimize;

namespace Grameen.Models
{
    public class ActivityView
    { 
        public DateTime Date { get; set; }

        public Calc Calculation { get; set; }
    }

    public class ErrorView
    {
        public DateTime Date { get; set; }

        public Calc Calculation { get; set; }

        public String error { get; set; }
    }
}
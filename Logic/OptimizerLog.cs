using System;
using System.Globalization;
using System.IO;
using System.Linq;
using Optimize;
using System.Web.Script.Serialization;
using OfficeOpenXml;

namespace Logic
{
    public class OptimizerLog
    { 
        private static ApplicationDbContext database = new ApplicationDbContext();

        public static  void ErrorLog(Calc calcOutput, string errorString)
        {
            try
            {
                database.Errors.Add(new Error()
                {
                    DateTime = DateTime.Now,
                    Calculation = calcOutput != null ? new JavaScriptSerializer().Serialize(calcOutput) : "null",
                    error = errorString
                });

                database.SaveChanges();
            }
            catch (Exception ex)
            {
                //throw ex;
            }
        }

        public static  void ActivityLog(Calc calcOutput)
        {
            try
            {
                database.Activities.Add(new Activity()
                {
                    DateTime = DateTime.Now,
                    Calculation = calcOutput != null ? new JavaScriptSerializer().Serialize(calcOutput) : "null"
                });

                database.SaveChanges();
            }
            catch (Exception ex)
            {
                //throw ex;
                ErrorLog(calcOutput, ex.ToString());
            }
        }
    }
}

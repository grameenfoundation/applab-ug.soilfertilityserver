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
        private static FileInfo errorLogFile = new FileInfo(@"C:\inetpub\wwwroot\Temp\Error.xlsm");

        private static FileInfo activityLogFile = new FileInfo(@"C:\inetpub\wwwroot\Temp\Activity.xlsm");

        private static string currentMonth = CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(DateTime.Now.Month);

        public static void ErrorLog(Calc calcOutput, string errorString)
        {
            try
            {
                using (ExcelPackage package = new ExcelPackage(errorLogFile))
                {
                    //Add Worksheet if it doesnot exist

                    if (package.Workbook.Worksheets.All(worksheet => worksheet.Name != currentMonth))
                        package.Workbook.Worksheets.Add(currentMonth);

                    ExcelWorksheet workSheet = package.Workbook.Worksheets[currentMonth];

                    //Add Header Row
                    if (workSheet.Dimension==null)
                    {
                        workSheet.Cells["A1"].Value = "Date Time";
                        workSheet.Cells["B1"].Value = "Calculation";
                        workSheet.Cells["C1"].Value = "Error";  
                    }
                    var lastRow = workSheet.Dimension.End.Row;

                    workSheet.Cells["A" + (lastRow + 1)].Value = DateTime.Now.ToString();
                    workSheet.Cells["B" + (lastRow + 1)].Value = calcOutput != null ? new JavaScriptSerializer().Serialize(calcOutput) : "null";
                    workSheet.Cells["C" + (lastRow + 1)].Value = errorString;
                     
                    package.Save();
                    package.Dispose();
                }
                
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public static void ActivityLog(Calc calcOutput)
        {
            try
            {
                using (ExcelPackage package = new ExcelPackage(activityLogFile))
                {
                    //Add Worksheet if it doesnot exist

                    if (package.Workbook.Worksheets.All(worksheet => worksheet.Name != currentMonth))
                        package.Workbook.Worksheets.Add(currentMonth);

                    ExcelWorksheet workSheet = package.Workbook.Worksheets[currentMonth];

                    //Add Header Row
                    if (workSheet.Dimension == null)
                    {
                        workSheet.Cells["A1"].Value = "Date Time";
                        workSheet.Cells["B1"].Value = "Calculation Details";
                    }
                    var lastRow = workSheet.Dimension.End.Row;

                    workSheet.Cells["A" + (lastRow + 1)].Value = DateTime.Now.ToString();
                    workSheet.Cells["B" + (lastRow + 1)].Value = calcOutput != null ? new JavaScriptSerializer().Serialize(calcOutput) : "null";

                    package.Save();
                    package.Dispose();
                }
            }
            catch (Exception ex)
            {
                //throw ex;
                ErrorLog(calcOutput, ex.ToString());
            }
        }
    }
}

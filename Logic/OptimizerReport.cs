using System;
using System.Collections.Generic;
using System.IO;
using Optimize;
using System.Web.Script.Serialization;
using OfficeOpenXml;
using Error = Optimize.Error;

namespace Logic
{
    public class OptimizerReport
    {
        static FileInfo errorFile = new FileInfo(@"C:\inetpub\wwwroot\Temp\Error.xlsm"); //Error file
        static FileInfo activityFile = new FileInfo(@"C:\inetpub\wwwroot\Temp\Activity.xlsm"); //Activity file
        static FileInfo newTempFile = new FileInfo(@"C:\inetpub\wwwroot\Temp\temp_" + DateTime.Now.Ticks + ".xlsm");

        public static List<Error> ErrorReport()
        {
            errorFile.CopyTo(newTempFile.ToString());
            try
            {
                var report = new List<Error>();
                using (ExcelPackage package = new ExcelPackage(newTempFile))
                {
                    foreach (ExcelWorksheet workSheet in package.Workbook.Worksheets)
                    {
                        if (workSheet.Dimension != null)
                        {
                            var start = workSheet.Dimension.Start;
                            var end = workSheet.Dimension.End;

                            for (int row = start.Row+1; row <= end.Row; row++)
                            {
                                var error = new Error
                                {
                                    dateTime = DateTime.Parse(workSheet.Cells[row, 1].Text
                                    //,CultureInfo.GetCultureInfo("en-US")
                                        ),
                                    calculation = workSheet.Cells[row, 2].Text != null
                                        ? new JavaScriptSerializer().Deserialize<Calc>(workSheet.Cells[row, 2].Text)
                                        : new Calc(),
                                    error = workSheet.Cells[row, 3].Text
                                };

                                report.Add(error);
                            }

                        }
                    }
                    package.Dispose();
                }
               newTempFile.Delete();
                return report;
            }
            catch (Exception ex)
            {
                newTempFile.Delete();
                throw ex;
                // OptimizerLog.ErrorLog(clientInputs, ex.ToString());
            }
        }

        public static List<Activity> ActivityReport()
        {
            activityFile.CopyTo(newTempFile.ToString());
            try
            {
                var report = new List<Activity>();
                using (ExcelPackage package = new ExcelPackage(newTempFile))
                {
                    foreach (ExcelWorksheet workSheet in package.Workbook.Worksheets)
                    {
                        if (workSheet.Dimension != null)
                        {
                            var start = workSheet.Dimension.Start;
                            var end = workSheet.Dimension.End;

                            for (int row = start.Row+1; row <= end.Row; row++)
                            {
                                var activity = new Activity
                                {
                                    dateTime = DateTime.Parse(workSheet.Cells[row, 1].Text
                                    //,CultureInfo.GetCultureInfo("en-US")
                                    ),
                                    calculation =
                                        workSheet.Cells[row, 2].Text != null
                                            ? new JavaScriptSerializer().Deserialize<Calc>(workSheet.Cells[row, 2].Text)
                                            : new Calc()
                                };
                                report.Add(activity);   
                            }
                        }
                    }
                    package.Dispose();
                }
                newTempFile.Delete();
                return report;
            }
            catch (Exception ex)
            {
                newTempFile.Delete();
                throw ex;
                 //OptimizerLog.ErrorLog(clientInputs, ex.ToString());

            }
        }
    }
}

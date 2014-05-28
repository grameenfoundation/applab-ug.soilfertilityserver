using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Services;
using System.Web.Script.Serialization;
//using Microsoft.CSharp.RuntimeBinder.Binder;

using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using Microsoft.Office.Interop.Excel;

namespace Optimizer
{
    /// <summary>
    /// Summary description for Service1
    /// </summary>
    [WebService(Namespace = "http://tempuri.org/")]
    [WebServiceBinding(ConformsTo = WsiProfiles.BasicProfile1_1)]
    [System.ComponentModel.ToolboxItem(false)]
    // To allow this Web Service to be called from script, using ASP.NET AJAX, uncomment the following line. 
    // [System.Web.Script.Services.ScriptService]
    public class Service1 : System.Web.Services.WebService
    {
        public static object key = new object();
        public static object key2 = new object();
        public static String[] filesInUse = new String[10];


        [WebMethod]
        public String Optimize(String json)
        {
            lock (key)
            {
                String jsonStr = openExel(json);
                return jsonStr;
            }
        }

        protected virtual bool IsFileLocked(FileInfo file)
        {
            FileStream stream = null;

            try
            {
                stream = file.Open(FileMode.Open, FileAccess.ReadWrite, FileShare.None);
            }
            catch (IOException)
            {
                return true;
            }
            finally
            {
                if (stream != null)
                    stream.Close();
            }

            //file is not locked
            return false;
        }

        private String getAvailableFile()
        {
            string sourcePath = AppDomain.CurrentDomain.BaseDirectory;
            string targetPath = System.IO.Path.Combine(sourcePath, "calculations");

            if (!System.IO.Directory.Exists(targetPath))
            {
                System.IO.Directory.CreateDirectory(targetPath);
            }

            String[] fileNames = Directory.GetFiles(targetPath, "*.xlsm");

            //log(fileNames.Length + "files found");

            String fileName = "";

            lock (key2)
            {
                do
                {
                    for (int i = 0; i < fileNames.Length; i++)
                    {
                        if (fileName == "" && !IsFileLocked(new FileInfo(fileNames[i])))
                        {
                            fileName = fileNames[i];
                            break;
                        }
                    }

                    if (fileName != "")
                        for (int i = 0; i < Service1.filesInUse.Length; i++)
                            if (Service1.filesInUse[i] == fileName)
                            {
                                fileName = "";
                                break;
                            }

                } while (fileName == "");

                for (int i = 0; i < Service1.filesInUse.Length; i++)
                {
                    if (Service1.filesInUse[i] == null)
                    {
                        Service1.filesInUse[i] = fileName;
                        break;
                    }
                }
            }

            return fileName;
        }

        public String openExel(String json)
        {

            //string to hold return value
            String jsonStr = "";

            Process[] P0, P1;
            P0 = Process.GetProcessesByName("Excel");
            Excel.Application excelApp = new Excel.Application();

            int I, J;
            P1 = Process.GetProcessesByName("Excel");
            I = 0;

            if (P1.Length > 1)
            {
                for (I = 0; I < P1.Length; I++)
                {
                    for (J = 0; J < P0.Length; J++)
                        if (P0[J].Id == P1[I].Id) break;
                    if (J == P0.Length) break;
                }
            }

            Process excelProcess_4_thisRequests = P1[I];

            excelApp.Visible = true;

            string targetFile = getAvailableFile();
            if (targetFile == "" || targetFile == null)
                return "ERROR";

            Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(targetFile,
                        0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "",
                        true, false, 0, true, false, false);
            if (excelWorkbook.HasPassword)
                excelWorkbook.Unprotect("52Fre");

            excelApp.Run("Reset_Form");

            try
            {

                Worksheet excelWorksheet = (Worksheet)excelWorkbook.Worksheets["Fertilizer Optimization"];

                try
                {
                    excelWorksheet.Unprotect("52Fre");
                }
                catch
                {
                    Log("Ooops, work sheet was not protected");
                }

                //Thread.Sleep(10000);
                //================================================
                excelApp.Run("Reset_Form");

                JavaScriptSerializer ser = new JavaScriptSerializer();
                Calc calc2 = ser.Deserialize<Calc>(json); 

                if (calc2.Id == null)
                {
                    Guid g = Guid.NewGuid();
                    calc2.Id = g.ToString();
                }

                calc2.File = targetFile;

                foreach (CalcCrop cc in calc2.CalcCrops)
                {
                    initializeCropCell(cc, excelWorksheet);
                }

                foreach (CalcFertilizer cf in calc2.CalcFertilizers)
                {
                    initializeFertilizerCell(cf, excelWorksheet);
                }


                excelWorksheet.get_Range("C34", "C34").Value = calc2.AmtAvailable;

                //=======================================
                excelApp.Run("Optimize_Solver");

                Double totReturns = getDoubleValue(excelWorksheet, "V76");
                calc2.TotNetReturns = totReturns;

                foreach (CalcCrop cc in calc2.CalcCrops)
                {
                    foreach (CalcFertilizer cf in calc2.CalcFertilizers)
                    {
                        Double val = getDoubleValue(excelWorksheet, getCalcCropFertilizerAmountCell(cc.Crop, cf.Fertilizer));
                        // ccfr.Amt = val;
                        if (val > 0.0)
                        {
                            CalcCropFertilizerRatio ccfr = new CalcCropFertilizerRatio(cc.Crop, cf.Fertilizer, getDoubleValue(excelWorksheet, getCalcCropFertilizerAmountCell(cc.Crop, cf.Fertilizer)));
                            calc2.CalcCropFertilizerRatios.Add(ccfr);
                        }
                    }
                }

                //============================================

                foreach (CalcFertilizer cf in calc2.CalcFertilizers)
                {
                    Double val = getDoubleValue(excelWorksheet, getTotalFertilizerRequiredCell(cf));
                    cf.TotalRequired = val;
                }

                //============================================

                foreach (CalcCrop cc in calc2.CalcCrops)
                {
                    Double val = getDoubleValue(excelWorksheet, getCropYeildIncreaseCell(cc));
                    cc.YieldIncrease = val;

                    val = getDoubleValue(excelWorksheet, getCropNetReturnsCell(cc));
                    cc.NetReturns = val;
                }

                jsonStr = ser.Serialize(calc2);

            }
            catch (Exception ex)
            {
                Log(ex.StackTrace);
            }
            finally
            {
                excelWorkbook.Save();
                excelWorkbook.Close();
                excelApp.Quit();

                //~~> Clean Up
                releaseObject(excelApp);
                releaseObject(excelWorkbook);

                //kill the damn process
                excelProcess_4_thisRequests.Kill();

                releaseFile(targetFile);
            }

            return jsonStr;
        }

        public void releaseFile(String fileName)
        {
            for (int i = 0; i < Service1.filesInUse.Length; i++)
            {
                if (Service1.filesInUse[i] == fileName)
                    Service1.filesInUse[i] = null;
            }
        }

        //~~> Release the objects
        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
            }
            finally
            {
                GC.Collect();
            }
        }


        private void initializeCropCell(CalcCrop c, Worksheet excelWorksheet)
        {
            switch (c.Crop.Name.ToLower())
            {
                case "maize":
                    populateCell(excelWorksheet, "C16", c.Area);
                    populateCell(excelWorksheet, "D16", c.Profit);
                    break;
                case "sorghum":
                    populateCell(excelWorksheet, "C17", c.Area);
                    populateCell(excelWorksheet, "D17", c.Profit);
                    break;
                case "upland rice, paddy":
                    populateCell(excelWorksheet, "C18", c.Area);
                    populateCell(excelWorksheet, "D18", c.Profit);
                    break;
                case "beans":
                    populateCell(excelWorksheet, "C19", c.Area);
                    populateCell(excelWorksheet, "D19", c.Profit);
                    break;
                case "soybeans":
                    populateCell(excelWorksheet, "C20", c.Area);
                    populateCell(excelWorksheet, "D20", c.Profit);
                    break;
                case "groundnuts, unshelled":
                    populateCell(excelWorksheet, "C21", c.Area);
                    populateCell(excelWorksheet, "D21", c.Profit);
                    break;

            }

        }

        private void initializeFertilizerCell(CalcFertilizer cf, Worksheet excelWorksheet)
        {
            switch (cf.Fertilizer.Name.ToLower())
            {
                case "urea":
                    populateCell(excelWorksheet, "F26", cf.Price);
                    break;
                case "triple super phosphate, tsp":
                    populateCell(excelWorksheet, "F27", cf.Price);
                    break;
                case "diammonium phosphate, dap":
                    populateCell(excelWorksheet, "F28", cf.Price);
                    break;
                case "murate of potash, kcl":
                    populateCell(excelWorksheet, "F29", cf.Price);
                    break;
            }
        }

        private void populateCell(Worksheet excelWorksheet, String cell, int value)
        {
            Range soyabeanHa = (Range)excelWorksheet.get_Range(cell, cell);
            soyabeanHa.Value = value;
        }

        private void populateCell(Worksheet excelWorksheet, String cell, double value)
        {
            Range soyabeanHa = (Range)excelWorksheet.get_Range(cell, cell);
            soyabeanHa.Value = value;
        }

        private Double getDoubleValue(Worksheet excelWorksheet, String cell)
        {
            Double val = (Double)excelWorksheet.get_Range(cell, cell).Value2;
            return val;
        }

        private int getIntValue(Worksheet excelWorksheet, String cell)
        {
            int val = (int)excelWorksheet.get_Range(cell, cell).Value2;
            return val;
        }

        private String getCalcCropFertilizerAmountCell(Crop crop, Fertilizer fert)
        {
            String fertilizerColumn = "";
            String cropRow = "";

            switch (fert.Name.ToLower())
            {
                case "urea":
                    fertilizerColumn = "O";
                    break;
                case "triple super phosphate, tsp":
                    fertilizerColumn = "P";
                    break;
                case "diammonium phosphate, dap":
                    fertilizerColumn = "Q";
                    break;
                case "murate of potash, kcl":
                    fertilizerColumn = "R";
                    break;
            }

            switch (crop.Name.ToLower())
            {
                case "maize":
                    cropRow = "32";
                    break;
                case "sorghum":
                    cropRow = "33";
                    break;
                case "upland rice, paddy":
                    cropRow = "34";
                    break;
                case "beans":
                    cropRow = "35";
                    break;
                case "soybeans":
                    cropRow = "36";
                    break;
                case "groundnuts, unshelled":
                    cropRow = "37";
                    break;

            }

            return fertilizerColumn + cropRow;
        }

        private String getTotalFertilizerRequiredCell(CalcFertilizer cf)
        {
            String fertilizerColumn = "";
            String cropRow = "";

            switch (cf.Fertilizer.Name.ToLower())
            {
                case "urea":
                    fertilizerColumn = "O";
                    break;
                case "triple super phosphate, tsp":
                    fertilizerColumn = "P";
                    break;
                case "diammonium phosphate, dap":
                    fertilizerColumn = "Q";
                    break;
                case "murate of potash, kcl":
                    fertilizerColumn = "R";
                    break;
            }

            cropRow = "39";

            return fertilizerColumn + cropRow;
        }

        private String getCropYeildIncreaseCell(CalcCrop cc)
        {
            String column = "C";
            String cropRow = "";

            switch (cc.Crop.Name.ToLower())
            {
                case "maize":
                    cropRow = "54";
                    break;
                case "sorghum":
                    cropRow = "55";
                    break;
                case "upland rice, paddy":
                    cropRow = "56";
                    break;
                case "beans":
                    cropRow = "57";
                    break;
                case "soybeans":
                    cropRow = "58";
                    break;
                case "groundnuts, unshelled":
                    cropRow = "59";
                    break;

            }

            return column + cropRow;
        }

        private String getCropNetReturnsCell(CalcCrop cc)
        {
            String column = "D";
            String cropRow = "";

            switch (cc.Crop.Name.ToLower())
            {
                case "maize":
                    cropRow = "54";
                    break;
                case "sorghum":
                    cropRow = "55";
                    break;
                case "upland rice, paddy":
                    cropRow = "56";
                    break;
                case "beans":
                    cropRow = "57";
                    break;
                case "soybeans":
                    cropRow = "58";
                    break;
                case "groundnuts, unshelled":
                    cropRow = "59";
                    break;

            }

            return column + cropRow;
        }

        public static void Log(string logMessage)
        {
            String logname = AppDomain.CurrentDomain.BaseDirectory + "logs\\" + DateTime.Now.Year+ "-" + DateTime.Now.Month.ToString("D2") + "-" + DateTime.Now.Day.ToString("D2") + ".log";
            using (StreamWriter w = File.AppendText(logname))
            {
                w.WriteLine("{0} {1}", DateTime.Now.ToLongTimeString(),
                    DateTime.Now.ToLongDateString());
                w.WriteLine("  ");
                w.WriteLine("  {0}", logMessage);
                w.WriteLine("-------------------------------");
            }
        }
    }
}
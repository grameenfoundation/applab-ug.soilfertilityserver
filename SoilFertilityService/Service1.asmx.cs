using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Services;
using System.Web.Script.Serialization;

using Microsoft.CSharp;

using System.Threading;
using System.IO;
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
            //Calculator c = new Calculator(key, this);
            Calculator c = new Calculator();

            return c.openExel(json);
        }
    }

    public class Calculator {

        public Calculator()//object key, Service1 s
        {
           // this.key = key;
           // this.s = s;
        }

        private String json;

        public string Json
        {
            get { return json; }
            set { json = value; }
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
                //the file is unavailable because it is:
                //still being written to
                //or being processed by another thread
                //or does not exist (has already been processed)
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

            log(fileNames.Length + "files found");

            String fileName = "";

            lock (Service1.key2)
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

        private void log(string message)
        {
            String timeNow = DateTime.Now.ToString(@"MM\/dd\/yyyy h\:mm\:ss tt");
            String logFile = AppDomain.CurrentDomain.BaseDirectory + "log.txt";
            lock (Service1.key)
            {
                File.AppendAllText(logFile, timeNow +" => " + Thread.CurrentThread.ManagedThreadId + " " +message+"\n");
            }
        }

        public String openExel(String json)//Object obj
        {
            this.json = "";
            // String json = (String)obj;

            String jsonStr = "";

            log("Starting process with " + json);

           string destFile = getAvailableFile();

           if (destFile == "" || destFile == null)
               return "ERROR";

           log("Using " + destFile);

            try
            {

                Application excelApp = new Application();
                excelApp.Visible = true;


                Workbook excelWorkbook = excelApp.Workbooks.Open(destFile,
                            0, false, 5, "", "", false, XlPlatform.xlWindows, "",
                            true, false, 0, true, false, false);
                excelWorkbook.Unprotect("52Fre");

                Worksheet excelWorksheet = (Worksheet)excelWorkbook.Worksheets["Fertilizer Optimization"];
                excelWorksheet.Unprotect("52Fre");

            // Thread.Sleep(10000);
            //================================================

            excelApp.Run("Reset_Form");

            try
            {
                //================================================

                JavaScriptSerializer ser = new JavaScriptSerializer();
                // String jsonStr = ser.Serialize(calc);

                Calc calc2 = ser.Deserialize<Calc>(json);

                if (calc2.Id == null)
                {
                    Guid g = Guid.NewGuid();
                    calc2.Id = g.ToString();
                }

                calc2.File = destFile;

                foreach (CalcCrop cc in calc2.CalcCrops)
                {
                    initializeCropCell(cc, excelWorksheet);
                }

                foreach (CalcFertilizer cf in calc2.CalcFertilizers)
                {
                    initializeFertilizerCell(cf, excelWorksheet);
                }


                excelWorksheet.get_Range("C34", "C34").Value = calc2.AmtAvailable;


                Range excelCell = (Range)excelWorksheet.get_Range("M1", "M1");

                excelCell.Value = json;

                log("B4 running Optimize_Solver");

                //~~> Run the macros by supplying the necessary arguments
                excelApp.Run("Optimize_Solver");//comma separated params

                log("AFTER running Optimize_Solver");


                Double totReturns = getDoubleValue(excelWorksheet, "V76");
                calc2.TotNetReturns = totReturns;

                log("B4 Copying Crop-Fertilizer-Amounts");

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


                log("AFTER Copying Crop-Fertilizer-Amounts");

                //================================================================================================

                log("B4 Copying Total Fertilizer Required");
                foreach (CalcFertilizer cf in calc2.CalcFertilizers)
                {
                    Double val = getDoubleValue(excelWorksheet, getTotalFertilizerRequiredCell(cf));
                    cf.TotalRequired = val;
                }
                log("AFTER Copying Total Fertilizer Required");

                //=================================================================================================

                log("B4 Copying Crop Yeild Increase and Net Returns");
                foreach (CalcCrop cc in calc2.CalcCrops)
                {
                    Double val = getDoubleValue(excelWorksheet, getCropYeildIncreaseCell(cc));
                    cc.YieldIncrease = val;

                    val = getDoubleValue(excelWorksheet, getCropNetReturnsCell(cc));
                    cc.NetReturns = val;
                }
                log("AFTER Copying Crop Yeild Increase and Net Returns");


                jsonStr = ser.Serialize(calc2);

                log("Closing");

                excelWorkbook.Save();
                excelWorkbook.Close();
                excelApp.Quit();

                //~~> Clean Up
                releaseObject(excelWorksheet);
                releaseObject(excelWorkbook);
                releaseObject(excelApp);

            }
            catch (Exception ex)
            {

                log("Closing after ERROR");

                log(ex.StackTrace);
                excelWorkbook.Save();
                excelWorkbook.Close();
                excelApp.Quit();

                //~~> Clean Up
                releaseObject(excelWorksheet);
                releaseObject(excelWorkbook);
                releaseObject(excelApp);
            }
            finally
            {
                 
            }

            }
            catch (Exception ex)
            {
                log("Some wired exception occured, see below");
                log(ex.StackTrace);
            }

            //this.json = jsonStr;
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

            //xlApp.Quit();

            //release all memory - stop EXCEL.exe from hanging around.
            //if (xlWorkBook != null) { Marshal.ReleaseComObject(xlWorkBook); } //release each workbook like this
            //if (xlWorkSheet != null) { Marshal.ReleaseComObject(xlWorkSheet); } //release each worksheet like this
            //if (xlApp != null) { Marshal.ReleaseComObject(xlApp); } //release the Excel application
            //xlWorkBook = null; //set each memory reference to null.
            //xlWorkSheet = null;
            // xlApp = null;
            //GC.Collect();

            try
            {
                log("Releasing obj " + obj.ToString());
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                log("Error releasing obj " + obj.ToString());
                log(ex.Message);

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

        private void initializeFertilizerCell(CalcFertilizer cf,  Worksheet excelWorksheet)
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
             Range soyabeanHa = ( Range)excelWorksheet.get_Range(cell, cell);
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
    }
}
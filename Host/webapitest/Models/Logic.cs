using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Script.Serialization;
using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using Optimize;
using System;
using System.Collections.Generic;
using System.IO;
using System.Web.Http;
using System.Web.Script.Serialization;
using OfficeOpenXml;
using Optimize;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Ajax.Utilities;
using Microsoft.Office.Interop.Excel;

namespace Optimize
{
    public class BI
    {

        public string Optimizer(string json)
        {
            //json = "{AmtAvailable":5.0E8,"CalcCrops":[{"Area":4.04686,"Crop":{"Name":"Maize"},"Profit":15.0,"Id":0}],"CalcFertilizers":[{"Fertilizer":{"Name":"Urea"},"Price":300,"id":0}],"CalcCropFertilizerRatios":[],"FarmerName":"564654654654654654654654","Id":"3862079c-da47-47af-8a2f-3d66c05eab33","Imei":"000000000000000"}"

            //For Testing purpose
            var calc = new Calc()
            {
                CalcCrops = new List<CalcCrop>() { new CalcCrop(new Crop() { Name = "maize" }, 20, 2) },
                AmtAvailable = 25000000,
                FarmerName = "Josh",
                //CalcFertilizers = new List<CalcFertilizer>(){new CalcFertilizer(new Fertilizer() { Name = "urea" }, 2500)}
            };
            json = new JavaScriptSerializer().Serialize(calc);
            //end of testing data
            try
            {
                FileInfo newFile = null;
                if (CreateCopyofFile(out newFile))
                {
                    if (InputClientEntries(newFile, json))
                    {
                        ExecuteOptimizeMacro(newFile.FullName);
                    }
                    //return true;
                }
                //return false;
                return json.ToString();
            }
            catch (Exception ex)
            {

                throw ex;
            }

        }

        public bool CreateCopyofFile(out FileInfo newFile)
        {
            //newFile = new FileInfo(outputDir.Name + @"C\inetpub\wwwroot\Temp\optimizer_" + DateTime.Now.Ticks + ".xlsm");//New file based on the time stamp
            newFile = new FileInfo(@"C:\inetpub\wwwroot\Temp\optimizer_" + DateTime.Now.Ticks + ".xlsm");//New file based on the time stamp

            FileInfo optimizerFile = new FileInfo(@"C:\inetpub\wwwroot\Temp\Optimizer.xlsm");//Template file

            if (optimizerFile.Exists)
            {
                optimizerFile.CopyTo(newFile.ToString()); // ensures we create a new workbook
                return true;
            }
            return false;
        }

        public bool InputClientEntries(FileInfo file, string json)
        {
            try
            {
                using (ExcelPackage package = new ExcelPackage(file))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets["Fertilizer Optimization"];

                    //Deserialise json string
                    Calc clientInputs = new JavaScriptSerializer().Deserialize<Calc>(json);

                    //Initialise Crop Values in excel worksheet
                    InitializeCrops(clientInputs.CalcCrops, worksheet);

                    //Initialise Fertilizer Values in excel worksheet
                    InitializeFertilizers(clientInputs.CalcFertilizers, worksheet);

                    //
                    worksheet.Cells["C34"].Value = clientInputs.AmtAvailable;
                    worksheet.Cells["F26"].Value = 2; //temp set for testing urea

                    package.Save();
                    package.Dispose();

                    return true;
                }

            }
            catch (Exception ex)
            {
                throw ex;
                return false;
            }

        }

        private void InitializeCrops(List<CalcCrop> calcCrops, ExcelWorksheet worksheet)
        {
            foreach (CalcCrop calcCrop in calcCrops)
            {
                switch (calcCrop.Crop.Name.ToLower())
                {
                    case "maize":
                        worksheet.Cells["C16"].Value = calcCrop.Area;
                        worksheet.Cells["D16"].Value = calcCrop.Profit;
                        break;
                    case "sorghum":
                        worksheet.Cells["C17"].Value = calcCrop.Area;
                        worksheet.Cells["D17"].Value = calcCrop.Profit;
                        break;
                    case "upland rice, paddy":
                        worksheet.Cells["C18"].Value = calcCrop.Area;
                        worksheet.Cells["D18"].Value = calcCrop.Profit;
                        break;
                    case "beans":
                        worksheet.Cells["C19"].Value = calcCrop.Area;
                        worksheet.Cells["D19"].Value = calcCrop.Profit;
                        break;
                    case "soybeans":
                        worksheet.Cells["C20"].Value = calcCrop.Area;
                        worksheet.Cells["D20"].Value = calcCrop.Profit;
                        break;
                    case "groundnuts, unshelled":
                        worksheet.Cells["C21"].Value = calcCrop.Area;
                        worksheet.Cells["D21"].Value = calcCrop.Profit;
                        break;

                }
            }

        }

        private void InitializeFertilizers(List<CalcFertilizer> calcFertilizers, ExcelWorksheet worksheet)
        {
            foreach (var calcFertilizer in calcFertilizers)
            {
                switch (calcFertilizer.Fertilizer.Name.ToLower())
                {
                    case "urea":
                        worksheet.Cells["F26"].Value = calcFertilizer.Price;
                        break;
                    case "triple super phosphate, tsp":
                        worksheet.Cells["F27"].Value = calcFertilizer.Price;
                        break;
                    case "diammonium phosphate, dap":
                        worksheet.Cells["F28"].Value = calcFertilizer.Price;
                        break;
                    case "murate of potash, kcl":
                        worksheet.Cells["F29"].Value = calcFertilizer.Price;
                        break;
                }
            }
        }

        private void ExecuteOptimizeMacro(string targetFile)
        {
            try
            {
                Excel.Application excelApp = new Excel.Application();

                Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(targetFile,
                    0, false, 5, "52Fre", "52Fre", false, Excel.XlPlatform.xlWindows, "",
                    true, false, 0, true, false, false);

                if (!excelWorkbook.Name.IsNullOrWhiteSpace())
                {
                    if (excelWorkbook.HasPassword)
                        excelWorkbook.Unprotect("52Fre");
                    Worksheet worksheet = excelWorkbook.Worksheets["Fertilizer Optimization"];

                    worksheet.Unprotect("52Fre");

                    //excelApp.Run("Reset_Form");
                    excelApp.Run("Optimize_Solver");


                    // Clean Up the memory
                    excelWorkbook.Save();
                    excelWorkbook.Close();

                    excelApp.Quit();

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(excelWorkbook);
                    GC.Collect();
                }
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }
    }
}
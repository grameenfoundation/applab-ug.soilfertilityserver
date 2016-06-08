using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Optimize;
using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using Microsoft.Ajax.Utilities;

namespace Logic
{
    public class Optimizer
    {
        public Calc calc;
         
 
        public string optimisationWorksheet = "Fertilizer Optimization";

        public Calc Optimize(Calc json)
        {
            calc = json;
            try
            {
                //calc.Database = database;//OptimizerManager.DatabaseCheck();

                FileInfo newFile = null;
                if (CreateCopyofFile(calc.Region, out newFile))
                {
                    if (InputClientEntries(newFile, calc))
                    {
                        ExecuteOptimizeMacro(newFile.FullName);
                    }
                    //return true;
                }
                //return false;
                return  OptimizerOutput.OutputResults(newFile, calc);
            }
            catch (Exception ex)
            { 
                OptimizerLog.ErrorLog( calc, ex.ToString());
                return calc;
            }

        }

        public bool CreateCopyofFile(int file, out FileInfo newFile)
        { 

            //New file based on the time stamp
            newFile = new FileInfo(@"C:\inetpub\wwwroot\Temp\Optimizer_"+file+"_" + DateTime.Now.Ticks + ".xlsm");
            //New file based on the time stamp

            FileInfo optimizerFile = new FileInfo(@"C:\inetpub\wwwroot\Temp\Optimizer_"+file+".xlsm"); //Template file

            if (!optimizerFile.Exists) return false;

            optimizerFile.CopyTo(newFile.ToString()); // ensures we create a new workbook
            return true;
        }

        public bool InputClientEntries(FileInfo file, Calc clientInputs)
        {
            try
            {
                using (ExcelPackage package = new ExcelPackage(file))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[optimisationWorksheet];

                    //Initialise Crop Values in excel worksheet
                    InitializeCrops(clientInputs.CalcCrops, worksheet);

                    //Initialise Fertilizer Values in excel worksheet
                    InitializeFertilizers(clientInputs.CalcFertilizers, worksheet);

                    //
                    worksheet.Cells["C34"].Value = clientInputs.AmtAvailable;
                   
                    package.Save();
                    package.Dispose();

                    return true;
                }

            }
            catch (Exception ex)
            {
                //throw ex;
                OptimizerLog.ErrorLog(clientInputs, ex.ToString());
                return false;
            }

        }

        private void InitializeCrops(List<CalcCrop> calcCrops, ExcelWorksheet worksheet)
        {
            //Cell Range with crop inputs
            var searchableCells = worksheet.Cells[16, 2, 24, 2];

            foreach (CalcCrop calcCrop in calcCrops)
            {
               var cropCells =   searchableCells.Where(a => a.Text.Contains(calcCrop.Crop.Name.Substring(0, 3))).ToList();

                if (cropCells.Any())
                {
                    var cropCell = cropCells.First().Address.Substring(1, 2);
                    worksheet.Cells["C"+cropCell].Value = calcCrop.Area;
                    worksheet.Cells["D" + cropCell].Value = calcCrop.Profit;
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
                    case "NPK":
                        worksheet.Cells["F30"].Value = calcFertilizer.Price;
                        break;
                }
            }
        }

        private void ExecuteOptimizeMacro(string targetFile)
        {
            try
            {
                Application excelApp = new Application();

                Workbook excelWorkbook = excelApp.Workbooks.Open(targetFile,
                    0, false, 5, "52Fre", "52Fre", false, XlPlatform.xlWindows, "",
                    true, false, 0, true, false, false);

                if (!excelWorkbook.Name.IsNullOrWhiteSpace())
                {
                    if (excelWorkbook.HasPassword)
                        excelWorkbook.Unprotect("52Fre");
                    Worksheet worksheet = excelWorkbook.Worksheets[optimisationWorksheet];

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
                OptimizerLog.ErrorLog(calc, ex.ToString());
                
            }
        }
       
    }
}

using System;
using System.IO;
using System.Linq;
using Optimize;
using OfficeOpenXml;

namespace Logic
{
    public class OptimizerOutput
    {
        public static string optimisationWorksheet = "Fertilizer Optimization";

        public static Calc OutputResults(FileInfo file, Calc calcOutput)
        {
            try
            {
                using (ExcelPackage package = new ExcelPackage(file))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[optimisationWorksheet];

                    calcOutput.TotNetReturns = worksheet.Cells["C62"].GetValue<double>();

                    //calcOutput.Id = "simpleId";
                    //calcOutput.File = file.FullName;

                    foreach (CalcCrop calcCrop in calcOutput.CalcCrops)
                    {
                        calcCrop.YieldIncrease = getCropYeildIncrease_NetReturns(calcCrop, worksheet)[0];

                        calcCrop.NetReturns = getCropYeildIncrease_NetReturns(calcCrop, worksheet)[1];

                        foreach (CalcFertilizer calcFertilizer in calcOutput.CalcFertilizers)
                        {
                            calcFertilizer.TotalRequired = getTotalFertilizerRequired(calcFertilizer, worksheet);

                            Double val = getCalcCropFertilizerAmount(calcCrop.Crop, calcFertilizer.Fertilizer, worksheet);

                            if (val > 0)
                            {
                                CalcCropFertilizerRatio calcCropFertilizerRatio =
                                    new CalcCropFertilizerRatio(calcCrop.Crop, calcFertilizer.Fertilizer, val);
                                calcOutput.CalcCropFertilizerRatios.Add(calcCropFertilizerRatio);
                            }
                        }
                    }

                    //Log activity for auditing purposes
                    OptimizerLog.ActivityLog(calcOutput);

                    package.Save();
                    package.Dispose();
                }

                //Do a file cleanup
                file.Delete();

                return calcOutput;
            }
            catch (Exception ex)
            {
                //throw ex;
                OptimizerLog.ErrorLog(calcOutput, ex.ToString());
                return calcOutput;
            }
        }

        private static double getCalcCropFertilizerAmount(Crop crop, Fertilizer fert, ExcelWorksheet worksheet)
        {
            String fertilizerColumn = "";
            String cropRow = "";

            switch (fert.Name.ToLower())
            {
                case "urea":
                    fertilizerColumn = "C";
                    break;
                case "triple super phosphate, tsp":
                    fertilizerColumn = "D";
                    break;
                case "diammonium phosphate, dap":
                    fertilizerColumn = "E";
                    break;
                case "murate of potash, kcl":
                    fertilizerColumn = "F";
                    break;
                case "NPK":
                    fertilizerColumn = "G";
                    break;
            }

            //Cell Range with crop OUTPUTS
            var searchableCells = worksheet.Cells[44,2,51,2];

            var cropCells = searchableCells.Where(a => a.Text.Contains(crop.Name.Substring(0, 3))).ToList();

            if (cropCells.Any())
            {
                cropRow = cropCells.First().Address.Substring(1, 2);
            }
           
            return worksheet.Cells[fertilizerColumn + cropRow].GetValue<double>();
        }

        private static double getTotalFertilizerRequired(CalcFertilizer cf,ExcelWorksheet worksheet)
        { 
            double value = 0;

             //Cell Range with crop:fertilizer ratios
            var searchableCells = worksheet.Cells[40,2,53,2];
            var totalFertilizerRows = searchableCells.Where(a => a.Text.Contains("Total fertilizer needed")).ToList();
            var neededRow = totalFertilizerRows.First().Address.Substring(1, 2);

            switch (cf.Fertilizer.Name.ToLower())
            {
                case "urea":
                    value = worksheet.Cells["C" + neededRow].GetValue<double>();
                    break;
                case "triple super phosphate, tsp":
                    value = worksheet.Cells["D" + neededRow].GetValue<double>();
                    break;
                case "diammonium phosphate, dap":
                    value = worksheet.Cells["E" + neededRow].GetValue<double>();
                    break;
                case "murate of potash, kcl":
                    value = worksheet.Cells["F" + neededRow].GetValue<double>();
                    break;
                case "NPK":
                    value = worksheet.Cells["G" + neededRow].GetValue<double>();
                    break;
            }

            return value;
        }

        private static double[] getCropYeildIncrease_NetReturns(CalcCrop calcCrop, ExcelWorksheet worksheet)
        { 
            double[] value = new double[] {};
            
            //Cell Range with crop OUTPUTS
            var searchableCells = worksheet.Cells[53, 2, 61, 2];
            var cropCells = searchableCells.Where(a => a.Text.Contains(calcCrop.Crop.Name.Substring(0, 3))).ToList();
            if (cropCells.Any())
            {
                var cropRow = cropCells.First().Address.Substring(1, 2);

                value = new[] { worksheet.Cells["C" + cropRow].GetValue<double>(), worksheet.Cells["D"+cropRow].GetValue<double>() };
            }
            return value;
        }

    }
}

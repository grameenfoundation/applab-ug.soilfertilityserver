using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Optimize;
using Microsoft.Ajax.Utilities;
using OfficeOpenXml;

namespace Logic
{
    public class OptimizerManager
    {
        static FileInfo databaseFile = new FileInfo(@"C:\inetpub\wwwroot\Temp\Database.xlsm"); //Database file
        static FileInfo newTempDatabaseFile = new FileInfo(@"C:\inetpub\wwwroot\Temp\Database" + DateTime.Now.Ticks + ".xlsm");

        static string regionWorksheet = "Regions";
        static string cropWorksheet = "Crops";
        static string regionCropWorksheet = "Region_Crops";
        static string updateWorksheet = "Update";
        static string optimisationWorksheet = "Fertilizer Optimization";

        //public static Database DatabaseCheck()
        //{
        //    try
        //    {
        //        //if (!databaseFile.Exists)
        //        //{
        //        //    databaseFile.Create().Flush();
                  
        //            using (ExcelPackage package = new ExcelPackage(databaseFile))
        //            {
        //                //Add missing Worksheets  

        //                if (package.Workbook.Worksheets.All(worksheet => worksheet.Name != regionWorksheet))
        //                {
        //                   var worksheet=  package.Workbook.Worksheets.Add(regionWorksheet);
        //                    worksheet.Cells["A1"].Value = "Id";
        //                    worksheet.Cells["B1"].Value = "Region Name";
        //                    worksheet.Cells["C1"].Value = "Units";
        //                }

        //                if (package.Workbook.Worksheets.All(worksheet => worksheet.Name != cropWorksheet))
        //                {
        //                   var worksheet= package.Workbook.Worksheets.Add(cropWorksheet); 
        //                    worksheet.Cells["B1"].Value = "Crop Name";
        //                }

        //                if (package.Workbook.Worksheets.All(worksheet => worksheet.Name != regionCropWorksheet))
        //                {
        //                    var worksheet =package.Workbook.Worksheets.Add(regionCropWorksheet);
        //                    worksheet.Cells["B1"].Value = "Region Id";
        //                    worksheet.Cells["C1"].Value = "Crop Name";
        //                }

        //                if (package.Workbook.Worksheets.All(worksheet => worksheet.Name != updateWorksheet))
        //                {
        //                    var worksheet= package.Workbook.Worksheets.Add(updateWorksheet);
        //                    worksheet.Cells["A1"].Value = "Last Date Modified";
        //                }

        //                package.Save();
        //                package.Dispose();
        //            }
        //        //}

        //        databaseFile.CopyTo(newTempDatabaseFile.ToString());

        //        var database = new Database(){Crops = new List<Crop>(),RegionCrops = new List<RegionCrop>(),Regions = new List<Region>()};

        //        using (ExcelPackage package = new ExcelPackage(newTempDatabaseFile))
        //        {
        //            //Add missing Worksheets  
 
        //            ExcelWorksheet regionsWorksheet = package.Workbook.Worksheets[regionWorksheet];
        //            ExcelWorksheet cropsWorksheet = package.Workbook.Worksheets[cropWorksheet];
        //            ExcelWorksheet regionCropsWorksheet = package.Workbook.Worksheets[regionCropWorksheet];

        //            //Get all Available Regions
        //            if (regionsWorksheet.Dimension != null)
        //            {
        //                var start = regionsWorksheet.Dimension.Start;
        //                var end = regionsWorksheet.Dimension.End;

        //                var allRegions = new List<Region>();

        //                for (int row = start.Row +1; row <= end.Row; row++)
        //                {
        //                    var region = new Region()
        //                    {
        //                        Id = Convert.ToInt32(regionsWorksheet.Cells[row, 1].Text),
        //                        Name = regionsWorksheet.Cells[row, 2].Text,
        //                        Units = regionsWorksheet.Cells[row, 3].Text
        //                    };
        //                    allRegions.Add(region);
        //                }
        //                database.Regions = allRegions;
        //            }

        //            //Get all Available Crops
        //            if (cropsWorksheet.Dimension != null)
        //            {
        //                var start = cropsWorksheet.Dimension.Start;
        //                var end = cropsWorksheet.Dimension.End;

        //                var allCrops = new List<Crop>();

        //                for (int row = start.Row +1; row <= end.Row; row++)
        //                {
        //                    var crop = new Crop()
        //                    {
        //                       // Id = Convert.ToInt32(cropsWorksheet.Cells[row, 1].Text),
        //                        Name = cropsWorksheet.Cells[row, 2].Text,
        //                    };
        //                    allCrops.Add(crop);
        //                }
        //                database.Crops = allCrops;
        //            }

        //            //Get all Available Regions with their Crops
        //            if (regionCropsWorksheet.Dimension != null)
        //            {
        //                var start = regionCropsWorksheet.Dimension.Start;
        //                var end = regionCropsWorksheet.Dimension.End;

        //                var allRegionCrops = new List<RegionCrop>();

        //                for (int row = start.Row +1; row <= end.Row; row++)
        //                {
        //                    var regionCrop = new RegionCrop()
        //                    {
        //                        //Id = regionCropsWorksheet.Cells[row, 1].Text,
        //                        RegionId = Convert.ToInt32(regionCropsWorksheet.Cells[row, 2].Text),
        //                        Crop = regionCropsWorksheet.Cells[row, 3].Text
        //                    };
        //                    allRegionCrops.Add(regionCrop);
        //                }
        //                database.RegionCrops = allRegionCrops;
        //            }

        //            package.Dispose();
        //        }
        //        newTempDatabaseFile.Delete();
        //        return database;
        //    }
        //    catch (Exception ex)
        //    {
        //        //Catch exception
        //        newTempDatabaseFile.Delete();
        //        throw ex;
               
        //    }

        //}

        //public static bool DatabaseUpdate(string newSpreadsheet,Region newRegion, Database currentDatabase)
        //{
        //    try
        //    {
        //        using (ExcelPackage package = new ExcelPackage(databaseFile))
        //        {
        //            ExcelWorksheet regionsWorksheet = package.Workbook.Worksheets[regionWorksheet];
        //            ExcelWorksheet cropsWorksheet = package.Workbook.Worksheets[cropWorksheet];
        //            ExcelWorksheet regionCropsWorksheet = package.Workbook.Worksheets[regionCropWorksheet];
        //            ExcelWorksheet updatesWorksheet = package.Workbook.Worksheets[updateWorksheet];
                   
        //            //Save New Region
        //            var regionsLastRow = regionsWorksheet.Dimension !=null? regionsWorksheet.Dimension.End.Row:1;
        //            regionsWorksheet.Cells["A" + (regionsLastRow + 1)].Value = newRegion.Id;
        //            regionsWorksheet.Cells["B" + (regionsLastRow + 1)].Value = newRegion.Name;
        //            regionsWorksheet.Cells["C" + (regionsLastRow + 1)].Value = newRegion.Units;

        //            //Save any New Crops
        //            var newRegionCrops= GetRegionCrops(newSpreadsheet);
        //            if (newRegionCrops.Count != 0)
        //            { 
        //                List<Crop> newCrops = newRegionCrops.Where(newRegionCrop => !currentDatabase.Crops.Select(a => a.Name).Contains(newRegionCrop.Name)).ToList();

        //                if (newCrops.Count != 0)
        //                {
        //                    var cropsLastRow = cropsWorksheet.Dimension !=null? cropsWorksheet.Dimension.End.Row:1;
        //                    foreach (Crop newCrop in newCrops)
        //                    {
        //                        cropsWorksheet.Cells["B" + (cropsLastRow + 1)].Value = newCrop.Name;
        //                        cropsLastRow++;
        //                    }
        //                }
        //                //Save Region_Crop Entries
        //                var regionCropsLastRow = regionCropsWorksheet.Dimension != null? regionCropsWorksheet.Dimension.End.Row:1;
        //                foreach (var regionCrop in newRegionCrops)
        //                {
        //                    regionCropsWorksheet.Cells["B" + (regionCropsLastRow + 1)].Value = newRegion.Id;
        //                    regionCropsWorksheet.Cells["C" + (regionCropsLastRow + 1)].Value = regionCrop.Name;
        //                    regionCropsLastRow++;
        //                }
        //            }
                    
        //            //Update the last date modified
        //            var updateLastRow = updatesWorksheet.Dimension!=null? updatesWorksheet.Dimension.End.Row:1;
        //             updatesWorksheet.Cells["A" + (updateLastRow + 1)].Value = DateTime.Now.ToString();

        //            package.Save();
        //            package.Dispose();

        //            return true;
        //        }
        //    }
        //    catch (Exception)
        //    {
        //        return false;
        //        throw;
        //    }
        //}

        public static bool DatabaseUpdate( Region region)
        {
            try
            {
                using (ExcelPackage package = new ExcelPackage(databaseFile))
                {
                    ExcelWorksheet regionsWorksheet = package.Workbook.Worksheets[regionWorksheet];
                    ExcelWorksheet cropsWorksheet = package.Workbook.Worksheets[cropWorksheet];
                    ExcelWorksheet regionCropsWorksheet = package.Workbook.Worksheets[regionCropWorksheet];
                    ExcelWorksheet updatesWorksheet = package.Workbook.Worksheets[updateWorksheet];

                    //Delete Region
                    var regionsLastRow = regionsWorksheet.Dimension !=null?regionsWorksheet.Dimension.End.Row:1;
                    var regionRows = regionsWorksheet.Cells[1,1,regionsLastRow,1].Where(a => a.Text.Equals(region.Id.ToString()));
                    foreach (var row in regionRows.ToList())
                    {
                        regionsWorksheet.DeleteRow(row.Start.Row,1,false);
                    }

                    //Delete Region_Crop Entries
                    var regionsCropsLastRow = regionCropsWorksheet.Dimension!=null?regionCropsWorksheet.Dimension.End.Row:1;
                    var regionCropsRow =  regionCropsWorksheet.Cells[1,2,regionsCropsLastRow,2].Where(a => a.Text.Equals(region.Id.ToString()));
                    foreach (var row in regionCropsRow.ToList())
                    {
                        regionCropsWorksheet.DeleteRow(row.Start.Row, 1, false);
                    }

                    //Update the last date modified
                    var updateLastRow = updatesWorksheet.Dimension!=null?updatesWorksheet.Dimension.End.Row:1;
                    updatesWorksheet.Cells["A" + (updateLastRow + 1)].Value = DateTime.Now.ToString();

                    package.Save();
                    package.Dispose();

                    return true;
                }
            }
            catch (Exception)
            {
                return false;
                throw;
            }
        }

        public static  List<Crop> GetRegionCrops(string newSpreadsheet)
        {
         FileInfo spreadSheet = new FileInfo(newSpreadsheet);

            using (ExcelPackage package = new ExcelPackage(spreadSheet))
            {
                List<Crop> crops = new List<Crop>();

                ExcelWorksheet worksheet = package.Workbook.Worksheets[optimisationWorksheet];

                var searchableCells = worksheet.Cells[16, 2, 23, 2];

                crops.AddRange(searchableCells.Select(cell => new Crop()
                {
                    Name = cell.Text
                }));

                package.Dispose();
                return crops.Count != 0 ? crops.Where(a => !a.Name.Contains("Total") && !a.Name.IsNullOrWhiteSpace()).ToList() : new List<Crop>();
            }
        }
    }
}

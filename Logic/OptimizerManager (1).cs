using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using Optimize;
using System.Web.Script.Serialization;
using OfficeOpenXml;
using Error = Optimize.Error;

namespace Logic
{
    public class OptimizerManager
    {
        static FileInfo databaseFile = new FileInfo(@"C:\inetpub\wwwroot\Temp\Database.xlsm"); //Database file
        static FileInfo newTempDatabaseFile = new FileInfo(@"C:\inetpub\wwwroot\Temp\Database" + DateTime.Now.Ticks + ".xlsm");

        static string regionWorksheet = "Regions";
        static string cropWorksheet = "Crops";
        static string regionCropWorksheet = "Region_Crops";

        public static Database DatabaseCheck()
        {
            try
            {
                databaseFile.CopyTo(newTempDatabaseFile.ToString());

                var database = new Database();

                using (ExcelPackage package = new ExcelPackage(newTempDatabaseFile))
                {
                    ExcelWorksheet regionsWorksheet = package.Workbook.Worksheets[regionWorksheet];
                    ExcelWorksheet cropsWorksheet = package.Workbook.Worksheets[cropWorksheet];
                    ExcelWorksheet regionCropsWorksheet = package.Workbook.Worksheets[regionCropWorksheet];

                    //Get all Available Regions
                    if (regionsWorksheet.Dimension != null)
                    {
                        var start = regionsWorksheet.Dimension.Start;
                        var end = regionsWorksheet.Dimension.End;

                        var allRegions = new List<Region>();

                        for (int row = start.Row + 1; row <= end.Row; row++)
                        {
                            var region = new Region()
                            {
                                Id = Convert.ToInt32(regionsWorksheet.Cells[row, 1].Text),
                                Name = regionsWorksheet.Cells[row, 2].Text,
                                Units = regionsWorksheet.Cells[row, 3].Text
                            };
                            allRegions.Add(region);
                        }
                        database.Regions = allRegions;
                    }

                    //Get all Available Crops
                    if (cropsWorksheet.Dimension != null)
                    {
                        var start = cropsWorksheet.Dimension.Start;
                        var end = cropsWorksheet.Dimension.End;

                        var allCrops = new List<Crop>();

                        for (int row = start.Row + 1; row <= end.Row; row++)
                        {
                            var crop = new Crop()
                            {
                                Id = Convert.ToInt32(cropsWorksheet.Cells[row, 1].Text),
                                Name = cropsWorksheet.Cells[row, 2].Text,
                            };
                            allCrops.Add(crop);
                        }
                        database.Crops = allCrops;
                    }

                    //Get all Available Regions with their Crops
                    if (regionCropsWorksheet.Dimension != null)
                    {
                        var start = regionCropsWorksheet.Dimension.Start;
                        var end = regionCropsWorksheet.Dimension.End;

                        var allRegionCrops = new List<RegionCrop>();

                        for (int row = start.Row + 1; row <= end.Row; row++)
                        {
                            var regionCrop = new RegionCrop()
                            {
                                //Id = regionCropsWorksheet.Cells[row, 1].Text,
                                RegionId = Convert.ToInt32(regionCropsWorksheet.Cells[row, 2].Text),
                                Crop = regionCropsWorksheet.Cells[row, 3].Text
                            };
                            allRegionCrops.Add(regionCrop);
                        }
                        database.RegionCrops = allRegionCrops;
                    }

                    package.Dispose();
                }
                newTempDatabaseFile.Delete();
                return database;
            }
            catch (Exception ex)
            {
                //Catch exception
                newTempDatabaseFile.Delete();
                throw ex;
               
            }

        } 
    }
}

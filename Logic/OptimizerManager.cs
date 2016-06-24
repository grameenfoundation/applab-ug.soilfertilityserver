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
        static string regionWorksheet = "Regions";
        static string cropWorksheet = "Crops";
        static string regionCropWorksheet = "Region_Crops";
        static string updateWorksheet = "Update";
        static string optimisationWorksheet = "Fertilizer Optimization";

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

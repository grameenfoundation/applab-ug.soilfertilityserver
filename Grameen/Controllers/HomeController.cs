using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Grameen.Models;
using Logic;
using Optimize;

namespace Grameen.Controllers
{
    public class HomeController : Controller
    {
        private ApplicationDbContext database = new ApplicationDbContext();

        public ActionResult Index()
        {
            ViewBag.Title = "Home Page";
             
            var report = new List<RegionCropView>();

            if (database.RegionCrops != null)
            {
                report = database.RegionCrops.GroupBy(a => a.RegionId).Select(a =>
                    new RegionCropView
                    {
                        Id = database.Regions.FirstOrDefault(b => b.Id == a.Key).Id,
                        Region = database.Regions.FirstOrDefault(b => b.Id == a.Key).Name,
                        Units = database.Regions.FirstOrDefault(b => b.Id == a.Key).Units,
                        Crops = a.Where(b => b.RegionId == a.Key).Select(x => x.Crop)
                        //database.Crops.Where(z=> a.Select(x => x.Crop).ToList().Contains(z.Id)).Select(b=>b.Name)
                    }).ToList();
            }

            return View(report);
        }

        public ActionResult AddRegion()
        {
            return View();
        }

        [HttpPost]
        public ActionResult AddRegion(Region newRegion, HttpPostedFileBase spreadSheet)
        {
            try
            {
                if (spreadSheet.ContentLength > 0)
                {
                    //var database = OptimizerManager.DatabaseCheck();

                    if (!database.Regions.Any(a => a.Name.Equals(newRegion.Name)))
                    {
                        int newRegionId =database.Regions.Count()!=0? database.Regions.Last().Id  + 1 : 0;
                        newRegion.Id = newRegionId;
                        string fileName = "Optimizer_" + newRegionId + ".xlsm";

                        string directory = new FileInfo(@"C:\inetpub\wwwroot\Temp\").DirectoryName;

                        if (!Directory.Exists(directory))
                        {
                            Directory.CreateDirectory(directory);
                        }

                        string path = Path.Combine(directory, fileName);
                        spreadSheet.SaveAs(path);
                        ////Update database with the new entries
                        //Save new Regions
                        database.Regions.Add(new Region(){Id = newRegionId,Name = newRegion.Name,Units = newRegion.Units});

                        var newRegionCrops = OptimizerManager.GetRegionCrops(path);
                        if (newRegionCrops.Count != 0)
                        {
                            List<Crop> newCrops = newRegionCrops.Where(newRegionCrop => !database.Crops.Select(a => a.Name).Contains(newRegionCrop.Name)).ToList();
                            //Save new Crops
                            if (newCrops.Count != 0)
                            {
                                foreach (Crop newCrop in newCrops)
                                {
                                    database.Crops.Add(newCrop);
                                }
                            }

                            //Save Region_Crop Entries
                            
                            foreach (var regionCrop in newRegionCrops)
                            {
                                database.RegionCrops.Add(new RegionCrop(){RegionId = newRegion.Id,Crop = regionCrop.Name});
                            }
                        }

                        database.SaveChanges();
                        database.Dispose();

                        return RedirectToAction("Index");
                    }
                    //Show Model Error message: Region with this name already exists
                    ModelState.AddModelError("", "Region with this name already exists");
                    return View(newRegion);
                }
                //Show Model Error message: Please select a valid excel macro enabled spreadsheet
                ModelState.AddModelError("", "Please select a valid excel macro enabled spreadsheet");
                return View(newRegion);

                return View();
            }
            catch (Exception)
            {
                ModelState.AddModelError("", "Kindly check your entries!");
                database.Dispose();
                return View(newRegion);
            }
        }

        public ActionResult EditRegion()
        {
            return View();
        }

        public ActionResult DeleteRegion(int id = 0)
        {
            //var database = OptimizerManager.DatabaseCheck();
            var region = database.Regions.First(a => a.Id == id);
            return View(region);
        }

        [HttpPost]
        public ActionResult DeleteRegion(Region region)
        {
            OptimizerManager.DatabaseUpdate(region); //Update the spreadsheet with the database entries

            var spreadSheet = new FileInfo(@"C:\inetpub\wwwroot\Temp\Optimizer_" + region.Id + ".xlsm");

            spreadSheet.Delete();

            return RedirectToAction("Index");
        }
    }
}
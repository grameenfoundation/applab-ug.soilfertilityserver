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
                        int newRegionId = database.Regions.Count() != 0 ? database.Regions.ToList().Last().Id + 1 : 0;
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
                        database.Regions.Add(new Region()
                        {
                            Id = newRegionId,
                            Name = newRegion.Name,
                            Units = newRegion.Units
                        });

                        var newRegionCrops = OptimizerManager.GetRegionCrops(path);
                        if (newRegionCrops.Count != 0)
                        {
                            List<Crop> newCrops =
                                newRegionCrops.Where(
                                    newRegionCrop => !database.Crops.Select(a => a.Name).Contains(newRegionCrop.Name))
                                    .ToList();
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
                                database.RegionCrops.Add(new RegionCrop()
                                {
                                    RegionId = newRegion.Id,
                                    Crop = regionCrop.Name
                                });
                            }

                        }
                        //add new last date modified value
                        database.Versions.Add(new Optimize.Version() {DateTime = DateTime.Now});

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
                 
            }
            catch (Exception)
            {
                ModelState.AddModelError("", "Kindly check your entries!");
                database.Dispose();
                return View(newRegion);
            }
        }

        public ActionResult EditRegion(int id = 0)
        {
            var region = database.Regions.First(a => a.Id == id);
            return View(region);
        }

        [HttpPost]
        public ActionResult EditRegion(Region region, HttpPostedFileBase spreadSheet)
        {
            try
            {
                //Only proceed when there is no other region with the same name as the one assigned
                if (!database.Regions.Any(a => a.Name.Equals(region.Name) && a.Id != region.Id))
                {
                    if (spreadSheet != null)
                    {
                        if (spreadSheet.ContentLength > 0)
                        {
                            string fileName = "Optimizer_" + region.Id + ".xlsm";

                            string directory = new FileInfo(@"C:\inetpub\wwwroot\Temp\").DirectoryName;

                            if (!Directory.Exists(directory))
                            {
                                Directory.CreateDirectory(directory);
                            }

                            string path = Path.Combine(directory, fileName);
                            spreadSheet.SaveAs(path);

                            //update the Region Crops database
                            var newRegionCrops = OptimizerManager.GetRegionCrops(path);
                            if (newRegionCrops.Count != 0)
                            {
                                List<Crop> newCrops =
                                    newRegionCrops.Where(
                                        newRegionCrop =>
                                            !database.Crops.Select(a => a.Name).Contains(newRegionCrop.Name))
                                        .ToList();
                                //Save new Crops
                                if (newCrops.Count != 0)
                                {
                                    foreach (Crop newCrop in newCrops)
                                    {
                                        database.Crops.Add(newCrop);
                                    }
                                }
                                //Remove all previous entries of the Region's Crops
                                var previousRegionCrops = database.RegionCrops.Where(a => a.RegionId == region.Id);
                                database.RegionCrops.RemoveRange(previousRegionCrops);

                                //Save Region_Crop Entries

                                foreach (var regionCrop in newRegionCrops)
                                {
                                    database.RegionCrops.Add(new RegionCrop()
                                    {
                                        RegionId = region.Id,
                                        Crop = regionCrop.Name
                                    });
                                }

                            }
                        }
                    }
                    ////Update region in database
                    var editedRegion = database.Regions.FirstOrDefault(a => a.Id == region.Id);
                    editedRegion.Units = region.Units;
                    editedRegion.Name = region.Name;


                    //add new last date modified value
                    database.Versions.Add(new Optimize.Version() {DateTime = DateTime.Now});

                    database.SaveChanges();
                    database.Dispose();
                    return RedirectToAction("Index");
                }
                //Show Model Error message: Region with this name already exists
                ModelState.AddModelError("", "Region with this name already exists");
                return View(region);
            }
            catch (Exception)
            {
                ModelState.AddModelError("", "Kindly check your entries!");
                database.Dispose();
                return View(region);
            }
        }


        public ActionResult DeleteRegion(int id = 0)
        { 
            var region = database.Regions.First(a => a.Id == id);
            return View(region);
        }

        [HttpPost]
        public ActionResult DeleteRegion(Region region)
        {
            //Delete Region and its RegionCrops from the database
            database.Regions.Remove(database.Regions.FirstOrDefault(a=>a.Id==region.Id));
            var regionCrops = database.RegionCrops.Where(a => a.RegionId == region.Id);
            database.RegionCrops.RemoveRange(regionCrops);

            //add new last date modified value
            database.Versions.Add(new Optimize.Version() { DateTime = DateTime.Now });
            database.SaveChanges();

            //Delete the region's spread sheet
            var spreadSheet = new FileInfo(@"C:\inetpub\wwwroot\Temp\Optimizer_" + region.Id + ".xlsm");

            spreadSheet.Delete();

            return RedirectToAction("Index");
        }
    }
}
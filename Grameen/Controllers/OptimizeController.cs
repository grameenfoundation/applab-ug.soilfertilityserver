using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.Http;
using System.Web.Script.Serialization;
using Grameen.Models;
using Logic;
using Optimize;

namespace Grameen.Controllers
{
    public class OptimizeController : ApiController
    {
        private ApplicationDbContext database = new ApplicationDbContext();

        //Get api/values/5
        public Calc Get([FromBody] Calc json)
        {
            if (json == null)
            {
                json = new JavaScriptSerializer().Deserialize<Calc>(
                   //{"AmtAvailable":250000.0,"CalcCrops":[{"Area":23.0,"Crop":{"Name":"Maize"},"Profit":300.0,"Id":0}],"CalcFertilizers":[{"Fertilizer":{"Name":"Urea"},"Price":150,"id":0}],"CalcCropFertilizerRatios":[],"database":{"Crops":[],"RegionCrops":[],"Regions":[],"VersionDateTime":"Feb 1, 1990 12:00:00 AM"},"FarmerName":"Josh","Id":"94d78060-9953-4d5e-8ffb-45db36e4b727","Imei":"000000000000000","Units":"Acres","Region":0}
                    @"{'AmtAvailable':2500000.0,'CalcCrops':[{'Area':20.0,'Crop':{'Name':'Maize'},'Profit':3000.0,'Id':0}],'CalcFertilizers':[{'Fertilizer':{'Name':'Urea'},'Price':5000,'id':0}],'CalcCropFertilizerRatios':[],'database':{'Crops':[],'RegionCrops':[],'Regions':[],'VersionDateTime':'Feb 1, 1990 12:00:00 AM'},'FarmerName':'Josh','Id':'af96c426-94ab-4ec2-bf86-42c342aacdb1','Imei':'000000000000000','Units':'Acres','Region':0}");
            }


            //Check database for version change
            //Check database for version change
            if (database.Versions.ToList().Last().DateTime > json.Database.VersionDateTime)
            {
                json.Database.Regions = database.Regions.ToList();
                json.Database.Crops = database.Crops.ToList();
                var regionCrops = database.RegionCrops.ToList().Select(regionCrop => new RegionCropAndroid()
                {
                    Id = regionCrop.Id, RegionId = regionCrop.RegionId, Crop = database.Crops.FirstOrDefault(a => a.Name == regionCrop.Crop)
                }).ToList();

                json.Database.RegionCrops = regionCrops;
                json.Database.VersionDateTime = new DateTime(); //database.Versions.ToList().Last().DateTime.Date;
            }
            var result = new Optimizer().Optimize(json);
            return result;
        }

        // POST api/values
        [HttpPost]
        public Calc Post([FromBody] Calc json)
        {

            //Check database for version change
            if (database.Versions.ToList().Last().DateTime > json.Database.VersionDateTime)
            {
                json.Database.Regions = database.Regions.ToList();
                json.Database.Crops = database.Crops.ToList();

                var regionCrops = database.RegionCrops.ToList().Select(regionCrop => new RegionCropAndroid()
                {
                    Id = regionCrop.Id,
                    RegionId = regionCrop.RegionId,
                    Crop = database.Crops.FirstOrDefault(a => a.Name == regionCrop.Crop)
                }).ToList();

                json.Database.RegionCrops = regionCrops;
                json.Database.VersionDateTime = new DateTime();//database.Versions.ToList().Last().DateTime.Date;
            }

            return new Optimizer().Optimize(json);
        }
    }
}
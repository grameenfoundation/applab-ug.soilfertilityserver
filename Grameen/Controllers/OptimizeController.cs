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
                    //@"{'AmtAvailable':50000000,'CalcCrops':[{'Area':4.04686,'Crop':{'Name':'Maize'},'Profit':15.0,'Id':0}],'CalcFertilizers':[{'Fertilizer':{'Name':'Urea'},'Price':300,'id':0}],'CalcCropFertilizerRatios':[],'FarmerName':'564654654654654654654654','Id':'testID','Imei':'000000000000000'}");
                    //@"{'Id':'c22fd646-0b89-417a-9fdf-43fe543a330e','File':null,'AmtAvailable':300000,'TotNetReturns':2997836.4390798109,'FarmerName':'joshua','Imei':'000000000000000','Region':0,'Units':'Acres','CalcFertilizers':[{'Id':0,'Fertilizer':{'Name':'Urea'},'Calc':null,'Price':200,'TotalRequired':52.792757777332326}],'CalcCrops':[{'Id':0,'Crop':{'Name':'Maize'},'Calc':null,'Area':3,'Profit':100,'YieldIncrease':276.09827670648059,'NetReturns':141411.6295393146}],'CalcCropFertilizerRatios':[{'Crop':{'Name':'Maize'},'Fert':{'Name':'Urea'},'Amt':10.103056868572388}]}");
                    @"{'AmtAvailable':2500000.0,'CalcCrops':[{'Area':20.0,'Crop':{'Name':'Maize'},'Profit':3000.0,'Id':0}],'CalcFertilizers':[{'Fertilizer':{'Name':'Urea'},'Price':5000,'id':0}],'CalcCropFertilizerRatios':[],'database':{'Crops':[],'RegionCrops':[],'Regions':[],'VersionDateTime':'Feb 1, 1990 12:00:00 AM'},'FarmerName':'Josh','Id':'af96c426-94ab-4ec2-bf86-42c342aacdb1','Imei':'000000000000000','Units':'Acres','Region':0}");
            }


            //Check database for version change
            //Check database for version change
            if (database.Versions.ToList().Last().DateTime > json.Database.VersionDateTime)
            {
                json.Database.Regions = database.Regions.ToList();
                json.Database.Crops = database.Crops.ToList();
                json.Database.RegionCrops = database.RegionCrops.ToList();
                json.Database.VersionDateTime = database.Versions.ToList().Last().DateTime;
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
                json.Database.RegionCrops = database.RegionCrops.ToList();
                json.Database.VersionDateTime = database.Versions.ToList().Last().DateTime;
            }

            return new Optimizer().Optimize(json);
        }
    }
}
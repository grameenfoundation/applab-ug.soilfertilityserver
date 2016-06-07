using System.Web.Http;
using System.Web.Script.Serialization;
using Logic;
using Optimize;

namespace Grameen.Controllers
{
    public class OptimizeController : ApiController
    {
        //Get api/values/5
        public Calc Get([FromBody] Calc json)
        {
            if (json == null)
            {
                json = new JavaScriptSerializer().Deserialize<Calc>(
                    //@"{'AmtAvailable':50000000,'CalcCrops':[{'Area':4.04686,'Crop':{'Name':'Maize'},'Profit':15.0,'Id':0}],'CalcFertilizers':[{'Fertilizer':{'Name':'Urea'},'Price':300,'id':0}],'CalcCropFertilizerRatios':[],'FarmerName':'564654654654654654654654','Id':'testID','Imei':'000000000000000'}");
                    @"{'Id':'c22fd646-0b89-417a-9fdf-43fe543a330e','File':null,'AmtAvailable':300000,'TotNetReturns':2997836.4390798109,'FarmerName':'joshua','Imei':'000000000000000','Region':0,'Units':'Acres','CalcFertilizers':[{'Id':0,'Fertilizer':{'Name':'Urea'},'Calc':null,'Price':200,'TotalRequired':52.792757777332326}],'CalcCrops':[{'Id':0,'Crop':{'Name':'Maize'},'Calc':null,'Area':3,'Profit':100,'YieldIncrease':276.09827670648059,'NetReturns':141411.6295393146}],'CalcCropFertilizerRatios':[{'Crop':{'Name':'Maize'},'Fert':{'Name':'Urea'},'Amt':10.103056868572388}]}");
            }

            var result = new Optimizer().Optimize(json);
            return result;
        }

        // POST api/values
        [HttpPost]
        public Calc Post([FromBody] Calc json)
        {
            //Optimize.;
            return new Optimizer().Optimize(json);
        }
    }
}
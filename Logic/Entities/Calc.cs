using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;

namespace Optimize
{
    public class Calc
    {
         public  Calc(){}

        private List<CalcCrop> calcCrops = new List<CalcCrop>();
        private List<CalcFertilizer> calcFertilizers = new List<CalcFertilizer>();
        private List<CalcCropFertilizerRatio> calcCropFertilizerRatios = new List<CalcCropFertilizerRatio>();

        public Database Database { get; set; } //Store working database for all crops, regions and their respective jurisdiction

        //[Key]
        //public string Id { get; set; }

        public string File { get; set; }

        public Double AmtAvailable { get; set; }

        public Double TotNetReturns { get; set; }
        public String FarmerName { get; set; }

        public String Imei { get; set; }

        public int Region { get; set; }

        public String Units { get; set; }

        public List<CalcFertilizer> CalcFertilizers
        {
            get { return calcFertilizers; }
            set { calcFertilizers = value; }
        }

        public List<CalcCrop> CalcCrops
        {
            get { return calcCrops; }
            set { calcCrops = value; }
        }

        public List<CalcCropFertilizerRatio> CalcCropFertilizerRatios
        {
            get { return calcCropFertilizerRatios; }
            set { calcCropFertilizerRatios = value; }
        }

       
    }
}
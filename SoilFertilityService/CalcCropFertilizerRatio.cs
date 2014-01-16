using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Optimizer
{
    public class CalcCropFertilizerRatio
    {
        private Crop crop;
        private Fertilizer fert;
        Double amt;

        public CalcCropFertilizerRatio() 
        {
        
        }

        public CalcCropFertilizerRatio(Crop crop, Fertilizer fert, Double amt)
        {
            this.crop = crop;
            this.fert = fert;
            this.amt = amt;
        }

        public Crop Crop
        {
            get { return crop; }
            set { crop = value; }
        }

       public Fertilizer Fert
       {
           get { return fert; } set { fert = value; }
       }

       public Double Amt
       {
           get { return amt; } set { amt = value; }
       }
    }
}
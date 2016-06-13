using System;
using System.ComponentModel.DataAnnotations;

namespace Optimize
{
    public class CalcCropFertilizerRatio
    {
        public CalcCropFertilizerRatio()
        {
           
        }
        public CalcCropFertilizerRatio(Crop crop, Fertilizer fert, Double amt)
        {
            this.Crop = crop;
            this.Fert = fert;
            this.Amt = amt;
        }

        [Key]
        public int Id { get; set; }

        public Crop Crop { get; set; }

        public Fertilizer Fert { get; set; }

        public Double Amt { get; set; }
    }
}
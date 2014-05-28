using System;
using System.Collections.Generic;
using System.Collections;
using System.Linq;
using System.Web;

namespace Optimizer
{
    public class Calc
    {
        private string id, file;
        private Double amtAvailable, totNetReturns;
        private string farmerName, imei;

        private List<CalcCrop> calcCrops = new List<CalcCrop>();
        private List<CalcFertilizer> calcFertilizers = new List<CalcFertilizer>();
        private List<CalcCropFertilizerRatio> calcCropFertilizerRatios = new List<CalcCropFertilizerRatio>();

        public Calc()
        {

        }

        public Calc(String id, String farmerName, String imei, int amtAvailable)
        {
            this.id = id;
            this.farmerName = farmerName;
            this.imei = imei;
            this.amtAvailable = amtAvailable;
        }

        public string Id 
        { 
            get { return id; } set { id = value; }
        }

        public string File
        {
            get { return file; }
            set { file = value; }
        }

        public Double AmtAvailable
        {
            get { return amtAvailable; } set { amtAvailable = value; }
        }

        public Double TotNetReturns
        {
            get { return totNetReturns; } set { totNetReturns = value; }
        }

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

        public String FarmerName
        {
            get { return farmerName; }
            set { farmerName = value; }
        }

        public String Imei
        {
            get { return imei; }
            set { imei = value; }
        }
    }
}
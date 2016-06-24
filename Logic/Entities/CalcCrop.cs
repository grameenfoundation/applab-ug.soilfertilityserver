using System;
using System.ComponentModel.DataAnnotations;

namespace Optimize
{
    public class CalcCrop
    {
        public  CalcCrop(){}

        //[Key]
        //public int Id { get; set; }

        public Crop Crop { get; set; }

        public Calc Calc { get; set; }

        public Double Area { get; set; }

        public double Profit { get; set; }

        public Double YieldIncrease { get; set; }

        public Double NetReturns { get; set; }
    }
}
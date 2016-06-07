using System;

namespace Optimize
{
    public class CalcFertilizer
    {
        public CalcFertilizer()
        {
        }

        public CalcFertilizer(Fertilizer fertilizer, double price)
        {
            this.Fertilizer = fertilizer;
            this.Price = price;
        }

        public int Id { get; set; }

        public Fertilizer Fertilizer { get; set; }

        public Calc Calc { get; set; }

        public double Price { get; set; }

        public Double TotalRequired { get; set; }
    }
}
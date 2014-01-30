using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Optimizer
{
    public class CalcFertilizer
    {
        private int id;
        private Fertilizer fertilizer;
        private Calc calc;

        private int price;
        private Double totalRequired;

        public CalcFertilizer()
        {

        }

        public CalcFertilizer(Fertilizer fertilizer, int price)
        {
            this.fertilizer = fertilizer;
            this.price = price;
        }

        public int Id
        {
            get { return id; }
            set { id = value; }
        }

        public Fertilizer Fertilizer
        {
            get { return fertilizer; }
            set { fertilizer = value; }
        }

        public Calc Calc
        {
            get { return calc; }
            set { calc = value; }
        }


        public int Price
        {
            get { return price; }
            set { price = value; }
        }

        public Double TotalRequired
        {
            get { return totalRequired; }
            set { totalRequired = value; }
        }
    }
}
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Optimizer
{
    public class CalcCrop
    {
        private int id;
        private Crop crop;
        private Calc calc;
        private Double profit;
        private Double area, yieldIncrease, netReturns;

        public CalcCrop()
        {

        }

        public CalcCrop(Crop crop, int area, int profit)
        {
            this.Area = area;
            this.crop = crop;
            this.profit = profit;
        }

        public int Id
        {
            get { return id; }
            set { id = value; }
        }

        public Crop Crop
        {
            get { return crop; }
            set { crop = value; }
        }

        public Calc Calc
        {
            get { return calc; }
            set { calc = value; }
        }

        public Double Area
        {
            get { return area; }
            set { area = value; }
        }

        public Double Profit
        {
            get { return profit; }
            set { profit = value; }
        }

        public Double YieldIncrease
        {
            get { return yieldIncrease; }
            set { yieldIncrease = value; }
        }

        public Double NetReturns
        {
            get { return netReturns; }
            set { netReturns = value; }
        }

    }
}
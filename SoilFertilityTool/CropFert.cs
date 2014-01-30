using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace MyTestService
{
    public class CropFert
    {
        String crop, fert;
        int amt;

        public CropFert() 
        {
        
        }

        public string Name
        {
            get { return crop; } set { crop = value; }
        }

        public String Fert
        {
            get { return fert; } set { fert = value; }
        }

        public int Amt
        {
            get { return amt; } set { amt = value; }
        }
    }
}
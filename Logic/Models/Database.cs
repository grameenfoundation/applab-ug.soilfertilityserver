using System;
using System.Collections.Generic;

namespace Optimize
{
    public class Database
    {
        public List<Region> Regions  { get; set; }

        public List<Crop> Crops { get; set; }

        public List<RegionCrop> RegionCrops { get; set; }

    }

}
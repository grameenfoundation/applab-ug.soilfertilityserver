using System;
using System.Collections.Generic;

namespace Optimize
{
    public class Database
    { 
        public int Id { get; set; }

        public DateTime VersionDateTime { get; set; }

        public List<Region> Regions  { get; set; }

        public List<Crop> Crops { get; set; }

        public List<RegionCropAndroid> RegionCrops { get; set; }

    }

    public class RegionCropAndroid
    {
        public int Id { get; set; }

        public int RegionId { get; set; }

        public Crop Crop { get; set; }
    }
}
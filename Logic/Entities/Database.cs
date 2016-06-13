using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;

namespace Optimize
{
    public class Database
    {
        [Key]
        public int Id { get; set; }

        public List<Region> Regions  { get; set; }

        public List<Crop> Crops { get; set; }

        public List<RegionCrop> RegionCrops { get; set; }

    }

}
using System.Collections.Generic;

namespace Grameen.Models
{
    public class RegionCropView
    {
        public int Id { get; set; }

        public string Region { get; set; }

        public string Units { get; set; }

        public IEnumerable<string> Crops { get; set; }
    }
}
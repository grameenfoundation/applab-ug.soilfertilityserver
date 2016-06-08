using System.ComponentModel.DataAnnotations;

namespace Optimize
{
    public class RegionCrop
    {
        //public string Id { get; set; }

        //public Region Region { get; set; }

        //public Crop Crop { get; set; }
        [Key]
        public int  Id { get; set; }

        public int RegionId { get; set; }

        public string Crop { get; set; }
    }

}
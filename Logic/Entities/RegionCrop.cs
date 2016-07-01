using System;
using System.ComponentModel.DataAnnotations;

namespace Optimize
{
    public class RegionCrop
    { 
        [Key]
        public int  Id { get; set; }

        public int RegionId { get; set; }

        public String Crop { get; set; }
    }

}
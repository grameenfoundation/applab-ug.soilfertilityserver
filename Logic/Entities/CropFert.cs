using System;
using System.ComponentModel.DataAnnotations;

namespace MyTestService
{
    public class CropFert
    {
        public  CropFert(){}
        [Key]
        public int Id { get; set; }

        public string Name { get; set; }

        public String Fert { get; set; }

        public double Amt { get; set; }
    }
}
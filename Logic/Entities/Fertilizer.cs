using System.ComponentModel.DataAnnotations;

namespace Optimize
{
    public class Fertilizer
    {
        public  Fertilizer(){}
        [Key]
        public string Name { get; set; }
    }
}
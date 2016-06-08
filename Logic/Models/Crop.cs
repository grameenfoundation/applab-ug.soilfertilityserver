using System.ComponentModel.DataAnnotations;

namespace Optimize
{
    public class Crop
    {
        //public int Id { get; set; }
        [Key]
        public string Name { get; set; }
    }

}
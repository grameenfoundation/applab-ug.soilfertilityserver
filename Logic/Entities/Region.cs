using System.ComponentModel.DataAnnotations;

namespace Optimize
{
    public class Region
    {
        [Key]
        public int Id { get; set; }

        [Required]
        public string Units { get; set; }

        [Required]
        public string Name { get; set; }
    }

}
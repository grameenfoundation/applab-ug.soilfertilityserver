using System;
using System.ComponentModel.DataAnnotations;

namespace Optimize
{
    public class Version
    {
        [Key]
        public int Id { get; set; }
 
        public DateTime DateTime { get; set; }
    }

}
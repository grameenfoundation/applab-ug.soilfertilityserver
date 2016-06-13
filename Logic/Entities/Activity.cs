using System;
using System.ComponentModel.DataAnnotations;

namespace Optimize
{
    public class Activity
    {
        [Key]
        public int Id { get; set; }

        public DateTime DateTime { get; set; }

        public string Calculation { get; set; }
    }

}
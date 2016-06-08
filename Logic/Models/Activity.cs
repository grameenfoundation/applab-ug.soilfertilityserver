using System;
using System.ComponentModel.DataAnnotations;

namespace Optimize
{
    public class Activity
    {
        [Key]
        public int Id { get; set; }

        public DateTime dateTime { get; set; }

        public Calc calculation { get; set; }
    }

}
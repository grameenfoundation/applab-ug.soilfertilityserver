using System;
using System.ComponentModel.DataAnnotations;

namespace Optimize
{
    public class Error
    {
        [Key]
        public int Id { get; set; }

        public DateTime DateTime { get; set; }

        public string  Calculation { get; set; }

        public string error { get; set; }

    }

}
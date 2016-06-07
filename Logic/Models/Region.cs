using System;
using System.Runtime.CompilerServices;
using OfficeOpenXml.FormulaParsing.Utilities;
using System.ComponentModel.DataAnnotations;

namespace Optimize
{
    public class Region
    {
        public int Id { get; set; }

        [Required]
        public string Units { get; set; }

        [Required]
        public string Name { get; set; }
    }

}
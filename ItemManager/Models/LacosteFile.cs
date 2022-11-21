using Microsoft.AspNetCore.Http;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Threading.Tasks;

namespace ItemManager.Models
{
    public class LacosteFile
    {
        [NotMapped] //Tell Entity Framework to ignore property
        public IFormFile File { get; set; }
    }
}

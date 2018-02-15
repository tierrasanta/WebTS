using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace WebTS2.Models
{
    public class DetailsProrrateoVM
    {
        public string DescripcionActividad { get; set; }
        public double monto { get; set; }
        public DateTime fechaingreso { get; set; }
    }
}
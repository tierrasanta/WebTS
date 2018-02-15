using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace WebTS2.Models
{
    public class ActividadesProrrateoVM
    {
        public string idlote { get; set; }
        public string descripcion { get; set; }
        public Nullable<decimal> costo { get; set; }
    }
}
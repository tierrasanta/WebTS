using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace WebTS2.Models
{
    public class DetalleCultivoVM
    {
        public int IdActividad { get; set; }
        public int IdCabecera { get; set; }
        public string NombreActividad { get; set; }
        public List<CultivoDetalle> detalle { get; set; }
    }
}
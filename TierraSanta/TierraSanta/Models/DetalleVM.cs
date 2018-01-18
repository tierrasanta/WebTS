using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace TierraSanta.Models
{
    public class DetalleVM
    {
        public int IdActividad { get; set; }
        public int IdCabecera { get; set; }
        public string NombreActividad { get; set; }
        public List<PlantillaCultivoDetalle> detalle { get; set; }
    }
}
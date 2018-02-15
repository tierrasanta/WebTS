using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace WebTS2.Models
{
    public class indexprorrateoVM
    {
        public string idempresa { get; set; }
        public int idplantilla { get; set; }
        public int idplantilladetalle { get; set; }
        public int idactividad { get; set; }
        public string idusuario { get; set; }
        public decimal cantidad { get; set; }
        public System.DateTime fechacreacion { get; set; }
        public Nullable<System.DateTime> fechacambio { get; set; }
        public Nullable<decimal> costo { get; set; }

        public virtual PlantillaCultivoCabecera PlantillaCultivoCabecera { get; set; }
        public virtual TablaActividades TablaActividades { get; set; }
    }
}
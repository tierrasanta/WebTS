using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace WebTS2.Models.Validacion
{
    public class CultivoValidacion
    {
        [Display(Name = "Fundo")]
        public string idfundo { get; set; }

        [Display(Name = "Lote")]
        public string idlote { get; set; }

        [Display(Name = "Plantilla")]
        public int idplantilla { get; set; }

        [Display(Name = "Area")]
        public Nullable<decimal> area { get; set; }

        [Display(Name = "Fecha inicial")]
        public System.DateTime fechainicio { get; set; }

        [Display(Name = "Fecha final")]
        public System.DateTime fechafin { get; set; }

        [Display(Name = "Fecha de creación")]
        public System.DateTime fechacreacion { get; set; }
    }
}
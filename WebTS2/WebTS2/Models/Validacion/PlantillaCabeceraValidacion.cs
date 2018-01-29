using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace WebTS2.Models.Validacion
{
    public class PlantillaCabeceraValidacion
    {
        [Required]
        [Display(Name = "Descripción")]
        public string descripcion { get; set; }

        [Display(Name = "Fecha de creación")]
        public System.DateTime fechacreacion { get; set; }

        [Display(Name = "Fecha de cambio")]
        public Nullable<System.DateTime> fechacambio { get; set; }
    }
}
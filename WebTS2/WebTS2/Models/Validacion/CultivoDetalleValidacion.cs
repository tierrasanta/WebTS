using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace WebTS2.Models.Validacion
{
    public class CultivoDetalleValidacion
    {
        [Display(Name = "Actividad")]
        public int idactividad { get; set; }

        [Display(Name = "Cantidad")]
        public decimal cantidad { get; set; }

        [Required]
        [Display(Name = "Fecha de actividad")]
        public System.DateTime fechallenado { get; set; }
    }
}
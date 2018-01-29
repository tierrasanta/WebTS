using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace WebTS2.Models.Validacion
{
    public class EmpresaValidacion
    {
        [Required]
        [Display(Name = "RUC")]
        public string ruc { get; set; }

        [Required]
        [Display(Name = "Razón social")]
        public string razonsocial { get; set; }

        [Required]
        [Display(Name = "Dirección")]
        public string direccion { get; set; }

        [Required]
        [Display(Name = "Abreviatura")]
        public string Abreviatura { get; set; }
        
        [Display(Name = "Fecha de creación")]
        public System.DateTime fechainicio { get; set; }
        
        [Display(Name = "Fecha de cambio")]
        public Nullable<System.DateTime> fechacambio { get; set; }
    }
}
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace WebTS2.Models.Validacion
{
    public class CuentaContableValidacion
    {
        [Required]
        [Display(Name = "Número de cuenta")]
        public string cuenta { get; set; }

        [Required]
        [Display(Name = "Cuenta")]
        public string descripcion { get; set; }
    }
}
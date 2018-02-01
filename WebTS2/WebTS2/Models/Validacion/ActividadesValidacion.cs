using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace WebTS2.Models.Validacion
{
    public class ActividadesValidacion
    {
        [Required]
        [Display(Name = "Descripción")]
        public string descripcion { get; set; }

        [Required]
        [Display(Name = "Abreviatura")]
        public string abreviatura { get; set; }

        [Display(Name = "Unidad de medida")]
        public int unimedida { get; set; }

        [Required]
         [Display(Name = "Costo")]
        public decimal costo1 { get; set; }

        [Display(Name = "Fecha de creación")]
        public System.DateTime fechacreacion { get; set; }

        [Display(Name = "Fecha de cambio")]
        public Nullable<System.DateTime> fechacambio { get; set; }

        [Required]
        [Display(Name = "Prorrateo")]
        public bool prorrateo { get; set; }
    }
}
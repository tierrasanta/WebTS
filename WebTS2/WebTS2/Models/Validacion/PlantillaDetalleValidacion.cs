﻿using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace WebTS2.Models.Validacion
{
    public class PlantillaDetalleValidacion
    {
        [Display(Name = "Actividad")]
        public int idactividad { get; set; }

        [Required]
        [Display(Name = "Cantidad")]
        public decimal cantidad { get; set; }
    }
}
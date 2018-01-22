using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace WebTS2.Models.Validacion
{
    //[MetadataType(typeof(UsuarioValidacion))]
    public class UsuarioValidacion
    {
        [Required]
        [Display(Name = "Nombre y apellidos")]
        public string Nombre { get; set; }

        [Required]
        [Display(Name = "Nombre de usuario")]
        public string Login { get; set; }

        [Required]
        [DataType(DataType.Password)]
        [Display(Name = "Clave de ingreso")]
        public string Clave { get; set; }

        [Required]
        [DataType(DataType.EmailAddress)]
        [Display(Name = "Email")]
        public string Email { get; set; }

        [Display(Name = "Estado(Activo/Desactivo)")]
        public bool Estado { get; set; }

        [Display(Name = "Es Administrador")]
        public bool EsAdministrador { get; set; }
    }

}
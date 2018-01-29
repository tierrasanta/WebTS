//------------------------------------------------------------------------------
// <auto-generated>
//     Este código se generó a partir de una plantilla.
//
//     Los cambios manuales en este archivo pueden causar un comportamiento inesperado de la aplicación.
//     Los cambios manuales en este archivo se sobrescribirán si se regenera el código.
// </auto-generated>
//------------------------------------------------------------------------------

namespace WebTS2.Models
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using Validacion;

    [MetadataType(typeof(CultivoDetalleValidacion))]
    public partial class CultivoDetalle
    {
        public string idempresa { get; set; }
        public string idusuario { get; set; }
        public int idcultivo { get; set; }
        public int idcultivodetalle { get; set; }
        public int idactividad { get; set; }
        public decimal cantidad { get; set; }
        public System.DateTime fechallenado { get; set; }
        public System.DateTime fechacreacion { get; set; }
        public Nullable<System.DateTime> fechacambio { get; set; }
    
        public virtual Cultivo Cultivo { get; set; }
        public virtual TablaActividades TablaActividades { get; set; }
    }
}

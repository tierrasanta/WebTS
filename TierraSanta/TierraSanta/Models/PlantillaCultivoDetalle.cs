//------------------------------------------------------------------------------
// <auto-generated>
//     Este código se generó a partir de una plantilla.
//
//     Los cambios manuales en este archivo pueden causar un comportamiento inesperado de la aplicación.
//     Los cambios manuales en este archivo se sobrescribirán si se regenera el código.
// </auto-generated>
//------------------------------------------------------------------------------

namespace TierraSanta.Models
{
    using System;
    using System.Collections.Generic;
    
    public partial class PlantillaCultivoDetalle
    {
        public string idempresa { get; set; }
        public int idplantilla { get; set; }
        public int idplantilladetalle { get; set; }
        public int idactividad { get; set; }
        public string idusuario { get; set; }
        public decimal cantidad { get; set; }
        public System.DateTime fechacreacion { get; set; }
        public Nullable<System.DateTime> fechacambio { get; set; }
    
        public virtual PlantillaCultivoCabecera PlantillaCultivoCabecera { get; set; }
        public virtual TablaActividades TablaActividades { get; set; }
    }
}

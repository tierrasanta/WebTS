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
    
    public partial class LOTES
    {
        public string idempresa { get; set; }
        public string idfinca { get; set; }
        public string idlote { get; set; }
        public string idusuario { get; set; }
        public string descripcion { get; set; }
        public decimal area { get; set; }
        public System.DateTime fechacreacion { get; set; }
        public Nullable<System.DateTime> fechacambio { get; set; }
    
        public virtual FINCAS FINCAS { get; set; }
    }
}

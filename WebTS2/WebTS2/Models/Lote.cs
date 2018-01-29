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

    [MetadataType(typeof(LoteValidacion))]
    public partial class Lote
    {
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2214:DoNotCallOverridableMethodsInConstructors")]
        public Lote()
        {
            this.Cultivo = new HashSet<Cultivo>();
        }
    
        public string idempresa { get; set; }
        public string idfundo { get; set; }
        public string idlote { get; set; }
        public string idusuario { get; set; }
        public string descripcion { get; set; }
        public decimal area { get; set; }
        public System.DateTime fechacreacion { get; set; }
        public Nullable<System.DateTime> fechacambio { get; set; }
    
        public virtual Fundo Fundo { get; set; }
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<Cultivo> Cultivo { get; set; }
    }
}

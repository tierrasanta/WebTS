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
    
    public partial class TABLAS_CULTIVOS
    {
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2214:DoNotCallOverridableMethodsInConstructors")]
        public TABLAS_CULTIVOS()
        {
            this.CULTIVOS = new HashSet<CULTIVOS>();
            this.PLANTILLAS_CULTIVOS = new HashSet<PLANTILLAS_CULTIVOS>();
            this.TABLAS_ACTIVIDADES = new HashSet<TABLAS_ACTIVIDADES>();
        }
    
        public string idempresa { get; set; }
        public string idcodigo { get; set; }
        public string idusuario { get; set; }
        public string descripcion { get; set; }
        public string abreviatura { get; set; }
        public decimal valornum { get; set; }
        public string valorchar { get; set; }
        public System.DateTime fechacreacion { get; set; }
        public Nullable<System.DateTime> fechacambio { get; set; }
    
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<CULTIVOS> CULTIVOS { get; set; }
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<PLANTILLAS_CULTIVOS> PLANTILLAS_CULTIVOS { get; set; }
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<TABLAS_ACTIVIDADES> TABLAS_ACTIVIDADES { get; set; }
    }
}

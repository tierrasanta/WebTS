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
    
    public partial class TablaCultivos
    {
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2214:DoNotCallOverridableMethodsInConstructors")]
        public TablaCultivos()
        {
            this.TablaActividades = new HashSet<TablaActividades>();
        }
    
        public string idempresa { get; set; }
        public int pktablacultivos { get; set; }
        public string idcodigo { get; set; }
        public string idusuario { get; set; }
        public string descripcion { get; set; }
        public string abreviatura { get; set; }
        public decimal valornum { get; set; }
        public string valorchar { get; set; }
        public System.DateTime fechacreacion { get; set; }
        public Nullable<System.DateTime> fechacambio { get; set; }
    
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<TablaActividades> TablaActividades { get; set; }
    }
}

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

    [MetadataType(typeof(CultivoValidacion))]
    public partial class Cultivo
    {
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2214:DoNotCallOverridableMethodsInConstructors")]
        public Cultivo()
        {
            this.CultivoDetalle = new HashSet<CultivoDetalle>();
        }
    
        public string idempresa { get; set; }
        public string idfundo { get; set; }
        public string idlote { get; set; }
        public int idcultivo { get; set; }
        public string idusuario { get; set; }
        public int idplantilla { get; set; }
        public Nullable<decimal> area { get; set; }
        public System.DateTime fechainicio { get; set; }
        public System.DateTime fechafin { get; set; }
        public System.DateTime fechacreacion { get; set; }
        public Nullable<System.DateTime> fechacambio { get; set; }
        public int idtablacultivos { get; set; }
    
        public virtual Fundo Fundo { get; set; }
        public virtual Lote Lote { get; set; }
        public virtual PlantillaCultivoCabecera PlantillaCultivoCabecera { get; set; }
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<CultivoDetalle> CultivoDetalle { get; set; }
        public virtual TablaCultivos TablaCultivos { get; set; }
    }
}

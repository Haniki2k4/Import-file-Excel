//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated from a template.
//
//     Manual changes to this file may cause unexpected behavior in your application.
//     Manual changes to this file will be overwritten if the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

namespace Import_file_Excel.Data
{
    using System;
    using System.Collections.Generic;
    
    public partial class TINH
    {
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2214:DoNotCallOverridableMethodsInConstructors")]
        public TINH()
        {
            this.HUYENs = new HashSet<HUYEN>();
        }
    
        public string MaTinh { get; set; }
        public string MaVung { get; set; }
        public string TenTinh { get; set; }
    
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<HUYEN> HUYENs { get; set; }
        public virtual VUNG VUNG { get; set; }
    }
}

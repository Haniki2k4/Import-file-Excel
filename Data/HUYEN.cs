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
    
    public partial class HUYEN
    {
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2214:DoNotCallOverridableMethodsInConstructors")]
        public HUYEN()
        {
            this.XAs = new HashSet<XA>();
        }
    
        public string MaHuyen { get; set; }
        public string MaTinh { get; set; }
        public string TenHuyen { get; set; }
    
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<XA> XAs { get; set; }
        public virtual TINH TINH { get; set; }
    }
}

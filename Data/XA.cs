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
    
    public partial class XA
    {
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2214:DoNotCallOverridableMethodsInConstructors")]
        public XA()
        {
            this.DIABANs = new HashSet<DIABAN>();
            this.THONGTINHOes = new HashSet<THONGTINHO>();
        }
    
        public string MaXa { get; set; }
        public string MaHuyen { get; set; }
        public string MaTinh { get; set; }
        public string TenXa { get; set; }
    
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<DIABAN> DIABANs { get; set; }
        public virtual HUYEN HUYEN { get; set; }
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<THONGTINHO> THONGTINHOes { get; set; }
        public virtual TINH TINH { get; set; }
    }
}

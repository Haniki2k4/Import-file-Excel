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
    
    public partial class THONGTINHO
    {
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2214:DoNotCallOverridableMethodsInConstructors")]
        public THONGTINHO()
        {
            this.THANHVIENTRONGHOes = new HashSet<THANHVIENTRONGHO>();
            this.THONGTINTUVONGs = new HashSet<THONGTINTUVONG>();
        }
    
        public string MaHo { get; set; }
        public string MaXa { get; set; }
        public string TenDB { get; set; }
        public string HoSo { get; set; }
        public string Nam { get; set; }
        public string HoTenChuHo { get; set; }
        public string DiaChi { get; set; }
        public string MaNV { get; set; }
        public Nullable<System.DateTime> NgayKThuc { get; set; }
        public Nullable<System.DateTime> NgayPVan { get; set; }
        public Nullable<decimal> KinhDo { get; set; }
        public Nullable<decimal> ViDo { get; set; }
        public string SDT { get; set; }
        public string TSNK { get; set; }
        public string TSNAM { get; set; }
        public string TSNU { get; set; }
        public string KT9 { get; set; }
        public string C45 { get; set; }
        public string KT14 { get; set; }
        public string NguoiXN { get; set; }
        public string NguoiTao { get; set; }
        public Nullable<System.DateTime> NgayTao { get; set; }
        public string PhienBan { get; set; }
    
        public virtual DIABAN DIABAN { get; set; }
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<THANHVIENTRONGHO> THANHVIENTRONGHOes { get; set; }
        public virtual XA XA { get; set; }
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<THONGTINTUVONG> THONGTINTUVONGs { get; set; }
    }
}

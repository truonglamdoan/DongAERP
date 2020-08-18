namespace Models.Framework
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("MarKetAsian")]
    public partial class MarKetAsian
    {
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2214:DoNotCallOverridableMethodsInConstructors")]
        public MarKetAsian()
        {
            MarKetDetails = new HashSet<MarKetDetail>();
        }

        public Guid APK { get; set; }

        [StringLength(50)]
        public string MarKetAsianID { get; set; }

        [Required]
        [StringLength(250)]
        public string MarKetAsianName { get; set; }

        public DateTime? CreateDate { get; set; }

        [StringLength(50)]
        public string CreateUserID { get; set; }

        [StringLength(50)]
        public string MarKetID { get; set; }

        public virtual MarKet MarKet { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<MarKetDetail> MarKetDetails { get; set; }
    }
}

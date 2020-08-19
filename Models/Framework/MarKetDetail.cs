namespace Models.Framework
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("MarKetDetail")]
    public partial class MarKetDetail
    {
        public Guid APK { get; set; }

        [StringLength(50)]
        public string MarKetDetailID { get; set; }

        [Required]
        [StringLength(250)]
        public string MarKetDetailName { get; set; }

        public DateTime? CreateDate { get; set; }

        [StringLength(50)]
        public string CreateUserID { get; set; }

        [StringLength(50)]
        public string MarKetID { get; set; }

        [StringLength(50)]
        public string MarKetAsianID { get; set; }

        public virtual MarKetAsian MarKetAsian { get; set; }

        public virtual MarKet MarKet { get; set; }
    }
}

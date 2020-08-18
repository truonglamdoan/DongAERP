namespace Models.Framework
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("Account")]
    public partial class Account
    {
        [Key]
        public Guid APK { get; set; }

        [Required]
        [StringLength(50)]
        public string EmployeeID { get; set; }

        [Required]
        [StringLength(250)]
        public string FullName { get; set; }

        [Required]
        [StringLength(250)]
        public string Password { get; set; }

        [Required]
        [StringLength(50)]
        public string DepartmentID { get; set; }

        public DateTime? BirthDay { get; set; }

        [StringLength(250)]
        public string Address { get; set; }

        [StringLength(100)]
        public string Tel { get; set; }

        [StringLength(100)]
        public string Email { get; set; }

        public byte? IsUserID { get; set; }

        public byte? Disabled { get; set; }

        public DateTime? CreateDate { get; set; }

        [StringLength(50)]
        public string CreateUserID { get; set; }

        public DateTime? LastModifyDate { get; set; }

        [StringLength(50)]
        public string LastModifyUserID { get; set; }
    }
}

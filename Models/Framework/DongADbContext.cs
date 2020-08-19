namespace Models.Framework
{
    using System;
    using System.Data.Entity;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Linq;

    public partial class DongADbContext : DbContext
    {
        public DongADbContext()
            : base("name=DongADbContext")
        {
        }

        public virtual DbSet<Account> Accounts { get; set; }
        public virtual DbSet<MarKetAsian> MarKetAsians { get; set; }
        public virtual DbSet<MarKetDetail> MarKetDetails { get; set; }
        public virtual DbSet<MarKet> MarKets { get; set; }
        public virtual DbSet<sysdiagram> sysdiagrams { get; set; }

        protected override void OnModelCreating(DbModelBuilder modelBuilder)
        {
            modelBuilder.Entity<Account>()
                .Property(e => e.CreateUserID)
                .IsUnicode(false);

            modelBuilder.Entity<Account>()
                .Property(e => e.LastModifyUserID)
                .IsUnicode(false);
        }
    }
}

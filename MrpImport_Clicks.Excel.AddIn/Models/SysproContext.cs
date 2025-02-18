using MrpImport_Clicks.Excel.AddIn.Models;
using System;
using System.ComponentModel.DataAnnotations.Schema;
using System.Data.Entity;
using System.Linq;

namespace MrpImport_Clicks.Excel.AddIn
{
    public partial class SysproContext : DbContext
    {
        public SysproContext()
            : base("name=SysproContext")
        {
        }

        public virtual DbSet<MrpForecast> MrpForecast { get; set; }
        public virtual DbSet<ArCustStkXref> ArCustStkXref { get; set; }

        protected override void OnModelCreating(DbModelBuilder modelBuilder)
        {
            //Database.SetInitializer<SysproContext>(null);
            //modelBuilder.Entity<usr_ForecastImport>().Property(t => t.FCastQtyPer01).HasPrecision(18, 6);
        }
    }
}

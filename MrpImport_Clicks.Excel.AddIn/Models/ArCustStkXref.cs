namespace MrpImport_Clicks.Excel.AddIn.Models
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("ArCustStkXref")]
    public partial class ArCustStkXref
    {
        [Key]
        [Column(Order = 0)]
        [StringLength(15)]
        public string Customer { get; set; }

        [Key]
        [Column(Order = 1)]
        [StringLength(30)]
        public string CustStockCode { get; set; }

        [StringLength(30)]
        public string StockCode { get; set; }

        [StringLength(50)]
        public string Description { get; set; }

        [StringLength(100)]
        public string LongDesc { get; set; }

    }
}

namespace MrpImport_Clicks.Excel.AddIn.Models
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("MrpForecast")]
    public partial class MrpForecast
    {
        [Key]
        [Column(Order = 0)]
        [StringLength(30)]
        public string StockCode { get; set; }

        [Key]
        [Column(Order = 1)]
        [StringLength(10)]
        public string ForecastWh { get; set; }

        [Key]
        [Column(Order = 2)]
        public DateTime ForecastDate { get; set; }

        [Key]
        [Column(Order = 3)]
        public decimal Line { get; set; }

        public decimal ForecastQtyOutst { get; set; }

        [StringLength(50)]
        public string Description { get; set; }

        [StringLength(30)]
        public string Reference { get; set; }

        [StringLength(1)]
        public string InactiveFlag { get; set; }

        public string ForecastType { get; set; }

        [NotMapped]
        public string ClicksStockCode { get; set; }


        //[Required]
        //[StringLength(1)]
        //public string ForecastType { get; set; }

        //[Required]
        //[StringLength(15)]
        //public string Customer { get; set; }

        //[Required]
        //[StringLength(30)]
        //public string ResourceParent { get; set; }
        //public decimal QtyInvoiced { get; set; }

        //public DateTime? OrigForecastDate { get; set; }

        //public decimal OriginalLine { get; set; }

        //public decimal OrigQtyOutst { get; set; }

        //[Required]
        //[StringLength(5)]
        //public string Version { get; set; }

        //[Required]
        //[StringLength(5)]
        //public string Release { get; set; }

    }
}

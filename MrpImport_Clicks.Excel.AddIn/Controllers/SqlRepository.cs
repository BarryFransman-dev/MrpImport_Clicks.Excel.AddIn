using MrpImport_Clicks.Excel.AddIn.Models;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Data.Entity.Migrations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MrpImport_Clicks.Excel.AddIn
{
    public class SqlRepository
    {
        public int ForecastAdd(List<MrpForecast> recs)
        {
            var savedRecs = 0;
            using (var dsContext = new SysproContext())
            {
                foreach (var item in recs)
                {
                    dsContext.MrpForecast.AddOrUpdate(p => new
                    { p.StockCode, p.ForecastWh, p.ForecastDate, p.Line, p.ForecastQtyOutst, p.Reference, p.InactiveFlag }, item);
                    savedRecs += dsContext.SaveChanges();
                }
            }
            return savedRecs;
        }

        public int ForecastAddUpdt(List<MrpForecast> recs)
        {
            var savedRecs = 0;
            //var fDate = DateTime.Now.ToString("yyyy-MM-dd");
            //var dt = DateTime.ParseExact(fDate, "yyyy-MM-dd", System.Globalization.CultureInfo.InvariantCulture);
            //var fdt = dt.Year + "-" + dt.Month + "-" + dt.Day;
            using (var dsContext = new SysproContext())
            {
                foreach (var item in recs)
                {
                    var existingEntity = dsContext.MrpForecast.FirstOrDefault(s => s.StockCode == item.StockCode & 
                                                                                s.ForecastWh == item.ForecastWh & 
                                                                                s.ForecastDate == item.ForecastDate);
                    if (existingEntity != null)
                    {
                        // ✅ Update existing record
                        existingEntity.ForecastQtyOutst = item.ForecastQtyOutst;
                        dsContext.Entry(existingEntity).State = EntityState.Modified;
                    }
                    else
                    {
                        // ✅ Insert new record
                        dsContext.MrpForecast.AddOrUpdate(p => new
                        { p.StockCode, p.ForecastWh, p.ForecastDate, p.Line, p.ForecastQtyOutst, p.Reference, p.InactiveFlag }, item);
                    }

                    savedRecs += dsContext.SaveChanges();
                }
            }
            return savedRecs;
        }

        public List<MrpForecast> GetAlternativeStockCodes(List<MrpForecast> stockCodes)
        {
            var xref = new SysproContext().ArCustStkXref.Where(f => f.LongDesc == "CLICKS").GroupBy(g => new XrefVM { StockCode = g.StockCode, CustStockCode = g.CustStockCode }).ToList();
            var xref2 = xref.Select(x => new XrefVM { StockCode = x.Key.StockCode, CustStockCode = x.Key.CustStockCode }).Distinct().ToList();
            var updatedStockList = stockCodes
                                    .GroupJoin(
                                        xref2,
                                        stock => stock.ClicksStockCode,
                                        alt => alt.CustStockCode,
                                        (stock, altMatches) =>
                                        {
                                            stock.StockCode = altMatches.FirstOrDefault()?.StockCode ?? "N/A";
                                            return stock;
                                        }
                                    ).ToList();
            return updatedStockList;
        }
    }
}

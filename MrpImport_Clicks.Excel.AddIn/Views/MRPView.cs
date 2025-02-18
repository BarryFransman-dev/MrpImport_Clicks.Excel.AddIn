using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Tools.Excel;
using System;
using System.Collections.Generic;
using Application = Microsoft.Office.Interop.Excel.Application;
using System.Windows.Forms;
using MrpImport_Clicks.Excel.AddIn.Models;
using System.Linq;
using System.Deployment.Application;

namespace MrpImport_Clicks.Excel.AddIn
{
    public partial class ForecastImport
    {
        private SqlRepository sqlRepository;
        private Application application;
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            application = Globals.ThisWorkbook.Application;
            sqlRepository = new SqlRepository();
            grpVersion.Label = "Ver:" + System.Windows.Forms.Application.ProductVersion.ToString();
            if (ApplicationDeployment.IsNetworkDeployed)
            {
                var ad = ApplicationDeployment.CurrentDeployment;
                grpVersion.Label = "Ver:" + ad.CurrentVersion.ToString();
            }
        }

        private void btnUpdate_Click(object sender, RibbonControlEventArgs e)
        {
            var sheetName = Globals.ThisWorkbook.Application.ActiveSheet.Name;

            if (sheetName.ToUpper() != "DETAILED")
            {
                MessageBox.Show("Please first select Detailed Tab");
                return;
            }

            if (MessageBox.Show("This will update the details from the " + sheetName + " Tab to Syspro's Forecast table. Continue?",
                "UPDATE\\INSERT", MessageBoxButtons.YesNo) == DialogResult.No)
            {
                return;
            }

            SetCursor("READING RECORDS...");

            try
            {
                var captureDetail = ReadRows();
                SetCursor("UPDATING SYSPRO MRP FORECAST...");

                var noMatch = captureDetail.Where(n => n.StockCode == "N/A").ToList();

                if (noMatch.Count > 0)
                {
                    MessageBox.Show("No records updated. Not all Clicks codes are matched", "UPDATE\\INSERT", MessageBoxButtons.OK);
                    return;
                }

                var res = sqlRepository.ForecastAddUpdt(captureDetail);
                ResetCursor();
                MessageBox.Show("Importing Complete - " + res + " records updated.", "UPDATE\\INSERT", MessageBoxButtons.OK);
            }
            catch (Exception ex)
            {
                ResetCursor();
                MessageBox.Show(@"PostClick: " + ex.Message);
            }
            finally
            {
                ResetCursor();
            }
        }

        private List<MrpForecast> ReadRows()
        {
            {
                Microsoft.Office.Interop.Excel.Worksheet worksheet = Globals.ThisWorkbook.Application.ActiveSheet;
                Range usedRange = worksheet.UsedRange;
                Range visibleCells = usedRange.SpecialCells(XlCellType.xlCellTypeVisible);

                List<MrpForecast> mrpList = new List<MrpForecast>();

                int totalRows = visibleCells.Rows.Count;
                int totalCols = visibleCells.Columns.Count;

                // Find header row containing dates
                int headerRowIndex = 10; // Adjust based on where your dates are

                // Loop through rows
                for (int row = 1; row <= totalRows; row++)
                {
                    Range currentRow = visibleCells.Rows[row];
                    var curRowVals = currentRow.Value;
                    if (curRowVals[1, 11] != null &&
                        curRowVals[1, 11] == "Total Vendor Order Plan")
                    {
                        var skuNumber = curRowVals[1, 5];
                        var skuDesc = curRowVals[1, 6];

                        for (int col = 12; col <= totalCols; col++) // Date columns start from 12th column
                        {
                            var dateCell = DateTime.Parse(usedRange.Cells[headerRowIndex, col].Value);
                            var quantityCell = Convert.ToDecimal(curRowVals[1, col]);

                            if (dateCell != null && quantityCell != null)
                            {
                                MrpForecast data = new MrpForecast
                                {
                                    ClicksStockCode = skuNumber.ToString(),
                                    Description = skuDesc,
                                    ForecastWh = "C",
                                    Line = 1,
                                    Reference = "Clicks",
                                    InactiveFlag = string.Empty,
                                    ForecastDate = dateCell,
                                    ForecastQtyOutst = quantityCell
                                };
                                mrpList.Add(data);
                            }
                        }
                    }
                }
                var retList = sqlRepository.GetAlternativeStockCodes(mrpList);
                CopyToSheet(retList);
                return retList;
            }
        }

        private void CopyToSheet(List<MrpForecast> mrpList)
        {
            // Create new sheet and write extracted data
            try
            {
                var existingSheet = Globals.ThisWorkbook.Application.Worksheets["Extracted Data"];
                Globals.ThisWorkbook.Application.DisplayAlerts = false;
                existingSheet.Delete();
                Globals.ThisWorkbook.Application.DisplayAlerts = true;
            }
            catch (Exception)
            {

            }

            Microsoft.Office.Interop.Excel.Worksheet newSheet = Globals.ThisWorkbook.Application.Sheets.Add();
            newSheet.Name = "Extracted Data";
            int firstRow = 1;

            // Write headers

            //newSheet.Rows[firstRow].Style.Font.Bold = true;

            newSheet.Cells[firstRow, 1].Value = "ClickStockCode";
            newSheet.Cells[firstRow, 2].Value = "ForecastWh";
            newSheet.Cells[firstRow, 3].Value = "ForecastDate";
            newSheet.Cells[firstRow, 4].Value = "Line";
            newSheet.Cells[firstRow, 5].Value = "ForecastQtyOutst";
            newSheet.Cells[firstRow, 6].Value = "Description";
            newSheet.Cells[firstRow, 7].Value = "Reference";
            newSheet.Cells[firstRow, 8].Value = "InactiveFlag";
            newSheet.Cells[firstRow, 9].Value = "SysproStockCode";
            newSheet.Columns[9].Style.Numberformat = "@";

            // Write data to new sheet
            int newRow = 2;
            foreach (var mrp in mrpList)
            {
                newSheet.Cells[newRow, 1].Value = mrp.ClicksStockCode;
                newSheet.Cells[newRow, 2].Value = mrp.ForecastWh;
                newSheet.Cells[newRow, 3].Value = mrp.ForecastDate;
                newSheet.Cells[newRow, 4].Value = mrp.Line;
                newSheet.Cells[newRow, 5].Value = mrp.ForecastQtyOutst;
                newSheet.Cells[newRow, 6].Value = mrp.Description;
                newSheet.Cells[newRow, 7].Value = mrp.Reference;
                newSheet.Cells[newRow, 8].Value = mrp.InactiveFlag;
                newSheet.Cells[newRow, 9].Value = mrp.StockCode;
                newRow++;
            }
            //newSheet.Cells[1, 5].Style.Numberformat = "#,##0";
            //newSheet.Cells[1, 5].Formula = $"=SUBTOTAL(9, E2:E{newRow - 1})";
        }

        private object GetColVal(string colName, Range area, int row)
        {
            var val = area.Find(colName, Type.Missing, XlFindLookIn.xlValues, XlLookAt.xlWhole, XlSearchOrder.xlByColumns, XlSearchDirection.xlNext, false, Type.Missing, Type.Missing);
            var retval = val != null ? (object)area[row, val.Column].Value2 : null;
            return retval;
        }

        private void SetCursor(string displayText)
        {
            this.application.Cursor = XlMousePointer.xlWait;
            this.application.StatusBar = displayText.ToUpper();
            this.application.ScreenUpdating = true;
        }

        private void ResetCursor()
        {
            this.application.Cursor = XlMousePointer.xlDefault;
            this.application.StatusBar = null;
            this.application.ScreenUpdating = true;
        }

    }
}

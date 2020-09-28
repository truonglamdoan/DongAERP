using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using DongA.Bussiness;
using DongA.Entities;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Configuration;
using System.Web.Mvc;

namespace DongAERP.Areas.Admin.Controllers
{
    public class ReportExcelDetailtByPartnerController : Controller
    {
        /// <summary>
        /// Content Type - Excel 2007 trở lên
        /// application/vnd.openxmlformats-officedocument.spreadsheetml.sheet
        /// </summary>
        public const string XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
        /// <summary>
        /// application/pdf
        /// </summary>
        public const string PDF = "application/pdf";

        // GET: Admin/ReportExcelDetailtByPartner
        public ActionResult Index()
        {
            return View();
        }

        /// <summary>
        /// Xuất Excel theo Ngày/Tháng/Năm
        /// </summary>
        /// <param name="fromDate"></param>
        /// <param name="toDate"></param>
        /// <returns></returns>
        public ActionResult CreateExcelForDayMonthYear(DateTime fromDate, DateTime toDate, int typeID, string reportTypeID, string marketID)
        {
            WorkbookDesigner designer = new WorkbookDesigner();
            string templatePath = "~/Content/Report/ReportDetailtByPartner.xlsx";
            // Get đường dẫn
            templatePath = System.Web.HttpContext.Current.Server.MapPath(templatePath);

            designer.Workbook = new DongA.Core.DongAExcel().OpenExcelFile(templatePath);
            designer.Workbook.CalculateFormula();

            // Set the glorbalization setting to change subtotal and grand total names
            GlobalizationSettings gsi = new GlobalizationSettingsImp();
            designer.Workbook.Settings.GlobalizationSettings = gsi;

            WorksheetCollection workSheets = designer.Workbook.Worksheets;
            Worksheet sheetReport = designer.Workbook.Worksheets[0];

            // Tạo title
            string typeReport = string.Empty;
            switch (typeID)
            {
                // Theo ngày
                case 1:
                    typeReport = "Chi tiết - Theo Ngày";
                    break;
                // Theo tháng
                case 2:
                    typeReport = "Chi tiết - Theo Tháng";
                    break;
                // Theo năm
                default:
                    typeReport = "Chi tiết - Theo Năm";
                    break;
            }

            CreateTitle("A2", "K2", sheetReport, typeReport, 14);

            // Get data report ngày
            List<ReportDetailtForTotalMoneyType> listReportData = new List<ReportDetailtForTotalMoneyType>();
            List<ReportDetailtForTotalMoneyType> listReportDataConvertUSD = new List<ReportDetailtForTotalMoneyType>();

            switch (typeID)
            {
                // Theo ngày
                case 1:
                    listReportData = new ReportBL().SearchReportDetailtMTForAll(fromDate, toDate, reportTypeID, marketID);
                    listReportDataConvertUSD = new ReportBL().SearchReportDetailtMTForAllConvert(fromDate, toDate, reportTypeID, marketID);
                    
                    if (!string.IsNullOrEmpty(marketID))
                    {
                        if (marketID == "0")
                        {
                            foreach (ReportDetailtForTotalMoneyType item in listReportData)
                            {
                                item.ReportID = item.MarketName;
                                item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;
                            }

                            // Quy USD
                            foreach (ReportDetailtForTotalMoneyType item in listReportDataConvertUSD)
                            {
                                item.ReportID = item.MarketName;
                                item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;
                            }

                        }
                        else
                        {
                            List<string> listAsianChild = new List<string>();
                            List<string> listAsianChildConvertUSD = new List<string>();
                            List<ReportDetailtForTotalMoneyType> listDataConvert = new List<ReportDetailtForTotalMoneyType>();
                            List<ReportDetailtForTotalMoneyType> listDataConvertUSD = new List<ReportDetailtForTotalMoneyType>();

                            foreach (ReportDetailtForTotalMoneyType item in listReportData)
                            {
                                item.ReportID = item.PartnerName;
                                item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;

                                if (!listAsianChild.Contains(item.MarketName))
                                {
                                    listAsianChild.Add(item.MarketName);
                                }
                            }

                            foreach (string item in listAsianChild)
                            {
                                List<ReportDetailtForTotalMoneyType> listDataAsianChild = listReportData.Where(x => x.MarketName == item).ToList();

                                listDataConvert.Add(
                                    new ReportDetailtForTotalMoneyType()
                                    {
                                        MarketName = "Châu Á",
                                        ReportID = item,
                                        VND = listDataAsianChild.Sum(x => x.VND),
                                        USD = listDataAsianChild.Sum(x => x.USD),
                                        EUR = listDataAsianChild.Sum(x => x.EUR),
                                        CAD = listDataAsianChild.Sum(x => x.CAD),
                                        AUD = listDataAsianChild.Sum(x => x.AUD),
                                        GBP = listDataAsianChild.Sum(x => x.GBP),
                                        TongDS = listDataAsianChild.Sum(x => x.TongDS)
                                    }
                                );
                            }

                            if (listDataConvert.Count > 0)
                            {
                                listReportData = new List<ReportDetailtForTotalMoneyType>(listDataConvert);
                            }

                            // Quy USD
                            foreach (ReportDetailtForTotalMoneyType item in listReportDataConvertUSD)
                            {
                                item.ReportID = item.PartnerName;
                                item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;

                                if (!listAsianChildConvertUSD.Contains(item.MarketName))
                                {
                                    listAsianChildConvertUSD.Add(item.MarketName);
                                }
                            }

                            foreach (string item in listAsianChildConvertUSD)
                            {
                                List<ReportDetailtForTotalMoneyType> listDataAsianChild = listReportDataConvertUSD.Where(x => x.MarketName == item).ToList();

                                listDataConvertUSD.Add(
                                    new ReportDetailtForTotalMoneyType()
                                    {
                                        MarketName = "Châu Á",
                                        ReportID = item,
                                        VND = listDataAsianChild.Sum(x => x.VND),
                                        USD = listDataAsianChild.Sum(x => x.USD),
                                        EUR = listDataAsianChild.Sum(x => x.EUR),
                                        CAD = listDataAsianChild.Sum(x => x.CAD),
                                        AUD = listDataAsianChild.Sum(x => x.AUD),
                                        GBP = listDataAsianChild.Sum(x => x.GBP),
                                        TongDS = listDataAsianChild.Sum(x => x.TongDS)
                                    }
                                );
                            }

                            if (listDataConvertUSD.Count > 0)
                            {
                                listReportDataConvertUSD = new List<ReportDetailtForTotalMoneyType>(listDataConvertUSD);
                            }
                        }
                    }

                    // Set from day and to day
                    sheetReport.Cells["E4"].PutValue(fromDate.ToString("dd/MM/yyyy"));
                    sheetReport.Cells["H4"].PutValue(toDate.ToString("dd/MM/yyyy"));
                    break;
                // Theo tháng
                case 2:

                    listReportData = new ReportBL().SearchReportDetailtMTForAllForMonth(fromDate, toDate, reportTypeID, marketID);
                    listReportDataConvertUSD = new ReportBL().SearchReportDetailtMTForAllForMonthConvert(fromDate, toDate, reportTypeID, marketID);

                    if (!string.IsNullOrEmpty(marketID))
                    {
                        if (marketID == "0")
                        {
                            foreach (ReportDetailtForTotalMoneyType item in listReportData)
                            {
                                item.ReportID = item.MarketName;
                                item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;
                            }

                            // Quy USD
                            foreach (ReportDetailtForTotalMoneyType item in listReportDataConvertUSD)
                            {
                                item.ReportID = item.MarketName;
                                item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;
                            }

                        }
                        else
                        {
                            List<string> listAsianChild = new List<string>();
                            List<string> listAsianChildConvertUSD = new List<string>();
                            List<ReportDetailtForTotalMoneyType> listDataConvert = new List<ReportDetailtForTotalMoneyType>();
                            List<ReportDetailtForTotalMoneyType> listDataConvertUSD = new List<ReportDetailtForTotalMoneyType>();

                            foreach (ReportDetailtForTotalMoneyType item in listReportData)
                            {
                                item.ReportID = item.PartnerName;
                                item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;

                                if (!listAsianChild.Contains(item.MarketName))
                                {
                                    listAsianChild.Add(item.MarketName);
                                }
                            }

                            foreach (string item in listAsianChild)
                            {
                                List<ReportDetailtForTotalMoneyType> listDataAsianChild = listReportData.Where(x => x.MarketName == item).ToList();

                                listDataConvert.Add(
                                    new ReportDetailtForTotalMoneyType()
                                    {
                                        MarketName = "Châu Á",
                                        ReportID = item,
                                        VND = listDataAsianChild.Sum(x => x.VND),
                                        USD = listDataAsianChild.Sum(x => x.USD),
                                        EUR = listDataAsianChild.Sum(x => x.EUR),
                                        CAD = listDataAsianChild.Sum(x => x.CAD),
                                        AUD = listDataAsianChild.Sum(x => x.AUD),
                                        GBP = listDataAsianChild.Sum(x => x.GBP),
                                        TongDS = listDataAsianChild.Sum(x => x.TongDS)
                                    }
                                );
                            }

                            if (listDataConvert.Count > 0)
                            {
                                listReportData = new List<ReportDetailtForTotalMoneyType>(listDataConvert);
                            }

                            // Quy USD
                            foreach (ReportDetailtForTotalMoneyType item in listReportDataConvertUSD)
                            {
                                item.ReportID = item.PartnerName;
                                item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;

                                if (!listAsianChildConvertUSD.Contains(item.MarketName))
                                {
                                    listAsianChildConvertUSD.Add(item.MarketName);
                                }
                            }

                            foreach (string item in listAsianChildConvertUSD)
                            {
                                List<ReportDetailtForTotalMoneyType> listDataAsianChild = listReportDataConvertUSD.Where(x => x.MarketName == item).ToList();

                                listDataConvertUSD.Add(
                                    new ReportDetailtForTotalMoneyType()
                                    {
                                        MarketName = "Châu Á",
                                        ReportID = item,
                                        VND = listDataAsianChild.Sum(x => x.VND),
                                        USD = listDataAsianChild.Sum(x => x.USD),
                                        EUR = listDataAsianChild.Sum(x => x.EUR),
                                        CAD = listDataAsianChild.Sum(x => x.CAD),
                                        AUD = listDataAsianChild.Sum(x => x.AUD),
                                        GBP = listDataAsianChild.Sum(x => x.GBP),
                                        TongDS = listDataAsianChild.Sum(x => x.TongDS)
                                    }
                                );
                            }

                            if (listDataConvertUSD.Count > 0)
                            {
                                listReportDataConvertUSD = new List<ReportDetailtForTotalMoneyType>(listDataConvertUSD);
                            }
                        }
                    }
                    // Set from day and to day
                    sheetReport.Cells["E4"].PutValue(string.Format("{0}/{1}", fromDate.Month, fromDate.Year));
                    sheetReport.Cells["H4"].PutValue(string.Format("{0}/{1}", toDate.Month, toDate.Year));
                    break;
                // Theo năm
                default:

                    listReportData = new ReportBL().SearchReportDetailtMTForAllForYear(fromDate, toDate, reportTypeID, marketID);
                    listReportDataConvertUSD = new ReportBL().SearchReportDetailtMTForAllForYearConvert(fromDate, toDate, reportTypeID, marketID);

                    if (!string.IsNullOrEmpty(marketID))
                    {
                        if (marketID == "0")
                        {
                            foreach (ReportDetailtForTotalMoneyType item in listReportData)
                            {
                                item.ReportID = item.MarketName;
                                item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;
                            }

                            // Quy USD
                            foreach (ReportDetailtForTotalMoneyType item in listReportDataConvertUSD)
                            {
                                item.ReportID = item.MarketName;
                                item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;
                            }

                        }
                        else
                        {
                            List<string> listAsianChild = new List<string>();
                            List<string> listAsianChildConvertUSD = new List<string>();
                            List<ReportDetailtForTotalMoneyType> listDataConvert = new List<ReportDetailtForTotalMoneyType>();
                            List<ReportDetailtForTotalMoneyType> listDataConvertUSD = new List<ReportDetailtForTotalMoneyType>();

                            foreach (ReportDetailtForTotalMoneyType item in listReportData)
                            {
                                item.ReportID = item.PartnerName;
                                item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;

                                if (!listAsianChild.Contains(item.MarketName))
                                {
                                    listAsianChild.Add(item.MarketName);
                                }
                            }

                            foreach (string item in listAsianChild)
                            {
                                List<ReportDetailtForTotalMoneyType> listDataAsianChild = listReportData.Where(x => x.MarketName == item).ToList();

                                listDataConvert.Add(
                                    new ReportDetailtForTotalMoneyType()
                                    {
                                        MarketName = "Châu Á",
                                        ReportID = item,
                                        VND = listDataAsianChild.Sum(x => x.VND),
                                        USD = listDataAsianChild.Sum(x => x.USD),
                                        EUR = listDataAsianChild.Sum(x => x.EUR),
                                        CAD = listDataAsianChild.Sum(x => x.CAD),
                                        AUD = listDataAsianChild.Sum(x => x.AUD),
                                        GBP = listDataAsianChild.Sum(x => x.GBP),
                                        TongDS = listDataAsianChild.Sum(x => x.TongDS)
                                    }
                                );
                            }

                            if (listDataConvert.Count > 0)
                            {
                                listReportData = new List<ReportDetailtForTotalMoneyType>(listDataConvert);
                            }

                            // Quy USD
                            foreach (ReportDetailtForTotalMoneyType item in listReportDataConvertUSD)
                            {
                                item.ReportID = item.PartnerName;
                                item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;

                                if (!listAsianChildConvertUSD.Contains(item.MarketName))
                                {
                                    listAsianChildConvertUSD.Add(item.MarketName);
                                }
                            }

                            foreach (string item in listAsianChildConvertUSD)
                            {
                                List<ReportDetailtForTotalMoneyType> listDataAsianChild = listReportDataConvertUSD.Where(x => x.MarketName == item).ToList();

                                listDataConvertUSD.Add(
                                    new ReportDetailtForTotalMoneyType()
                                    {
                                        MarketName = "Châu Á",
                                        ReportID = item,
                                        VND = listDataAsianChild.Sum(x => x.VND),
                                        USD = listDataAsianChild.Sum(x => x.USD),
                                        EUR = listDataAsianChild.Sum(x => x.EUR),
                                        CAD = listDataAsianChild.Sum(x => x.CAD),
                                        AUD = listDataAsianChild.Sum(x => x.AUD),
                                        GBP = listDataAsianChild.Sum(x => x.GBP),
                                        TongDS = listDataAsianChild.Sum(x => x.TongDS)
                                    }
                                );
                            }

                            if (listDataConvertUSD.Count > 0)
                            {
                                listReportDataConvertUSD = new List<ReportDetailtForTotalMoneyType>(listDataConvertUSD);
                            }
                        }
                    }
                    // Set from day and to day
                    sheetReport.Cells["E4"].PutValue(fromDate.Year.ToString());
                    sheetReport.Cells["H4"].PutValue(toDate.Year.ToString());
                    break;
            }

            DataTable dataTable = new DataTable();

            if (listReportData.Count > 0)
            {
                // Add dòng tổng vào list danh sách
                ReportDetailtForTotalMoneyType reportForMaket = new ReportDetailtForTotalMoneyType()
                {
                    ReportID = "Tổng",
                    VND = listReportData.Sum(x => x.VND),
                    USD = listReportData.Sum(x => x.USD),
                    EUR = listReportData.Sum(x => x.EUR),
                    CAD = listReportData.Sum(x => x.CAD),
                    AUD = listReportData.Sum(x => x.AUD),
                    GBP = listReportData.Sum(x => x.GBP),
                    TongDS = listReportData.Sum(x => x.TongDS),
                };
                listReportData.Add(reportForMaket);

                // Tạo các col cho table
                dataTable = CreateDataTableFormart();

                // Danh sách dataSet của báo cáo ngày
                DataSet dataReport = ConvertListObjectToDataSet(listReportData);

                // Đổ data vào datatble mới
                FillData(dataReport.Tables[0], dataTable);
            }
            

            // Tạo cột hearder cho Quy đô
            string title = "Đơn vị";
            CreateTitle("G6", "G6", sheetReport, title, 12);


            // Tạo cột hearder cho Quy đô
            title = "Nguyên tệ";
            CreateTitle("H6", "H6", sheetReport, title, 12);

            // Set border
            Style style = new CellsFactory().CreateStyle();
            style.SetBorder(BorderType.BottomBorder, CellBorderType.Thin, Color.Black);
            style.SetBorder(BorderType.LeftBorder, CellBorderType.Thin, Color.Black);
            style.SetBorder(BorderType.RightBorder, CellBorderType.Thin, Color.Black);
            style.SetBorder(BorderType.TopBorder, CellBorderType.Thin, Color.Black);

            int totalRowTable1 = 0;

            if (dataTable.Rows.Count > 0)
            {
                int stepRow = 0;
                // total row = row start + số row hiện có
                int totalRow = dataTable.Rows.Count + 7;
                totalRowTable1 = totalRow;
                // Số dòng của row
                for (int a = 7; a < totalRow; a++)
                {
                    int stepColumn = 0;
                    // Số cột trong báo cáo cần hiển thị
                    // Tổng số cột hiển thị = Số cột hiển thị bắt đầu + tổng số cột cần hiển thị
                    // Số 6 là cột marketName
                    int totalCol = 1 + 7;
                    for (int b = 1; b < totalCol; b++)
                    {
                        // Giá trị của value trong table
                        string valueOfTable = dataTable.Rows[stepRow][stepColumn].ToString();

                        // Insert vào dòng cột xác định trong Excel
                        sheetReport.Cells[a, b].PutValue(valueOfTable, true);

                        // set style cho number
                        if (b == 1)
                        {
                            style.Custom = "";
                        }
                        else
                        {
                            style.Custom = "#,##0.00";
                        }

                        // set border
                        sheetReport.Cells[a, b].SetStyle(style);

                        // Cột tổng cộng
                        if (b.Equals(totalCol - 1))
                        {
                            sheetReport.Cells[a, b].PutValue(valueOfTable, true, true);
                            style.Font.IsBold = true;
                            sheetReport.Cells[a, b].SetStyle(style);
                        }

                        // Trường hợp thuộc dòng cuối Tổng
                        if (a.Equals(totalRow - 1))
                        {
                            sheetReport.Cells[a, b].PutValue(valueOfTable, true, true);
                            style.Font.IsBold = true;
                            sheetReport.Cells[a, b].SetStyle(style);
                        }

                        // Set lại giá trị mặt định
                        style.Font.IsBold = false;
                        // Tăng cột theo dòng của table
                        stepColumn++;
                    }
                    // Tăng dòng của table lên
                    stepRow++;
                }
            }
            else
            {
                sheetReport.Cells["D10"].PutValue("Không có dữ liệu");
            }

            // Quy USD
            dataTable = new DataTable();

            if (listReportDataConvertUSD.Count > 0)
            {
                // Add dòng tổng vào list danh sách
                ReportDetailtForTotalMoneyType reportForMaket = new ReportDetailtForTotalMoneyType()
                {
                    ReportID = "Tổng",
                    VND = listReportDataConvertUSD.Sum(x => x.VND),
                    USD = listReportDataConvertUSD.Sum(x => x.USD),
                    EUR = listReportDataConvertUSD.Sum(x => x.EUR),
                    CAD = listReportDataConvertUSD.Sum(x => x.CAD),
                    AUD = listReportDataConvertUSD.Sum(x => x.AUD),
                    GBP = listReportDataConvertUSD.Sum(x => x.GBP),
                    TongDS = listReportDataConvertUSD.Sum(x => x.TongDS),
                };
                listReportDataConvertUSD.Add(reportForMaket);

                // Tạo các col cho table
                dataTable = CreateDataTableFormart(true);

                // Danh sách dataSet của báo cáo ngày
                DataSet dataReport = ConvertListObjectToDataSet(listReportDataConvertUSD);

                // Đổ data vào datatble mới
                FillData(dataReport.Tables[0], dataTable, true);
            }

            // Tạo cột hearder cho Quy đô
            title = "Thị trường";
            CreateTitle(string.Format("B{0}", totalRowTable1 + 4), string.Format("B{0}", totalRowTable1 + 4), sheetReport, title, 12, true);

            title = "VND";
            CreateTitle(string.Format("C{0}", totalRowTable1 + 4), string.Format("C{0}", totalRowTable1 + 4), sheetReport, title, 12, true);

            title = "USD";
            CreateTitle(string.Format("D{0}", totalRowTable1 + 4), string.Format("D{0}", totalRowTable1 + 4), sheetReport, title, 12, true);

            title = "EUR";
            CreateTitle(string.Format("E{0}", totalRowTable1 + 4), string.Format("E{0}", totalRowTable1 + 4), sheetReport, title, 12, true);

            title = "CAD";
            CreateTitle(string.Format("F{0}", totalRowTable1 + 4), string.Format("F{0}", totalRowTable1 + 4), sheetReport, title, 12, true);

            title = "AUD";
            CreateTitle(string.Format("G{0}", totalRowTable1 + 4), string.Format("G{0}", totalRowTable1 + 4), sheetReport, title, 12, true);

            title = "GBP";
            CreateTitle(string.Format("H{0}", totalRowTable1 + 4), string.Format("H{0}", totalRowTable1 + 4), sheetReport, title, 12, true);

            title = "Tổng";
            CreateTitle(string.Format("I{0}", totalRowTable1 + 4), string.Format("I{0}", totalRowTable1 + 4), sheetReport, title, 12, true);

            // Tạo cột hearder cho Quy đô
            title = "Đơn vị";
            CreateTitle(string.Format("H{0}", totalRowTable1 + 3), string.Format("H{0}", totalRowTable1 + 3), sheetReport, title, 12);


            // Tạo cột hearder cho Quy đô
            title = "Quy USD";
            CreateTitle(string.Format("I{0}", totalRowTable1 + 3), string.Format("I{0}", totalRowTable1 + 3), sheetReport, title, 12);

            if (dataTable.Rows.Count > 0)
            {
                int stepRow = 0;
                // total row = row start + số row hiện có
                int totalRow = dataTable.Rows.Count + 4 + totalRowTable1;
                // Số dòng của row
                for (int a = totalRowTable1 + 4; a < totalRow; a++)
                {
                    int stepColumn = 0;
                    // Số cột trong báo cáo cần hiển thị
                    // Tổng số cột hiển thị = Số cột hiển thị bắt đầu + tổng số cột cần hiển thị
                    // Số 6 là cột marketName
                    int totalCol = 1 + 8;
                    for (int b = 1; b < totalCol; b++)
                    {
                        // Giá trị của value trong table
                        string valueOfTable = dataTable.Rows[stepRow][stepColumn].ToString();

                        // Insert vào dòng cột xác định trong Excel
                        sheetReport.Cells[a, b].PutValue(valueOfTable, true);

                        // set style cho number
                        if (b == 1)
                        {
                            style.Custom = "";
                        }
                        else
                        {
                            style.Custom = "#,##0.00";
                        }

                        // set border
                        sheetReport.Cells[a, b].SetStyle(style);

                        // Cột tổng cộng
                        if (b.Equals(totalCol - 1))
                        {
                            sheetReport.Cells[a, b].PutValue(valueOfTable, true, true);
                            style.Font.IsBold = true;
                            sheetReport.Cells[a, b].SetStyle(style);
                        }

                        // Trường hợp thuộc dòng cuối Tổng
                        if (a.Equals(totalRow - 1))
                        {
                            sheetReport.Cells[a, b].PutValue(valueOfTable, true, true);
                            style.Font.IsBold = true;
                            sheetReport.Cells[a, b].SetStyle(style);
                        }

                        // Set lại giá trị mặt định
                        style.Font.IsBold = false;
                        // Tăng cột theo dòng của table
                        stepColumn++;
                    }
                    // Tăng dòng của table lên
                    stepRow++;
                }
            }
            else
            {
                sheetReport.Cells["D10"].PutValue("Không có dữ liệu");
            }

            //sheetReport.Cells.ConvertStringToNumericValue();
            // Chạy process
            designer.Process();

            switch (typeID)
            {
                // Theo ngày
                case 1:
                    return ExportReport("ReportDay", designer);
                    break;
                // Theo tháng
                case 2:
                    return ExportReport("ReportMonth", designer);
                    break;
                // Theo năm
                default:
                    return ExportReport("ReportYear", designer);
                    break;
            }
        }


        /// <summary>
        /// Xuất Excel theo Ngày/Tháng/Năm
        /// </summary>
        /// <param name="fromDate"></param>
        /// <param name="toDate"></param>
        /// <returns></returns>
        public ActionResult CreateExcelForDayMonthYearForOne(DateTime fromDate, DateTime toDate, int typeID, string reportTypeID, string marketID)
        {
            WorkbookDesigner designer = new WorkbookDesigner();
            string templatePath = "~/Content/Report/ReportDetailtByPartner.xlsx";
            // Get đường dẫn
            templatePath = System.Web.HttpContext.Current.Server.MapPath(templatePath);

            designer.Workbook = new DongA.Core.DongAExcel().OpenExcelFile(templatePath);
            designer.Workbook.CalculateFormula();

            // Set the glorbalization setting to change subtotal and grand total names
            GlobalizationSettings gsi = new GlobalizationSettingsImp();
            designer.Workbook.Settings.GlobalizationSettings = gsi;

            WorksheetCollection workSheets = designer.Workbook.Worksheets;
            Worksheet sheetReport = designer.Workbook.Worksheets[0];
            
            // Tạo title
            string typeReport = string.Empty;
            switch (typeID)
            {
                // Theo ngày
                case 1:
                    typeReport = "Chi tiết - Theo Ngày";
                    break;
                // Theo tháng
                case 2:
                    typeReport = "Chi tiết - Theo Tháng";
                    break;
                // Theo năm
                default:
                    typeReport = "Chi tiết - Theo Năm";
                    break;
            }

            CreateTitle("A2", "K2", sheetReport, typeReport, 14);

            // Get data report ngày
            List<ReportDetailtForTotalMoneyType> listReportData = new List<ReportDetailtForTotalMoneyType>();
            List<ReportDetailtForTotalMoneyType> listReportDataUSD = new List<ReportDetailtForTotalMoneyType>();
            // Get danh sách thị trường
            List<string> listMarket = new List<string>();

            switch (typeID)
            {
                // Theo ngày
                case 1:
                    listReportData = new ReportBL().SearchReportDetailtMTForOne(fromDate, toDate, reportTypeID, marketID);
                    // Quy USD
                    listReportDataUSD = new ReportBL().SearchReportDetailtMTForOneConvert(fromDate, toDate, reportTypeID, marketID);

                    if (marketID.Contains("005"))
                    {
                        listMarket = new List<string>();

                        foreach (ReportDetailtForTotalMoneyType item in listReportData)
                        {
                            if (!listMarket.Contains(item.MarketName))
                            {
                                listMarket.Add(item.MarketName);
                            }
                            item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;
                        }

                        List<ReportDetailtForTotalMoneyType> listDataConvert = new List<ReportDetailtForTotalMoneyType>();
                        foreach (string item in listMarket)
                        {
                            List<ReportDetailtForTotalMoneyType> listDataItem = listReportData.Where(x => x.MarketName == item).ToList();

                            listDataConvert.Add(
                                new ReportDetailtForTotalMoneyType()
                                {
                                    MarketName = "Thị trường Châu Á",
                                    PartnerName = item,
                                    ReportID = item,
                                    VND = listDataItem.Sum(x => x.VND),
                                    USD = listDataItem.Sum(x => x.USD),
                                    EUR = listDataItem.Sum(x => x.EUR),
                                    CAD = listDataItem.Sum(x => x.CAD),
                                    AUD = listDataItem.Sum(x => x.AUD),
                                    GBP = listDataItem.Sum(x => x.GBP),
                                }
                            );
                        }

                        if (listDataConvert.Count > 0)
                        {
                            listReportData = new List<ReportDetailtForTotalMoneyType>(listDataConvert);
                        }
                    }

                    foreach (ReportDetailtForTotalMoneyType item in listReportData)
                    {
                        item.ReportID = item.PartnerName;
                        item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;
                    }

                    // Quy USD
                    if (marketID.Contains("005"))
                    {
                        listMarket = new List<string>();

                        foreach (ReportDetailtForTotalMoneyType item in listReportDataUSD)
                        {
                            if (!listMarket.Contains(item.MarketName))
                            {
                                listMarket.Add(item.MarketName);
                            }
                            item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;
                        }

                        List<ReportDetailtForTotalMoneyType> listDataConvert = new List<ReportDetailtForTotalMoneyType>();
                        foreach (string item in listMarket)
                        {
                            List<ReportDetailtForTotalMoneyType> listDataItem = listReportDataUSD.Where(x => x.MarketName == item).ToList();

                            listDataConvert.Add(
                                new ReportDetailtForTotalMoneyType()
                                {
                                    MarketName = "Thị trường Châu Á",
                                    PartnerName = item,
                                    ReportID = item,
                                    VND = listDataItem.Sum(x => x.VND),
                                    USD = listDataItem.Sum(x => x.USD),
                                    EUR = listDataItem.Sum(x => x.EUR),
                                    CAD = listDataItem.Sum(x => x.CAD),
                                    AUD = listDataItem.Sum(x => x.AUD),
                                    GBP = listDataItem.Sum(x => x.GBP),
                                }
                            );
                        }

                        if (listDataConvert.Count > 0)
                        {
                            listReportDataUSD = new List<ReportDetailtForTotalMoneyType>(listDataConvert);
                        }
                    }

                    foreach (ReportDetailtForTotalMoneyType item in listReportDataUSD)
                    {
                        item.ReportID = item.PartnerName;
                        item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;
                    }

                    // Set from day and to day
                    sheetReport.Cells["E4"].PutValue(fromDate.ToString("dd/MM/yyyy"));
                    sheetReport.Cells["H4"].PutValue(toDate.ToString("dd/MM/yyyy"));
                    break;
                // Theo tháng
                case 2:

                    listReportData = new ReportBL().SearchReportDetailtMTForOneForMonth(fromDate, toDate, reportTypeID, marketID);
                    listReportDataUSD = new ReportBL().SearchReportDetailtMTForOneForMonthConvert(fromDate, toDate, reportTypeID, marketID);

                    // Nguyên tệ
                    if (marketID.Contains("005"))
                    {
                        listMarket = new List<string>();

                        foreach (ReportDetailtForTotalMoneyType item in listReportData)
                        {
                            if (!listMarket.Contains(item.MarketName))
                            {
                                listMarket.Add(item.MarketName);
                            }
                            item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;
                        }

                        List<ReportDetailtForTotalMoneyType> listDataConvert = new List<ReportDetailtForTotalMoneyType>();
                        foreach (string item in listMarket)
                        {
                            List<ReportDetailtForTotalMoneyType> listDataItem = listReportData.Where(x => x.MarketName == item).ToList();

                            listDataConvert.Add(
                                new ReportDetailtForTotalMoneyType()
                                {
                                    MarketName = "Thị trường Châu Á",
                                    PartnerName = item,
                                    ReportID = item,
                                    VND = listDataItem.Sum(x => x.VND),
                                    USD = listDataItem.Sum(x => x.USD),
                                    EUR = listDataItem.Sum(x => x.EUR),
                                    CAD = listDataItem.Sum(x => x.CAD),
                                    AUD = listDataItem.Sum(x => x.AUD),
                                    GBP = listDataItem.Sum(x => x.GBP),
                                }
                            );
                        }

                        if (listDataConvert.Count > 0)
                        {
                            listReportData = new List<ReportDetailtForTotalMoneyType>(listDataConvert);
                        }
                    }

                    foreach (ReportDetailtForTotalMoneyType item in listReportData)
                    {
                        item.ReportID = item.PartnerName;
                        item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;
                    }

                    // Quy USD
                    if (marketID.Contains("005"))
                    {
                        listMarket = new List<string>();

                        foreach (ReportDetailtForTotalMoneyType item in listReportDataUSD)
                        {
                            if (!listMarket.Contains(item.MarketName))
                            {
                                listMarket.Add(item.MarketName);
                            }
                            item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;
                        }

                        List<ReportDetailtForTotalMoneyType> listDataConvert = new List<ReportDetailtForTotalMoneyType>();
                        foreach (string item in listMarket)
                        {
                            List<ReportDetailtForTotalMoneyType> listDataItem = listReportDataUSD.Where(x => x.MarketName == item).ToList();

                            listDataConvert.Add(
                                new ReportDetailtForTotalMoneyType()
                                {
                                    MarketName = "Thị trường Châu Á",
                                    PartnerName = item,
                                    ReportID = item,
                                    VND = listDataItem.Sum(x => x.VND),
                                    USD = listDataItem.Sum(x => x.USD),
                                    EUR = listDataItem.Sum(x => x.EUR),
                                    CAD = listDataItem.Sum(x => x.CAD),
                                    AUD = listDataItem.Sum(x => x.AUD),
                                    GBP = listDataItem.Sum(x => x.GBP),
                                }
                            );
                        }

                        if (listDataConvert.Count > 0)
                        {
                            listReportDataUSD = new List<ReportDetailtForTotalMoneyType>(listDataConvert);
                        }
                    }

                    foreach (ReportDetailtForTotalMoneyType item in listReportDataUSD)
                    {
                        item.ReportID = item.PartnerName;
                        item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;
                    }

                    // Set from day and to day
                    sheetReport.Cells["E4"].PutValue(string.Format("{0}/{1}", fromDate.Month, fromDate.Year));
                    sheetReport.Cells["H4"].PutValue(string.Format("{0}/{1}", toDate.Month, toDate.Year));
                    break;
                // Theo năm
                default:

                    listReportData = new ReportBL().SearchReportDetailtMTForOneForYear(fromDate, toDate, reportTypeID, marketID);
                    listReportDataUSD = new ReportBL().SearchReportDetailtMTForOneForYearConvert(fromDate, toDate, reportTypeID, marketID);

                    // Nguyên tệ
                    if (marketID.Contains("005"))
                    {
                        listMarket = new List<string>();

                        foreach (ReportDetailtForTotalMoneyType item in listReportData)
                        {
                            if (!listMarket.Contains(item.MarketName))
                            {
                                listMarket.Add(item.MarketName);
                            }
                            item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;
                        }

                        List<ReportDetailtForTotalMoneyType> listDataConvert = new List<ReportDetailtForTotalMoneyType>();
                        foreach (string item in listMarket)
                        {
                            List<ReportDetailtForTotalMoneyType> listDataItem = listReportData.Where(x => x.MarketName == item).ToList();

                            listDataConvert.Add(
                                new ReportDetailtForTotalMoneyType()
                                {
                                    MarketName = "Thị trường Châu Á",
                                    PartnerName = item,
                                    ReportID = item,
                                    VND = listDataItem.Sum(x => x.VND),
                                    USD = listDataItem.Sum(x => x.USD),
                                    EUR = listDataItem.Sum(x => x.EUR),
                                    CAD = listDataItem.Sum(x => x.CAD),
                                    AUD = listDataItem.Sum(x => x.AUD),
                                    GBP = listDataItem.Sum(x => x.GBP),
                                }
                            );
                        }

                        if (listDataConvert.Count > 0)
                        {
                            listReportData = new List<ReportDetailtForTotalMoneyType>(listDataConvert);
                        }
                    }

                    foreach (ReportDetailtForTotalMoneyType item in listReportData)
                    {
                        item.ReportID = item.PartnerName;
                        item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;
                    }

                    // Quy USD
                    if (marketID.Contains("005"))
                    {
                        listMarket = new List<string>();

                        foreach (ReportDetailtForTotalMoneyType item in listReportDataUSD)
                        {
                            if (!listMarket.Contains(item.MarketName))
                            {
                                listMarket.Add(item.MarketName);
                            }
                            item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;
                        }

                        List<ReportDetailtForTotalMoneyType> listDataConvert = new List<ReportDetailtForTotalMoneyType>();
                        foreach (string item in listMarket)
                        {
                            List<ReportDetailtForTotalMoneyType> listDataItem = listReportDataUSD.Where(x => x.MarketName == item).ToList();

                            listDataConvert.Add(
                                new ReportDetailtForTotalMoneyType()
                                {
                                    MarketName = "Thị trường Châu Á",
                                    PartnerName = item,
                                    ReportID = item,
                                    VND = listDataItem.Sum(x => x.VND),
                                    USD = listDataItem.Sum(x => x.USD),
                                    EUR = listDataItem.Sum(x => x.EUR),
                                    CAD = listDataItem.Sum(x => x.CAD),
                                    AUD = listDataItem.Sum(x => x.AUD),
                                    GBP = listDataItem.Sum(x => x.GBP),
                                }
                            );
                        }

                        if (listDataConvert.Count > 0)
                        {
                            listReportDataUSD = new List<ReportDetailtForTotalMoneyType>(listDataConvert);
                        }
                    }

                    foreach (ReportDetailtForTotalMoneyType item in listReportDataUSD)
                    {
                        item.ReportID = item.PartnerName;
                        item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;
                    }

                    // Set from day and to day
                    sheetReport.Cells["E4"].PutValue(fromDate.Year.ToString());
                    sheetReport.Cells["H4"].PutValue(toDate.Year.ToString());
                    break;
            }

            DataTable dataTable = new DataTable();

            if (listReportData.Count > 0)
            {
                // Add dòng tổng vào list danh sách
                ReportDetailtForTotalMoneyType reportForMaket = new ReportDetailtForTotalMoneyType()
                {
                    STT = "",
                    ReportID = "Tổng tất cả",
                    VND = listReportData.Sum(item => item.VND),
                    USD = listReportData.Sum(item => item.USD),
                    EUR = listReportData.Sum(item => item.EUR),
                    CAD = listReportData.Sum(item => item.CAD),
                    AUD = listReportData.Sum(item => item.AUD),
                    GBP = listReportData.Sum(item => item.GBP),
                    TongDS = listReportData.Sum(item => item.TongDS),
                };
                listReportData.Add(reportForMaket);

                // Tạo các col cho table
                dataTable = CreateDataTableFormart();

                // Danh sách dataSet của báo cáo ngày
                DataSet dataReport = ConvertListObjectToDataSet(listReportData);

                // Đổ data vào datatble mới
                FillData(dataReport.Tables[0], dataTable);
            }

            // Set border
            Style style = new CellsFactory().CreateStyle();
            style.SetBorder(BorderType.TopBorder, CellBorderType.Thin, Color.Black);
            style.SetBorder(BorderType.BottomBorder, CellBorderType.Thin, Color.Black);
            style.SetBorder(BorderType.LeftBorder, CellBorderType.Thin, Color.Black);
            style.SetBorder(BorderType.RightBorder, CellBorderType.Thin, Color.Black);

            // Tạo cột hearder cho Quy đô
            string title = "Đối tác";
            CreateTitle("B7", "B7", sheetReport, title, 12, true);

            // Tạo cột hearder cho Quy đô
            title = "Đơn vị";
            CreateTitle("G6", "G6", sheetReport, title, 12);


            // Tạo cột hearder cho Quy đô
            title = "Nguyên tệ";
            CreateTitle("H6", "H6", sheetReport, title, 12);

            int totalRowTable1 = 0;

            if (dataTable.Rows.Count > 0)
            {
                // List Duplicate
                bool valueCheck = true;
                // listMarket
                int count = 0;
                int stepRow = 0;
                // total row = row start + số row hiện có
                // Thêm 1 vào để thêm 1 dòng thị trường
                int totalRow = dataTable.Rows.Count + 7;
                // Get total row from table 1
                totalRowTable1 = totalRow;
                // check trùng
                List<string> listDuplicate = new List<string>();
                // Số dòng của row
                for (int a = 7; a < totalRow; a++)
                {
                    int stepColumn = 0;
                    // Số cột trong báo cáo cần hiển thị
                    // Tổng số cột hiển thị = Số cột hiển thị bắt đầu + tổng số cột cần hiển thị
                    // Số 6 là cột marketName
                    int totalCol = 1 + 7;
                    for (int b = 1; b < totalCol; b++)
                    {
                        // Giá trị của value trong table
                        string valueOfTable = dataTable.Rows[stepRow][stepColumn].ToString();

                        // Insert vào dòng cột xác định trong Excel
                        sheetReport.Cells[a, b].PutValue(valueOfTable, true);

                        // set style cho number
                        if (b == 1)
                        {
                            style.Custom = "";
                        }
                        else
                        {
                            style.Custom = "#,##0.00";
                        }

                        // set border
                        sheetReport.Cells[a, b].SetStyle(style);

                        // Cột tổng cộng
                        if (b.Equals(totalCol - 1))
                        {
                            sheetReport.Cells[a, b].PutValue(valueOfTable, true, true);
                            style.Font.IsBold = true;
                            sheetReport.Cells[a, b].SetStyle(style);
                        }

                        // Trường hợp thuộc dòng cuối Tổng
                        if (a.Equals(totalRow - 1))
                        {
                            sheetReport.Cells[a, b].PutValue(valueOfTable, true, true);
                            style.Font.IsBold = true;
                            sheetReport.Cells[a, b].SetStyle(style);
                        }

                        // Set lại giá trị mặt định
                        style.Font.IsBold = false;
                        // Tăng cột theo dòng của table
                        stepColumn++;
                    }
                    // Tăng dòng của table lên
                    stepRow++;
                }
                
            }
            else
            {
                sheetReport.Cells["D10"].PutValue("Không có dữ liệu");
            }

            // Quy USD
            dataTable = new DataTable();

            if (listReportDataUSD.Count > 0)
            {
                // Add dòng tổng vào list danh sách
                ReportDetailtForTotalMoneyType reportForMaket = new ReportDetailtForTotalMoneyType()
                {
                    ReportID = "Tổng",
                    VND = listReportDataUSD.Sum(x => x.VND),
                    USD = listReportDataUSD.Sum(x => x.USD),
                    EUR = listReportDataUSD.Sum(x => x.EUR),
                    CAD = listReportDataUSD.Sum(x => x.CAD),
                    AUD = listReportDataUSD.Sum(x => x.AUD),
                    GBP = listReportDataUSD.Sum(x => x.GBP),
                    TongDS = listReportDataUSD.Sum(x => x.TongDS),
                };
                listReportDataUSD.Add(reportForMaket);

                // Tạo các col cho table
                dataTable = CreateDataTableFormart(true);

                // Danh sách dataSet của báo cáo ngày
                DataSet dataReport = ConvertListObjectToDataSet(listReportDataUSD);

                // Đổ data vào datatble mới
                FillData(dataReport.Tables[0], dataTable, true);
            }

            // Tạo cột hearder cho Quy đô
            title = "Đối tác";
            CreateTitle(string.Format("B{0}", totalRowTable1 + 4), string.Format("B{0}", totalRowTable1 + 4), sheetReport, title, 12, true);

            title = "VND";
            CreateTitle(string.Format("C{0}", totalRowTable1 + 4), string.Format("C{0}", totalRowTable1 + 4), sheetReport, title, 12, true);

            title = "USD";
            CreateTitle(string.Format("D{0}", totalRowTable1 + 4), string.Format("D{0}", totalRowTable1 + 4), sheetReport, title, 12, true);

            title = "EUR";
            CreateTitle(string.Format("E{0}", totalRowTable1 + 4), string.Format("E{0}", totalRowTable1 + 4), sheetReport, title, 12, true);

            title = "CAD";
            CreateTitle(string.Format("F{0}", totalRowTable1 + 4), string.Format("F{0}", totalRowTable1 + 4), sheetReport, title, 12, true);

            title = "AUD";
            CreateTitle(string.Format("G{0}", totalRowTable1 + 4), string.Format("G{0}", totalRowTable1 + 4), sheetReport, title, 12, true);

            title = "GBP";
            CreateTitle(string.Format("H{0}", totalRowTable1 + 4), string.Format("H{0}", totalRowTable1 + 4), sheetReport, title, 12, true);

            title = "Tổng";
            CreateTitle(string.Format("I{0}", totalRowTable1 + 4), string.Format("I{0}", totalRowTable1 + 4), sheetReport, title, 12, true);

            // Tạo cột hearder cho Quy đô
            title = "Đơn vị";
            CreateTitle(string.Format("H{0}", totalRowTable1 + 3), string.Format("H{0}", totalRowTable1 + 3), sheetReport, title, 12);


            // Tạo cột hearder cho Quy đô
            title = "Quy USD";
            CreateTitle(string.Format("I{0}", totalRowTable1 + 3), string.Format("I{0}", totalRowTable1 + 3), sheetReport, title, 12);

            if (dataTable.Rows.Count > 0)
            {
                int stepRow = 0;
                // total row = row start + số row hiện có
                int totalRow = dataTable.Rows.Count + 4 + totalRowTable1;
                // Số dòng của row
                for (int a = totalRowTable1 + 4; a < totalRow; a++)
                {
                    int stepColumn = 0;
                    // Số cột trong báo cáo cần hiển thị
                    // Tổng số cột hiển thị = Số cột hiển thị bắt đầu + tổng số cột cần hiển thị
                    // Số 6 là cột marketName
                    int totalCol = 1 + 8;
                    for (int b = 1; b < totalCol; b++)
                    {
                        // Giá trị của value trong table
                        string valueOfTable = dataTable.Rows[stepRow][stepColumn].ToString();

                        // Insert vào dòng cột xác định trong Excel
                        sheetReport.Cells[a, b].PutValue(valueOfTable, true);

                        // set style cho number
                        if (b == 1)
                        {
                            style.Custom = "";
                        }
                        else
                        {
                            style.Custom = "#,##0.00";
                        }

                        // set border
                        sheetReport.Cells[a, b].SetStyle(style);

                        // Cột tổng cộng
                        if (b.Equals(totalCol - 1))
                        {
                            sheetReport.Cells[a, b].PutValue(valueOfTable, true, true);
                            style.Font.IsBold = true;
                            sheetReport.Cells[a, b].SetStyle(style);
                        }

                        // Trường hợp thuộc dòng cuối Tổng
                        if (a.Equals(totalRow - 1))
                        {
                            sheetReport.Cells[a, b].PutValue(valueOfTable, true, true);
                            style.Font.IsBold = true;
                            sheetReport.Cells[a, b].SetStyle(style);
                        }

                        // Set lại giá trị mặt định
                        style.Font.IsBold = false;
                        // Tăng cột theo dòng của table
                        stepColumn++;
                    }
                    // Tăng dòng của table lên
                    stepRow++;
                }
            }
            else
            {
                sheetReport.Cells["D10"].PutValue("Không có dữ liệu");
            }

            //sheetReport.Cells.ConvertStringToNumericValue();
            // Chạy process
            designer.Process();

            switch (typeID)
            {
                // Theo ngày
                case 1:
                    return ExportReport("ReportDay", designer);
                    break;
                // Theo tháng
                case 2:
                    return ExportReport("ReportMonth", designer);
                    break;
                // Theo năm
                default:
                    return ExportReport("ReportYear", designer);
                    break;
            }
        }

        /// <summary>
        /// Tạo mẫu cho Excel cho so sánh theo giai đoạn cho tất cả
        /// </summary>
        /// <param name="gradationID"></param>
        /// <param name="year"></param>
        /// <param name="typeID"></param>
        /// <returns></returns>
        public ActionResult CreateExcelForGradationCompare(string gradationID, int year, string reportTypeID, string marketID)
        {
            WorkbookDesigner designer = new WorkbookDesigner();
            string templatePath = "~/Content/Report/ReportDetailPartnerForGaration.xlsx";
            // Get đường dẫn
            templatePath = System.Web.HttpContext.Current.Server.MapPath(templatePath);

            designer.Workbook = new DongA.Core.DongAExcel().OpenExcelFile(templatePath);
            designer.Workbook.CalculateFormula();

            WorksheetCollection workSheets = designer.Workbook.Worksheets;
            Worksheet sheetReport = designer.Workbook.Worksheets[0];

            // Tạo title
            string typeReport = "So sánh - Theo giai đoạn - Tất cả";

            string text = string.Format(" tháng năm {0}", year);
            string textValue = "T";
            switch (gradationID)
            {
                case "1":
                    text = string.Concat("3", text);
                    textValue = string.Concat("3", textValue);
                    break;
                case "2":
                    text = string.Concat("6", text);
                    textValue = string.Concat("6", textValue);
                    break;
                case "3":
                    text = string.Concat("9", text);
                    textValue = string.Concat("9", textValue);
                    break;
                default:
                    text = string.Concat("12", text);
                    textValue = string.Concat("12", textValue);
                    break;
            }

            // Tạo title
            CreateTitle("A2", "K2", sheetReport, typeReport, 14);

            // Tạo title detailt
            string titleDetailt = string.Format("Giai đoạn: {0}", text);
            CreateTitle("A3", "K3", sheetReport, titleDetailt, 12);


            // Tạo giá trị cho cột dữ liệu của Chi quầy/ Chi nhà/ Chuyển khoản
            sheetReport.Cells["C62"].PutValue(string.Format("Năm {0} ", year - 1));
            sheetReport.Cells["D62"].PutValue(string.Format("Năm {0} ", year));

            sheetReport.Cells["E62"].PutValue(string.Format("Năm {0} ", year - 1));
            sheetReport.Cells["F62"].PutValue(string.Format("Năm {0} ", year));

            sheetReport.Cells["G62"].PutValue(string.Format("Năm {0} ", year - 1));
            sheetReport.Cells["H62"].PutValue(string.Format("Năm {0} ", year));

            sheetReport.Cells["I62"].PutValue(string.Format("Năm {0} ", year - 1));
            sheetReport.Cells["J62"].PutValue(string.Format("Năm {0} ", year));

            List<ReportDetailtForTotalMoneyType> listDataGradation = new ReportBL().ReportDetailtMTGradationCompareForAllConvert(year, int.Parse(gradationID), reportTypeID, marketID);

            // Khởi tạo datatable
            DataTable table = new DataTable();
            // Tạo các cột cho datatable
            table.Columns.Add("MarketName", typeof(String));
            table.Columns.Add("VND1", typeof(double));
            table.Columns.Add("VND2", typeof(double));

            table.Columns.Add("USD1", typeof(double));
            table.Columns.Add("USD2", typeof(double));

            table.Columns.Add("EUR1", typeof(double));
            table.Columns.Add("EUR2", typeof(double));

            table.Columns.Add("CAD1", typeof(double));
            table.Columns.Add("CAD2", typeof(double));

            table.Columns.Add("AUD1", typeof(double));
            table.Columns.Add("AUD2", typeof(double));

            table.Columns.Add("GBP1", typeof(double));
            table.Columns.Add("GBP2", typeof(double));

            table.Columns.Add("TDS1", typeof(double));
            table.Columns.Add("TDS2", typeof(double));

            List<string> listMarket = new List<string>();

            // Trường hợp chọn tất cả thị trường
            if (marketID.Equals("0"))
            {
                foreach (ReportDetailtForTotalMoneyType item in listDataGradation)
                {
                    item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;
                }

                // Danh sách mã thị trường của Tất cả
                listMarket = new List<string>() { "003", "005", "001", "002", "014", "004" };

                foreach (string item in listMarket)
                {
                    // Cùng kì
                    ReportDetailtForTotalMoneyType dataItemLastYear = listDataGradation.Find(x => x.MarketCode == item && x.Year == (year - 1).ToString());
                    ReportDetailtForTotalMoneyType dataItemYear = listDataGradation.Find(x => x.MarketCode == item && x.Year == year.ToString());

                    // add item vào table
                    table.Rows.Add(dataItemLastYear.MarketName
                        , dataItemLastYear.VND, dataItemYear.VND, dataItemLastYear.USD, dataItemYear.USD, dataItemLastYear.EUR, dataItemYear.EUR
                        , dataItemLastYear.CAD, dataItemYear.CAD, dataItemLastYear.AUD, dataItemYear.AUD, dataItemLastYear.GBP, dataItemYear.GBP
                        , dataItemLastYear.TongDS, dataItemYear.TongDS
                        );
                }

                DataRow row = table.NewRow();
                row["MarketName"] = "Tổng";
                row["VND1"] = table.Compute("Sum(VND1)", "");
                row["VND2"] = table.Compute("Sum(VND2)", "");

                row["USD1"] = table.Compute("Sum(USD1)", "");
                row["USD2"] = table.Compute("Sum(USD2)", "");

                row["EUR1"] = table.Compute("Sum(EUR1)", "");
                row["EUR2"] = table.Compute("Sum(EUR2)", "");

                row["CAD1"] = table.Compute("Sum(CAD1)", "");
                row["CAD2"] = table.Compute("Sum(CAD2)", "");

                row["AUD1"] = table.Compute("Sum(AUD1)", "");
                row["AUD2"] = table.Compute("Sum(AUD2)", "");

                row["GBP1"] = table.Compute("Sum(GBP1)", "");
                row["GBP2"] = table.Compute("Sum(GBP2)", "");

                row["TDS1"] = table.Compute("Sum(TDS1)", "");
                row["TDS2"] = table.Compute("Sum(TDS2)", "");
                table.Rows.Add(row);

            }
            else
            {
                // List thị trường
                listMarket = new List<string>();

                foreach (ReportDetailtForTotalMoneyType item in listDataGradation)
                {
                    if (!listMarket.Contains(item.MarketName))
                    {
                        listMarket.Add(item.MarketName);
                    }
                    item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;
                }

                foreach (string item in listMarket)
                {
                    List<ReportDetailtForTotalMoneyType> listDetailtYear = listDataGradation.Where(x => x.MarketName == item && x.Year == year.ToString()).ToList();
                    List<ReportDetailtForTotalMoneyType> listDetailtLastYear = listDataGradation.Where(x => x.MarketName == item && x.Year == (year - 1).ToString()).ToList();

                    if (listDetailtYear.Count == 0)
                    {
                        listDetailtYear = new List<ReportDetailtForTotalMoneyType>();
                    }

                    if (listDetailtLastYear.Count == 0)
                    {
                        listDetailtLastYear = new List<ReportDetailtForTotalMoneyType>();
                    }

                    // Year
                    double sumVNDYear = listDetailtYear.Sum(x => x.VND);
                    double sumUSDYear = listDetailtYear.Sum(x => x.USD);
                    double sumEURYear = listDetailtYear.Sum(x => x.EUR);
                    double sumCADYear = listDetailtYear.Sum(x => x.CAD);
                    double sumAUDYear = listDetailtYear.Sum(x => x.AUD);
                    double sumGBPYear = listDetailtYear.Sum(x => x.GBP);
                    double sumTongDSYear = sumVNDYear + sumUSDYear + sumEURYear + sumCADYear + sumAUDYear + sumGBPYear;

                    // LastYear
                    double sumVNDLastYear = listDetailtLastYear.Sum(x => x.VND);
                    double sumUSDLastYear = listDetailtLastYear.Sum(x => x.USD);
                    double sumEURLastYear = listDetailtLastYear.Sum(x => x.EUR);
                    double sumCADLastYear = listDetailtLastYear.Sum(x => x.CAD);
                    double sumAUDLastYear = listDetailtLastYear.Sum(x => x.AUD);
                    double sumGBPLastYear = listDetailtLastYear.Sum(x => x.GBP);
                    double sumTongDSLastYear = sumVNDLastYear + sumUSDLastYear + sumEURLastYear + sumCADLastYear + sumAUDLastYear + sumGBPLastYear;

                    table.Rows.Add(item
                        , sumVNDLastYear, sumVNDYear, sumUSDYear, sumUSDLastYear, sumEURLastYear, sumEURYear
                        , sumCADLastYear, sumCADYear, sumAUDLastYear, sumAUDYear, sumGBPLastYear, sumGBPYear
                        , sumTongDSLastYear, sumTongDSYear
                        );
                }

                DataRow row = table.NewRow();
                row["MarketName"] = "Tổng";
                row["VND1"] = table.Compute("Sum(VND1)", "");
                row["VND2"] = table.Compute("Sum(VND2)", "");

                row["USD1"] = table.Compute("Sum(USD1)", "");
                row["USD2"] = table.Compute("Sum(USD2)", "");

                row["EUR1"] = table.Compute("Sum(EUR1)", "");
                row["EUR2"] = table.Compute("Sum(EUR2)", "");

                row["CAD1"] = table.Compute("Sum(CAD1)", "");
                row["CAD2"] = table.Compute("Sum(CAD2)", "");

                row["AUD1"] = table.Compute("Sum(AUD1)", "");
                row["AUD2"] = table.Compute("Sum(AUD2)", "");

                row["GBP1"] = table.Compute("Sum(GBP1)", "");
                row["GBP2"] = table.Compute("Sum(GBP2)", "");

                row["TDS1"] = table.Compute("Sum(TDS1)", "");
                row["TDS2"] = table.Compute("Sum(TDS2)", "");
                table.Rows.Add(row);
            }

            // Tạo cột hearder cho table3
            // VND
            string title = string.Format("{0} {1}", textValue, year - 1);
            CreateTitle("B82", "B82", sheetReport, title, 12, true);

            title = string.Format("{0} {1}", textValue, year);
            CreateTitle("C82", "C82", sheetReport, title, 12, true);

            // USD
            title = string.Format("{0} {1}", textValue, year - 1);
            CreateTitle("D82", "D82", sheetReport, title, 12, true);

            title = string.Format("{0} {1}", textValue, year);
            CreateTitle("E82", "E82", sheetReport, title, 12, true);

            //EUR
            title = string.Format("{0} {1}", textValue, year - 1);
            CreateTitle("F82", "F82", sheetReport, title, 12, true);

            title = string.Format("{0} {1}", textValue, year);
            CreateTitle("G82", "G82", sheetReport, title, 12, true);

            //CAD
            title = string.Format("{0} {1}", textValue, year - 1);
            CreateTitle("H82", "H82", sheetReport, title, 12, true);

            title = string.Format("{0} {1}", textValue, year);
            CreateTitle("I82", "I82", sheetReport, title, 12, true);

            // AUD
            title = string.Format("{0} {1}", textValue, year - 1);
            CreateTitle("J82", "J82", sheetReport, title, 12, true);

            title = string.Format("{0} {1}", textValue, year);
            CreateTitle("K82", "K82", sheetReport, title, 12, true);

            //GBP
            title = string.Format("{0} {1}", textValue, year - 1);
            CreateTitle("L82", "L82", sheetReport, title, 12, true);

            title = string.Format("{0} {1}", textValue, year);
            CreateTitle("M82", "M82", sheetReport, title, 12, true);

            // Tổng
            title = string.Format("{0} {1}", textValue, year - 1);
            CreateTitle("N82", "N82", sheetReport, title, 12, true);

            title = string.Format("{0} {1}", textValue, year);
            CreateTitle("O82", "O82", sheetReport, title, 12, true);



            // Set border
            Style style = new CellsFactory().CreateStyle();
            style.SetBorder(BorderType.BottomBorder, CellBorderType.Thin, Color.Black);
            style.SetBorder(BorderType.LeftBorder, CellBorderType.Thin, Color.Black);
            style.SetBorder(BorderType.RightBorder, CellBorderType.Thin, Color.Black);

            int totalRowTable1 = 0;

            if (table.Rows.Count > 0)
            {
                int stepRow = 0;
                // total row = row start + số row hiện có
                int totalRow = table.Rows.Count + 82;
                totalRowTable1 = totalRow;
                // Số dòng của row
                for (int a = 82; a < totalRow; a++)
                {
                    int stepColumn = 0;
                    // Số cột trong báo cáo cần hiển thị
                    // Tổng số cột hiển thị = Số cột hiển thị bắt đầu + tổng số cột cần hiển thị
                    int totalCol = 0 + 15;
                    for (int b = 0; b < totalCol; b++)
                    {
                        // Giá trị của value trong table
                        string valueOfTable = table.Rows[stepRow][stepColumn].ToString();

                        // Tô màu cho các dòng có giá trị tăng giảm
                        if (b >= 18)
                        {
                            decimal tryParseValue = 0;
                            decimal.TryParse(valueOfTable, out tryParseValue);
                            style.Font.Color = Color.Green;

                            if (tryParseValue < 0)
                            {
                                style.Font.Color = Color.Red;
                            }
                        }

                        // Insert vào dòng cột xác định trong Excel
                        sheetReport.Cells[a, b].PutValue(valueOfTable, true);

                        // set style cho number
                        style.Custom = "#,##0.00";

                        // set border
                        sheetReport.Cells[a, b].SetStyle(style);

                        // Cột tổng cộng
                        if (b.Equals(totalCol - 1))
                        {
                            sheetReport.Cells[a, b].PutValue(valueOfTable, true, true);
                            style.Font.IsBold = true;
                            sheetReport.Cells[a, b].SetStyle(style);
                        }

                        // Trường hợp thuộc 2 dòng cuối
                        if (a.Equals(totalRow - 1))
                        {
                            sheetReport.Cells[a, b].PutValue(valueOfTable, true, true);
                            style.Font.IsBold = true;
                            sheetReport.Cells[a, b].SetStyle(style);
                        }

                        // Set lại giá trị mặt định
                        style.Font.IsBold = false;
                        // Tăng cột theo dòng của table
                        stepColumn++;
                    }
                    // Tăng dòng của table lên
                    stepRow++;

                    // Set lại color cho dòng hiện tại 
                    style.Font.Color = Color.Black;
                }
            }
            else
            {
                sheetReport.Cells["D10"].PutValue("Không có dữ liệu");
            }


            // Tăng/Giảm so với cùng kì năm trước
            listDataGradation = new ReportBL().ReportDetailtMTGradationCompareForAllConvert(year, int.Parse(gradationID), reportTypeID, marketID);

            foreach (ReportDetailtForTotalMoneyType item in listDataGradation)
            {
                item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;
            }

            // Danh sách mã thị trường của Tất cả
            listMarket = new List<string>() { "003", "005", "001", "002", "014", "004" };

            // Khởi tạo datatable
            table = new DataTable();
            // Tạo các cột cho datatable
            table.Columns.Add("MarketName", typeof(String));
            table.Columns.Add("VND1", typeof(double));
            table.Columns.Add("USD1", typeof(double));
            table.Columns.Add("EUR1", typeof(double));
            table.Columns.Add("CAD1", typeof(double));
            table.Columns.Add("AUD1", typeof(double));
            table.Columns.Add("GBP1", typeof(double));
            table.Columns.Add("TDS1", typeof(double));

            // Trường hợp chọn tất cả thị trường
            if (marketID.Equals("0"))
            {
                foreach (string item in listMarket)
                {
                    // Cùng kì
                    ReportDetailtForTotalMoneyType dataItemLastYear = listDataGradation.Find(x => x.MarketCode == item && x.Year == (year - 1).ToString());
                    ReportDetailtForTotalMoneyType dataItemYear = listDataGradation.Find(x => x.MarketCode == item && x.Year == year.ToString());

                    // Trường hợp năm trước có đối tác và năm nay không có
                    if (dataItemLastYear != null && dataItemYear == null)
                    {
                        dataItemYear = new ReportDetailtForTotalMoneyType();
                        dataItemYear.MarketCode = dataItemLastYear.MarketCode;
                        dataItemYear.MarketName = dataItemLastYear.MarketName;
                        dataItemYear.VND = 0;
                        dataItemYear.USD = 0;
                        dataItemYear.EUR = 0;
                        dataItemYear.CAD = 0;
                        dataItemYear.AUD = 0;
                        dataItemYear.GBP = 0;
                        dataItemYear.Year = year.ToString();
                    }

                    // Trường hợp năm trước không có đối tác và năm nay có đối tác
                    if (dataItemYear != null && dataItemLastYear == null)
                    {
                        dataItemLastYear = new ReportDetailtForTotalMoneyType();
                        dataItemLastYear.MarketCode = dataItemYear.MarketCode;
                        dataItemLastYear.MarketName = dataItemYear.MarketName;
                        dataItemLastYear.VND = 0;
                        dataItemLastYear.USD = 0;
                        dataItemLastYear.EUR = 0;
                        dataItemLastYear.CAD = 0;
                        dataItemLastYear.AUD = 0;
                        dataItemLastYear.GBP = 0;
                        dataItemLastYear.Year = (year - 1).ToString();
                    }

                    double sumVND = dataItemYear.VND - dataItemLastYear.VND;
                    double sumUSD = dataItemYear.USD - dataItemLastYear.USD;
                    double sumEUR = dataItemYear.EUR - dataItemLastYear.EUR;
                    double sumCAD = dataItemYear.CAD - dataItemLastYear.CAD;
                    double sumAUD = dataItemYear.AUD - dataItemLastYear.AUD;
                    double sumGBP = dataItemYear.GBP - dataItemLastYear.GBP;
                    double sumTongDS = sumVND + sumUSD + sumEUR + sumCAD + sumAUD + sumGBP;

                    // add item vào table
                    table.Rows.Add(dataItemLastYear.MarketName
                        , Math.Round(sumVND, 2, MidpointRounding.ToEven)
                        , Math.Round(sumUSD, 2, MidpointRounding.ToEven)
                        , Math.Round(sumEUR, 2, MidpointRounding.ToEven)
                        , Math.Round(sumCAD, 2, MidpointRounding.ToEven)
                        , Math.Round(sumAUD, 2, MidpointRounding.ToEven)
                        , Math.Round(sumGBP, 2, MidpointRounding.ToEven)
                        , Math.Round(sumTongDS, 2, MidpointRounding.ToEven)
                        );
                }

                DataRow row = table.NewRow();
                row["MarketName"] = "Tổng";
                row["VND1"] = table.Compute("Sum(VND1)", "");
                row["USD1"] = table.Compute("Sum(USD1)", "");
                row["EUR1"] = table.Compute("Sum(EUR1)", "");
                row["CAD1"] = table.Compute("Sum(CAD1)", "");
                row["AUD1"] = table.Compute("Sum(AUD1)", "");
                row["GBP1"] = table.Compute("Sum(GBP1)", "");
                row["TDS1"] = table.Compute("Sum(TDS1)", "");
                table.Rows.Add(row);
            }
            else
            {
                // List thị trường
                listMarket = new List<string>();

                foreach (ReportDetailtForTotalMoneyType item in listDataGradation)
                {
                    if (!listMarket.Contains(item.MarketName))
                    {
                        listMarket.Add(item.MarketName);
                    }
                    item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;
                }

                foreach (string item in listMarket)
                {
                    List<ReportDetailtForTotalMoneyType> listDetailtYear = listDataGradation.Where(x => x.MarketName == item && x.Year == year.ToString()).ToList();
                    List<ReportDetailtForTotalMoneyType> listDetailtLastYear = listDataGradation.Where(x => x.MarketName == item && x.Year == (year - 1).ToString()).ToList();

                    // Year
                    double sumVNDYear = listDetailtYear.Sum(x => x.VND);
                    double sumUSDYear = listDetailtYear.Sum(x => x.USD);
                    double sumEURYear = listDetailtYear.Sum(x => x.EUR);
                    double sumCADYear = listDetailtYear.Sum(x => x.CAD);
                    double sumAUDYear = listDetailtYear.Sum(x => x.AUD);
                    double sumGBPYear = listDetailtYear.Sum(x => x.GBP);
                    double sumTongDSYear = sumVNDYear + sumUSDYear + sumEURYear + sumCADYear + sumAUDYear + sumGBPYear;

                    // LastYear
                    double sumVNDLastYear = listDetailtLastYear.Sum(x => x.VND);
                    double sumUSDLastYear = listDetailtLastYear.Sum(x => x.USD);
                    double sumEURLastYear = listDetailtLastYear.Sum(x => x.EUR);
                    double sumCADLastYear = listDetailtLastYear.Sum(x => x.CAD);
                    double sumAUDLastYear = listDetailtLastYear.Sum(x => x.AUD);
                    double sumGBPLastYear = listDetailtLastYear.Sum(x => x.GBP);
                    double sumTongDSLastYear = sumVNDLastYear + sumUSDLastYear + sumEURLastYear + sumCADLastYear + sumAUDLastYear + sumGBPLastYear;

                    double sumVND = sumVNDYear - sumVNDLastYear;
                    double sumUSD = sumUSDYear - sumUSDLastYear;
                    double sumEUR = sumEURYear - sumEURLastYear;
                    double sumCAD = sumCADYear - sumCADLastYear;
                    double sumAUD = sumAUDYear - sumAUDLastYear;
                    double sumGBP = sumGBPYear - sumGBPLastYear;
                    double sumTongDS = sumVND + sumUSD + sumEUR + sumCAD + sumAUD + sumGBP;

                    // add item vào table
                    table.Rows.Add(item
                        , Math.Round(sumVND, 2, MidpointRounding.ToEven)
                        , Math.Round(sumUSD, 2, MidpointRounding.ToEven)
                        , Math.Round(sumEUR, 2, MidpointRounding.ToEven)
                        , Math.Round(sumCAD, 2, MidpointRounding.ToEven)
                        , Math.Round(sumAUD, 2, MidpointRounding.ToEven)
                        , Math.Round(sumGBP, 2, MidpointRounding.ToEven)
                        , Math.Round(sumTongDS, 2, MidpointRounding.ToEven)
                        );
                }

                DataRow row = table.NewRow();
                row["MarketName"] = "Tổng";
                row["VND1"] = table.Compute("Sum(VND1)", "");
                row["USD1"] = table.Compute("Sum(USD1)", "");
                row["EUR1"] = table.Compute("Sum(EUR1)", "");
                row["CAD1"] = table.Compute("Sum(CAD1)", "");
                row["AUD1"] = table.Compute("Sum(AUD1)", "");
                row["GBP1"] = table.Compute("Sum(GBP1)", "");
                row["TDS1"] = table.Compute("Sum(TDS1)", "");
                table.Rows.Add(row);
            }

            // Tạo hearder
            title = "Thị trường";
            CreateTitle(string.Format("A{0}", totalRowTable1 + 3), string.Format("A{0}", totalRowTable1 + 4), sheetReport, title, 12, true);

            title = string.Format("Tăng/Giảm so với cùng kì năm {0}", year - 1);
            CreateTitle(string.Format("B{0}", totalRowTable1 + 3), string.Format("H{0}", totalRowTable1 + 3), sheetReport, title, 12, true);

            // Tạo cột hearder cho table3
            // VND
            title = "VND";
            CreateTitle(string.Format("B{0}", totalRowTable1 + 4), string.Format("B{0}", totalRowTable1 + 4), sheetReport, title, 12, true);

            // USD
            title = "USD";
            CreateTitle(string.Format("C{0}", totalRowTable1 + 4), string.Format("C{0}", totalRowTable1 + 4), sheetReport, title, 12, true);

            //EUR
            title = "EUR";
            CreateTitle(string.Format("D{0}", totalRowTable1 + 4), string.Format("D{0}", totalRowTable1 + 4), sheetReport, title, 12, true);

            //CAD
            title = "CAD";
            CreateTitle(string.Format("E{0}", totalRowTable1 + 4), string.Format("E{0}", totalRowTable1 + 4), sheetReport, title, 12, true);

            // AUD
            title = "AUD";
            CreateTitle(string.Format("F{0}", totalRowTable1 + 4), string.Format("F{0}", totalRowTable1 + 4), sheetReport, title, 12, true);

            //GBP
            title = "GBP";
            CreateTitle(string.Format("G{0}", totalRowTable1 + 4), string.Format("G{0}", totalRowTable1 + 4), sheetReport, title, 12, true);

            // Tổng
            title = "Tổng";
            CreateTitle(string.Format("H{0}", totalRowTable1 + 4), string.Format("H{0}", totalRowTable1 + 4), sheetReport, title, 12, true);

            // Set border
            style = new CellsFactory().CreateStyle();
            style.SetBorder(BorderType.BottomBorder, CellBorderType.Thin, Color.Black);
            style.SetBorder(BorderType.LeftBorder, CellBorderType.Thin, Color.Black);
            style.SetBorder(BorderType.RightBorder, CellBorderType.Thin, Color.Black);
            style.SetBorder(BorderType.TopBorder, CellBorderType.Thin, Color.Black);

            int totalRowTable2 = 0;

            if (table.Rows.Count > 0)
            {
                int stepRow = 0;
                // total row = row start + số row hiện có
                int totalRow = table.Rows.Count + totalRowTable1 + 4;
                totalRowTable2 = totalRow;
                // Số dòng của row
                for (int a = totalRowTable1 + 4; a < totalRow; a++)
                {
                    int stepColumn = 0;
                    // Số cột trong báo cáo cần hiển thị
                    // Tổng số cột hiển thị = Số cột hiển thị bắt đầu + tổng số cột cần hiển thị
                    int totalCol = 0 + 8;
                    for (int b = 0; b < totalCol; b++)
                    {
                        // Giá trị của value trong table
                        string valueOfTable = table.Rows[stepRow][stepColumn].ToString();

                        // Tô màu cho các dòng có giá trị tăng giảm
                        if (b > 0)
                        {
                            decimal tryParseValue = 0;
                            decimal.TryParse(valueOfTable, out tryParseValue);
                            style.Font.Color = Color.Green;

                            if (tryParseValue < 0)
                            {
                                style.Font.Color = Color.Red;
                            }
                        }

                        // Insert vào dòng cột xác định trong Excel
                        sheetReport.Cells[a, b].PutValue(valueOfTable, true);

                        // set style cho number
                        style.Custom = "#,##0.00";

                        // set border
                        sheetReport.Cells[a, b].SetStyle(style);

                        // Cột tổng cộng
                        if (b.Equals(totalCol - 1))
                        {
                            sheetReport.Cells[a, b].PutValue(valueOfTable, true, true);
                            style.Font.IsBold = true;
                            sheetReport.Cells[a, b].SetStyle(style);
                        }

                        // Trường hợp thuộc 2 dòng cuối
                        if (a.Equals(totalRow - 1))
                        {
                            sheetReport.Cells[a, b].PutValue(valueOfTable, true, true);
                            style.Font.IsBold = true;
                            sheetReport.Cells[a, b].SetStyle(style);
                        }

                        // Set lại giá trị mặt định
                        style.Font.IsBold = false;
                        // Tăng cột theo dòng của table
                        stepColumn++;
                    }
                    // Tăng dòng của table lên
                    stepRow++;

                    // Set lại color cho dòng hiện tại 
                    style.Font.Color = Color.Black;
                }
            }
            else
            {
                sheetReport.Cells["D10"].PutValue("Không có dữ liệu");
            }

            // Vẽ biểu đồ cột cho Excel
            //Chart reference
            Aspose.Cells.Charts.Chart leadSourceColumnVND;
            //Add Pie Chart
            // VND
            int chartIndex = sheetReport.Charts.Add(ChartType.Column3DClustered, 6, 0, 19, 6);
            leadSourceColumnVND = sheetReport.Charts[chartIndex];

            //Chart title
            leadSourceColumnVND.Title.Text = string.Format("Doanh số dịch vụ từng loại tiền - VND \n Giai đoạn: {0} \n", text);
            leadSourceColumnVND.Title.Font.Color = Color.Silver;

            // Adding SeriesCollection (chart data source) to the chart ranging from "A1" cell to "B3"
            string totalRowData = string.Format("B83:C{0}", 83 + table.Rows.Count - 2);
            leadSourceColumnVND.NSeries.Add(totalRowData, true);

            // Set the category data covering the range A2:A5.
            // Tổng số dòng cần hiển thị là số dòng hiện tại  + số dòng trong listData -2 (trong đó 1 là dòng tổng cuối cùng)
            string categoryData = "A83:A88";
            leadSourceColumnVND.NSeries.CategoryData = categoryData;

            // Set the names of the chart series taken from cells.
            leadSourceColumnVND.NSeries[0].Name = "=B82";
            leadSourceColumnVND.NSeries[1].Name = "=C82";

            // Set the 1st series fill color.
            leadSourceColumnVND.NSeries[0].Area.ForegroundColor = Color.Orange;
            leadSourceColumnVND.NSeries[0].Area.Formatting = FormattingType.Custom;

            // Set the 2nd series fill color.
            leadSourceColumnVND.NSeries[1].Area.ForegroundColor = Color.Green;
            leadSourceColumnVND.NSeries[1].Area.Formatting = FormattingType.Custom;


            // Set plot area formatting as none and hide its border.
            leadSourceColumnVND.PlotArea.Area.FillFormat.FillType = FillType.None;
            leadSourceColumnVND.PlotArea.Border.IsVisible = false;

            // Set value axis major tick mark as none and hide axis line. 
            // Also set the color of value axis major grid lines.
            leadSourceColumnVND.ValueAxis.MajorTickMark = TickMarkType.None;
            leadSourceColumnVND.ValueAxis.AxisLine.IsVisible = false;
            //leadSourceColumnChiNha.ValueAxis.IsAutomaticMajorUnit = false;
            //leadSourceColumnChiNha.ValueAxis.MajorUnit = 10000000;
            leadSourceColumnVND.ValueAxis.MajorGridLines.Color = Color.FromArgb(217, 217, 217);

            //Chart reference
            Aspose.Cells.Charts.Chart leadSourceColumnUSD;
            // USD
            chartIndex = sheetReport.Charts.Add(ChartType.Column3DClustered, 7, 7, 19, 13);
            leadSourceColumnUSD = sheetReport.Charts[chartIndex];

            //Chart title
            leadSourceColumnUSD.Title.Text = string.Format("Doanh số dịch vụ từng loại tiền - USD \n Giai đoạn: {0} \n", text);
            leadSourceColumnUSD.Title.Font.Color = Color.Silver;

            // Adding SeriesCollection (chart data source) to the chart ranging from "A1" cell to "B3"
            totalRowData = string.Format("D83:E{0}", 83 + table.Rows.Count - 2);
            leadSourceColumnUSD.NSeries.Add(totalRowData, true);

            // Set the category data covering the range A2:A5.
            // Tổng số dòng cần hiển thị là số dòng hiện tại  + số dòng trong listData -2 (trong đó 1 là dòng tổng cuối cùng)
            categoryData = "A83:A88";
            leadSourceColumnUSD.NSeries.CategoryData = categoryData;

            // Set the names of the chart series taken from cells.
            leadSourceColumnUSD.NSeries[0].Name = "=B82";
            leadSourceColumnUSD.NSeries[1].Name = "=C82";


            // Set the 1st series fill color.
            leadSourceColumnUSD.NSeries[0].Area.ForegroundColor = Color.Orange;
            leadSourceColumnUSD.NSeries[0].Area.Formatting = FormattingType.Custom;

            // Set the 2nd series fill color.
            leadSourceColumnUSD.NSeries[1].Area.ForegroundColor = Color.Green;
            leadSourceColumnUSD.NSeries[1].Area.Formatting = FormattingType.Custom;


            // Set plot area formatting as none and hide its border.
            leadSourceColumnUSD.PlotArea.Area.FillFormat.FillType = FillType.None;
            leadSourceColumnUSD.PlotArea.Border.IsVisible = false;

            // Set value axis major tick mark as none and hide axis line. 
            // Also set the color of value axis major grid lines.
            leadSourceColumnUSD.ValueAxis.MajorTickMark = TickMarkType.None;
            leadSourceColumnUSD.ValueAxis.AxisLine.IsVisible = false;
            leadSourceColumnUSD.ValueAxis.MajorGridLines.Color = Color.FromArgb(217, 217, 217);


            //Chart reference
            Aspose.Cells.Charts.Chart leadSourceColumnEUR;
            // EUR
            chartIndex = sheetReport.Charts.Add(ChartType.Column3DClustered, 20, 0, 33, 6);
            leadSourceColumnEUR = sheetReport.Charts[chartIndex];

            //Chart title
            leadSourceColumnEUR.Title.Text = string.Format("Doanh số dịch vụ từng loại tiền - EUR \n Giai đoạn: {0} \n", text);
            leadSourceColumnEUR.Title.Font.Color = Color.Silver;

            // Adding SeriesCollection (chart data source) to the chart ranging from "A1" cell to "B3"
            totalRowData = string.Format("F83:G{0}", 83 + table.Rows.Count - 2);
            leadSourceColumnEUR.NSeries.Add(totalRowData, true);

            // Set the category data covering the range A2:A5.
            // Tổng số dòng cần hiển thị là số dòng hiện tại  + số dòng trong listData -2 (trong đó 1 là dòng tổng cuối cùng)
            categoryData = "A83:A88";
            leadSourceColumnEUR.NSeries.CategoryData = categoryData;

            // Set the names of the chart series taken from cells.
            leadSourceColumnEUR.NSeries[0].Name = "=B82";
            leadSourceColumnEUR.NSeries[1].Name = "=C82";


            // Set the 1st series fill color.
            leadSourceColumnEUR.NSeries[0].Area.ForegroundColor = Color.Orange;
            leadSourceColumnEUR.NSeries[0].Area.Formatting = FormattingType.Custom;

            // Set the 2nd series fill color.
            leadSourceColumnEUR.NSeries[1].Area.ForegroundColor = Color.Green;
            leadSourceColumnEUR.NSeries[1].Area.Formatting = FormattingType.Custom;


            // Set plot area formatting as none and hide its border.
            leadSourceColumnEUR.PlotArea.Area.FillFormat.FillType = FillType.None;
            leadSourceColumnEUR.PlotArea.Border.IsVisible = false;

            // Set value axis major tick mark as none and hide axis line. 
            // Also set the color of value axis major grid lines.
            leadSourceColumnEUR.ValueAxis.MajorTickMark = TickMarkType.None;
            leadSourceColumnEUR.ValueAxis.AxisLine.IsVisible = false;
            leadSourceColumnEUR.ValueAxis.MajorGridLines.Color = Color.FromArgb(217, 217, 217);


            //Chart reference
            Aspose.Cells.Charts.Chart leadSourceColumnCAD;
            // CAD
            chartIndex = sheetReport.Charts.Add(ChartType.Column3DClustered, 20, 7, 33, 13);
            leadSourceColumnCAD = sheetReport.Charts[chartIndex];

            //Chart title
            leadSourceColumnCAD.Title.Text = string.Format("Doanh số dịch vụ từng loại tiền - CAD \n Giai đoạn: {0} \n", text);
            leadSourceColumnCAD.Title.Font.Color = Color.Silver;

            // Adding SeriesCollection (chart data source) to the chart ranging from "A1" cell to "B3"
            totalRowData = string.Format("H83:I{0}", 83 + table.Rows.Count - 2);
            leadSourceColumnCAD.NSeries.Add(totalRowData, true);

            // Set the category data covering the range A2:A5.
            // Tổng số dòng cần hiển thị là số dòng hiện tại  + số dòng trong listData -2 (trong đó 1 là dòng tổng cuối cùng)
            categoryData = "A83:A88";
            leadSourceColumnCAD.NSeries.CategoryData = categoryData;

            // Set the names of the chart series taken from cells.
            leadSourceColumnCAD.NSeries[0].Name = "=B82";
            leadSourceColumnCAD.NSeries[1].Name = "=C82";


            // Set the 1st series fill color.
            leadSourceColumnCAD.NSeries[0].Area.ForegroundColor = Color.Orange;
            leadSourceColumnCAD.NSeries[0].Area.Formatting = FormattingType.Custom;

            // Set the 2nd series fill color.
            leadSourceColumnCAD.NSeries[1].Area.ForegroundColor = Color.Green;
            leadSourceColumnCAD.NSeries[1].Area.Formatting = FormattingType.Custom;


            // Set plot area formatting as none and hide its border.
            leadSourceColumnCAD.PlotArea.Area.FillFormat.FillType = FillType.None;
            leadSourceColumnCAD.PlotArea.Border.IsVisible = false;

            // Set value axis major tick mark as none and hide axis line. 
            // Also set the color of value axis major grid lines.
            leadSourceColumnCAD.ValueAxis.MajorTickMark = TickMarkType.None;
            leadSourceColumnCAD.ValueAxis.AxisLine.IsVisible = false;
            leadSourceColumnCAD.ValueAxis.MajorGridLines.Color = Color.FromArgb(217, 217, 217);


            //Chart reference
            Aspose.Cells.Charts.Chart leadSourceColumnAUD;
            // AUD
            chartIndex = sheetReport.Charts.Add(ChartType.Column3DClustered, 34, 0, 47, 6);
            leadSourceColumnAUD = sheetReport.Charts[chartIndex];

            //Chart title
            leadSourceColumnAUD.Title.Text = string.Format("Doanh số dịch vụ từng loại tiền - AUD \n Giai đoạn: {0} \n", text);
            leadSourceColumnAUD.Title.Font.Color = Color.Silver;

            // Adding SeriesCollection (chart data source) to the chart ranging from "A1" cell to "B3"
            totalRowData = string.Format("J83:K{0}", 83 + table.Rows.Count - 2);
            leadSourceColumnAUD.NSeries.Add(totalRowData, true);

            // Set the category data covering the range A2:A5.
            // Tổng số dòng cần hiển thị là số dòng hiện tại  + số dòng trong listData -2 (trong đó 1 là dòng tổng cuối cùng)
            categoryData = "A83:A88";
            leadSourceColumnAUD.NSeries.CategoryData = categoryData;

            // Set the names of the chart series taken from cells.
            leadSourceColumnAUD.NSeries[0].Name = "=B82";
            leadSourceColumnAUD.NSeries[1].Name = "=C82";


            // Set the 1st series fill color.
            leadSourceColumnAUD.NSeries[0].Area.ForegroundColor = Color.Orange;
            leadSourceColumnAUD.NSeries[0].Area.Formatting = FormattingType.Custom;

            // Set the 2nd series fill color.
            leadSourceColumnAUD.NSeries[1].Area.ForegroundColor = Color.Green;
            leadSourceColumnAUD.NSeries[1].Area.Formatting = FormattingType.Custom;


            // Set plot area formatting as none and hide its border.
            leadSourceColumnAUD.PlotArea.Area.FillFormat.FillType = FillType.None;
            leadSourceColumnAUD.PlotArea.Border.IsVisible = false;

            // Set value axis major tick mark as none and hide axis line. 
            // Also set the color of value axis major grid lines.
            leadSourceColumnAUD.ValueAxis.MajorTickMark = TickMarkType.None;
            leadSourceColumnAUD.ValueAxis.AxisLine.IsVisible = false;
            leadSourceColumnAUD.ValueAxis.MajorGridLines.Color = Color.FromArgb(217, 217, 217);


            //Chart reference
            Aspose.Cells.Charts.Chart leadSourceColumnGBP;
            // AUD
            chartIndex = sheetReport.Charts.Add(ChartType.Column3DClustered, 34, 7, 47, 13);
            leadSourceColumnGBP = sheetReport.Charts[chartIndex];

            //Chart title
            leadSourceColumnGBP.Title.Text = string.Format("Doanh số dịch vụ từng loại tiền - GBP \n Giai đoạn: {0} \n", text);
            leadSourceColumnGBP.Title.Font.Color = Color.Silver;

            // Adding SeriesCollection (chart data source) to the chart ranging from "A1" cell to "B3"
            totalRowData = string.Format("L83:M{0}", 83 + table.Rows.Count - 2);
            leadSourceColumnGBP.NSeries.Add(totalRowData, true);

            // Set the category data covering the range A2:A5.
            // Tổng số dòng cần hiển thị là số dòng hiện tại  + số dòng trong listData -2 (trong đó 1 là dòng tổng cuối cùng)
            categoryData = "A83:A88";
            leadSourceColumnGBP.NSeries.CategoryData = categoryData;

            // Set the names of the chart series taken from cells.
            leadSourceColumnGBP.NSeries[0].Name = "=B82";
            leadSourceColumnGBP.NSeries[1].Name = "=C82";


            // Set the 1st series fill color.
            leadSourceColumnGBP.NSeries[0].Area.ForegroundColor = Color.Orange;
            leadSourceColumnGBP.NSeries[0].Area.Formatting = FormattingType.Custom;

            // Set the 2nd series fill color.
            leadSourceColumnGBP.NSeries[1].Area.ForegroundColor = Color.Green;
            leadSourceColumnGBP.NSeries[1].Area.Formatting = FormattingType.Custom;


            // Set plot area formatting as none and hide its border.
            leadSourceColumnGBP.PlotArea.Area.FillFormat.FillType = FillType.None;
            leadSourceColumnGBP.PlotArea.Border.IsVisible = false;

            // Set value axis major tick mark as none and hide axis line. 
            // Also set the color of value axis major grid lines.
            leadSourceColumnGBP.ValueAxis.MajorTickMark = TickMarkType.None;
            leadSourceColumnGBP.ValueAxis.AxisLine.IsVisible = false;
            leadSourceColumnGBP.ValueAxis.MajorGridLines.Color = Color.FromArgb(217, 217, 217);

            // Tạo chart cột tỉ trọng cho các thị trường
            //Add Pie Chart
            //Chart reference
            Aspose.Cells.Charts.Chart leadSourceColumn;
            chartIndex = sheetReport.Charts.Add(ChartType.Column3DStacked, 50, 0, 77, 13);
            leadSourceColumn = sheetReport.Charts[chartIndex];


            //Chart title
            leadSourceColumn.Title.Text = string.Format("Tỉ trọng từng thị trường \n Giai đoạn: {0} \n", text);
            leadSourceColumn.Title.Font.Color = Color.Silver;

            // count cho 3 loại: chi Quầy, Chi nhà, Chuyển khoản
            // list thị trường

            listDataGradation = new ReportBL().ReportDetailtMTGradationCompareForAllPercent(year, int.Parse(gradationID), reportTypeID, marketID);

            List<string> listMarketCurrent = new List<string>();
            foreach (ReportDetailtForTotalMoneyType item in listDataGradation)
            {
                if (!listMarketCurrent.Contains(item.MarketName))
                {
                    listMarketCurrent.Add(item.MarketName);
                }
            }

            // List dữ liệu dataRow
            string[] listTotalRowData = new string[listMarketCurrent.Count];
            int i = 0;
            foreach (string item in listMarketCurrent)
            {
                // Năm trước
                List<ReportDetailtForTotalMoneyType> dataItemLastYear = listDataGradation.Where(x => x.MarketName == item && x.Year == (year - 1).ToString()).ToList();
                List<ReportDetailtForTotalMoneyType> dataItemYear = listDataGradation.Where(x => x.MarketName == item && x.Year == year.ToString()).ToList();


                listTotalRowData[i++] = string.Concat("{"
                    , string.Format("{0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9}, {10}, {11}"
                    , dataItemYear.Sum(x => x.VND), dataItemLastYear.Sum(x => x.VND)
                    , dataItemYear.Sum(x => x.USD), dataItemLastYear.Sum(x => x.USD)
                    , dataItemYear.Sum(x => x.EUR), dataItemLastYear.Sum(x => x.EUR)
                    , dataItemYear.Sum(x => x.CAD), dataItemLastYear.Sum(x => x.CAD)
                    , dataItemYear.Sum(x => x.AUD), dataItemLastYear.Sum(x => x.AUD)
                    , dataItemYear.Sum(x => x.GBP), dataItemLastYear.Sum(x => x.GBP))
                    , "}");
            }

            int count = 0;
            string yearCurent = string.Format("Năm {0}", year);
            string lastYear = string.Format("Năm {0}", year - 1);
            foreach (string item in listTotalRowData)
            {
                totalRowData = item;
                leadSourceColumn.NSeries.Add(totalRowData, true);

                categoryData = string.Concat("{", string.Format("VND {0}, VND {1}, USD {0}, USD {1}, EUR {0},  EUR {1}, CAD {0}, CAD {1}, AUD {0}, AUD {1}, GBP {0}, GBP {1}", yearCurent, lastYear), "}");
                leadSourceColumn.NSeries.CategoryData = categoryData;

                leadSourceColumn.NSeries[count].Name = listMarketCurrent[count];

                switch (count)
                {
                    case 0:
                        leadSourceColumn.NSeries[count].Area.ForegroundColor = Color.Orange;
                        leadSourceColumn.NSeries[count].Area.Formatting = FormattingType.Custom;
                        break;
                    case 1:
                        leadSourceColumn.NSeries[count].Area.ForegroundColor = Color.Green;
                        leadSourceColumn.NSeries[count].Area.Formatting = FormattingType.Custom;
                        break;
                    case 2:
                        leadSourceColumn.NSeries[count].Area.ForegroundColor = Color.Blue;
                        leadSourceColumn.NSeries[count].Area.Formatting = FormattingType.Custom;
                        break;
                    case 3:
                        leadSourceColumn.NSeries[count].Area.ForegroundColor = Color.Yellow;
                        leadSourceColumn.NSeries[count].Area.Formatting = FormattingType.Custom;
                        break;
                    case 4:
                        leadSourceColumn.NSeries[count].Area.ForegroundColor = Color.Brown;
                        leadSourceColumn.NSeries[count].Area.Formatting = FormattingType.Custom;
                        break;
                    case 5:
                        leadSourceColumn.NSeries[count].Area.ForegroundColor = Color.Olive;
                        leadSourceColumn.NSeries[count].Area.Formatting = FormattingType.Custom;
                        break;
                    default:
                        leadSourceColumn.NSeries[count].Area.ForegroundColor = Color.Pink;
                        leadSourceColumn.NSeries[count].Area.Formatting = FormattingType.Custom;
                        break;
                }
                count++;
            }

            // Set plot area formatting as none and hide its border.
            leadSourceColumn.PlotArea.Area.FillFormat.FillType = FillType.None;
            leadSourceColumn.PlotArea.Border.IsVisible = false;

            // Set value axis major tick mark as none and hide axis line. 
            // Also set the color of value axis major grid lines.
            leadSourceColumn.ValueAxis.MajorTickMark = TickMarkType.None;
            leadSourceColumn.ValueAxis.MaxValue = 100;
            leadSourceColumn.ValueAxis.AxisLine.IsVisible = false;
            leadSourceColumn.ValueAxis.MajorGridLines.Color = Color.FromArgb(217, 217, 217);

            // Chạy process
            designer.Process();
            return ExportReport("ReportGradationCompare", designer);
        }

        /// <summary>
        /// Tạo mẫu cho Excel cho so sánh theo giai đoạn cho tất cả
        /// </summary>
        /// <param name="gradationID"></param>
        /// <param name="year"></param>
        /// <param name="typeID"></param>
        /// <returns></returns>
        public ActionResult CreateExcelForGradationCompareForOne(string gradationID, int year, string reportTypeID, string marketID)
        {
            WorkbookDesigner designer = new WorkbookDesigner();
            string templatePath = "~/Content/Report/ReportDetailPartnerForGaration.xlsx";
            // Get đường dẫn
            templatePath = System.Web.HttpContext.Current.Server.MapPath(templatePath);

            designer.Workbook = new DongA.Core.DongAExcel().OpenExcelFile(templatePath);
            designer.Workbook.CalculateFormula();

            WorksheetCollection workSheets = designer.Workbook.Worksheets;
            Worksheet sheetReport = designer.Workbook.Worksheets[0];

            // Tạo title
            string typeReport = "So sánh - Theo giai đoạn - Tất cả";

            string text = string.Format(" tháng năm {0}", year);
            string textValue = "T";
            switch (gradationID)
            {
                case "1":
                    text = string.Concat("3", text);
                    textValue = string.Concat("3", textValue);
                    break;
                case "2":
                    text = string.Concat("6", text);
                    textValue = string.Concat("6", textValue);
                    break;
                case "3":
                    text = string.Concat("9", text);
                    textValue = string.Concat("9", textValue);
                    break;
                default:
                    text = string.Concat("12", text);
                    textValue = string.Concat("12", textValue);
                    break;
            }

            // Tạo title
            CreateTitle("A2", "K2", sheetReport, typeReport, 14);

            // Tạo title detailt
            string titleDetailt = string.Format("Giai đoạn: {0}", text);
            CreateTitle("A3", "K3", sheetReport, titleDetailt, 12);


            // Tạo giá trị cho cột dữ liệu của Chi quầy/ Chi nhà/ Chuyển khoản
            sheetReport.Cells["C62"].PutValue(string.Format("Năm {0} ", year - 1));
            sheetReport.Cells["D62"].PutValue(string.Format("Năm {0} ", year));

            sheetReport.Cells["E62"].PutValue(string.Format("Năm {0} ", year - 1));
            sheetReport.Cells["F62"].PutValue(string.Format("Năm {0} ", year));

            sheetReport.Cells["G62"].PutValue(string.Format("Năm {0} ", year - 1));
            sheetReport.Cells["H62"].PutValue(string.Format("Năm {0} ", year));

            sheetReport.Cells["I62"].PutValue(string.Format("Năm {0} ", year - 1));
            sheetReport.Cells["J62"].PutValue(string.Format("Năm {0} ", year));

            List<ReportDetailtForTotalMoneyType> listDataGradation = new ReportBL().ReportDetailtMTGradationCompareForOneConvert(year, int.Parse(gradationID), reportTypeID, marketID);

            if (marketID.Contains("005"))
            {
                List<string> listMarket = new List<string>();
                List<ReportDetailtForTotalMoneyType> listDataGradationConvert = new List<ReportDetailtForTotalMoneyType>();

                foreach (ReportDetailtForTotalMoneyType item in listDataGradation)
                {
                    if (!listMarket.Contains(item.MarketName))
                    {
                        listMarket.Add(item.MarketName);
                    }
                }

                foreach (string item in listMarket)
                {
                    List<ReportDetailtForTotalMoneyType> listDataItemYear = listDataGradation.Where(x => x.MarketName == item && x.Year == year.ToString()).ToList();
                    List<ReportDetailtForTotalMoneyType> listDataItemLastYear = listDataGradation.Where(x => x.MarketName == item && x.Year == (year - 1).ToString()).ToList();

                    // Year
                    listDataGradationConvert.Add(
                        new ReportDetailtForTotalMoneyType()
                        {
                            PartnerName = item,
                            MarketName = "Thị trường Châu Á",
                            VND = listDataItemYear.Sum(x => x.VND),
                            USD = listDataItemYear.Sum(x => x.USD),
                            EUR = listDataItemYear.Sum(x => x.EUR),
                            CAD = listDataItemYear.Sum(x => x.CAD),
                            AUD = listDataItemYear.Sum(x => x.AUD),
                            GBP = listDataItemYear.Sum(x => x.GBP),
                            Year = year.ToString()
                        }
                    );

                    // Last Year
                    listDataGradationConvert.Add(
                        new ReportDetailtForTotalMoneyType()
                        {
                            PartnerName = item,
                            MarketName = "Thị trường Châu Á",
                            VND = listDataItemLastYear.Sum(x => x.VND),
                            USD = listDataItemLastYear.Sum(x => x.USD),
                            EUR = listDataItemLastYear.Sum(x => x.EUR),
                            CAD = listDataItemLastYear.Sum(x => x.CAD),
                            AUD = listDataItemLastYear.Sum(x => x.AUD),
                            GBP = listDataItemLastYear.Sum(x => x.GBP),
                            Year = (year - 1).ToString()
                        }
                    );
                }

                if (listDataGradationConvert.Count > 0)
                {
                    listDataGradation = new List<ReportDetailtForTotalMoneyType>(listDataGradationConvert);
                }
            }

            foreach (ReportDetailtForTotalMoneyType item in listDataGradation)
            {
                item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;
            }

            // Khởi tạo datatable
            DataTable table = new DataTable();
            // Tạo các cột cho datatable
            table.Columns.Add("PartnerName", typeof(String));
            table.Columns.Add("VND1", typeof(double));
            table.Columns.Add("VND2", typeof(double));

            table.Columns.Add("USD1", typeof(double));
            table.Columns.Add("USD2", typeof(double));

            table.Columns.Add("EUR1", typeof(double));
            table.Columns.Add("EUR2", typeof(double));

            table.Columns.Add("CAD1", typeof(double));
            table.Columns.Add("CAD2", typeof(double));

            table.Columns.Add("AUD1", typeof(double));
            table.Columns.Add("AUD2", typeof(double));

            table.Columns.Add("GBP1", typeof(double));
            table.Columns.Add("GBP2", typeof(double));

            table.Columns.Add("TDS1", typeof(double));
            table.Columns.Add("TDS2", typeof(double));

            table.Columns.Add("MarketName", typeof(String));


            foreach (ReportDetailtForTotalMoneyType item in listDataGradation)
            {

                // Cùng kì
                ReportDetailtForTotalMoneyType dataItemLastYear = listDataGradation.Find(x => x.PartnerName == item.PartnerName && x.Year == (year - 1).ToString());
                ReportDetailtForTotalMoneyType dataItemYear = listDataGradation.Find(x => x.PartnerName == item.PartnerName && x.Year == year.ToString());

                // Trường hợp năm trước có đối tác và năm nay không có
                if (dataItemLastYear != null && dataItemYear == null)
                {
                    dataItemYear = new ReportDetailtForTotalMoneyType();
                    dataItemYear.PartnerName = dataItemLastYear.PartnerName;
                    dataItemYear.Year = year.ToString();
                }

                // Trường hợp năm trước không có đối tác và năm nay có đối tác
                if (dataItemYear != null && dataItemLastYear == null)
                {
                    dataItemLastYear = new ReportDetailtForTotalMoneyType();
                    dataItemLastYear.PartnerName = dataItemYear.PartnerName;
                    dataItemLastYear.Year = (year - 1).ToString();
                }

                // Check tồn tại của item
                string value = string.Format("PartnerName='{0}'", item.PartnerName);
                DataRow[] foundRows = table.Select(value);
                if (dataItemLastYear != null && dataItemYear != null && foundRows.Count() == 0)
                {
                    // add item vào table
                    table.Rows.Add(item.PartnerName
                        , dataItemLastYear.VND, dataItemYear.VND, dataItemLastYear.USD, dataItemYear.USD, dataItemLastYear.EUR, dataItemYear.EUR
                        , dataItemLastYear.CAD, dataItemYear.CAD, dataItemLastYear.AUD, dataItemYear.AUD, dataItemLastYear.GBP, dataItemYear.GBP
                        , dataItemLastYear.TongDS, dataItemYear.TongDS
                        , item.MarketName
                        );
                }
            }

            DataRow row = table.NewRow();
            row["PartnerName"] = "Tổng";
            row["VND1"] = table.Compute("Sum(VND1)", "");
            row["VND2"] = table.Compute("Sum(VND2)", "");

            row["USD1"] = table.Compute("Sum(USD1)", "");
            row["USD2"] = table.Compute("Sum(USD2)", "");

            row["EUR1"] = table.Compute("Sum(EUR1)", "");
            row["EUR2"] = table.Compute("Sum(EUR2)", "");

            row["CAD1"] = table.Compute("Sum(CAD1)", "");
            row["CAD2"] = table.Compute("Sum(CAD2)", "");

            row["AUD1"] = table.Compute("Sum(AUD1)", "");
            row["AUD2"] = table.Compute("Sum(AUD2)", "");

            row["GBP1"] = table.Compute("Sum(GBP1)", "");
            row["GBP2"] = table.Compute("Sum(GBP2)", "");

            row["TDS1"] = table.Compute("Sum(TDS1)", "");
            row["TDS2"] = table.Compute("Sum(TDS2)", "");
            row["MarketName"] = "";

            table.Rows.Add(row);

            // Tạo cột hearder cho table3
            string title = "Đối tác";
            CreateTitle("A81", "A82", sheetReport, title, 12, true);

            // VND
            title = string.Format("{0} {1}", textValue, year - 1);
            CreateTitle("B82", "B82", sheetReport, title, 12, true);

            title = string.Format("{0} {1}", textValue, year);
            CreateTitle("C82", "C82", sheetReport, title, 12, true);

            // USD
            title = string.Format("{0} {1}", textValue, year - 1);
            CreateTitle("D82", "D82", sheetReport, title, 12, true);

            title = string.Format("{0} {1}", textValue, year);
            CreateTitle("E82", "E82", sheetReport, title, 12, true);

            //EUR
            title = string.Format("{0} {1}", textValue, year - 1);
            CreateTitle("F82", "F82", sheetReport, title, 12, true);

            title = string.Format("{0} {1}", textValue, year);
            CreateTitle("G82", "G82", sheetReport, title, 12, true);

            //CAD
            title = string.Format("{0} {1}", textValue, year - 1);
            CreateTitle("H82", "H82", sheetReport, title, 12, true);

            title = string.Format("{0} {1}", textValue, year);
            CreateTitle("I82", "I82", sheetReport, title, 12, true);

            // AUD
            title = string.Format("{0} {1}", textValue, year - 1);
            CreateTitle("J82", "J82", sheetReport, title, 12, true);

            title = string.Format("{0} {1}", textValue, year);
            CreateTitle("K82", "K82", sheetReport, title, 12, true);

            //GBP
            title = string.Format("{0} {1}", textValue, year - 1);
            CreateTitle("L82", "L82", sheetReport, title, 12, true);

            title = string.Format("{0} {1}", textValue, year);
            CreateTitle("M82", "M82", sheetReport, title, 12, true);

            // Tổng
            title = string.Format("{0} {1}", textValue, year - 1);
            CreateTitle("N82", "N82", sheetReport, title, 12, true);

            title = string.Format("{0} {1}", textValue, year);
            CreateTitle("O82", "O82", sheetReport, title, 12, true);



            // Set border
            Style style = new CellsFactory().CreateStyle();
            style.SetBorder(BorderType.BottomBorder, CellBorderType.Thin, Color.Black);
            style.SetBorder(BorderType.LeftBorder, CellBorderType.Thin, Color.Black);
            style.SetBorder(BorderType.RightBorder, CellBorderType.Thin, Color.Black);

            int totalRowTable1 = 0;

            if (table.Rows.Count > 0)
            {
                int stepRow = 0;
                // total row = row start + số row hiện có
                int totalRow = table.Rows.Count + 82;
                totalRowTable1 = totalRow;
                // Số dòng của row
                for (int a = 82; a < totalRow; a++)
                {
                    int stepColumn = 0;
                    // Số cột trong báo cáo cần hiển thị
                    // Tổng số cột hiển thị = Số cột hiển thị bắt đầu + tổng số cột cần hiển thị
                    int totalCol = 0 + 15;
                    for (int b = 0; b < totalCol; b++)
                    {
                        // Giá trị của value trong table
                        string valueOfTable = table.Rows[stepRow][stepColumn].ToString();

                        // Tô màu cho các dòng có giá trị tăng giảm
                        if (b >= 18)
                        {
                            decimal tryParseValue = 0;
                            decimal.TryParse(valueOfTable, out tryParseValue);
                            style.Font.Color = Color.Green;

                            if (tryParseValue < 0)
                            {
                                style.Font.Color = Color.Red;
                            }
                        }

                        // Insert vào dòng cột xác định trong Excel
                        sheetReport.Cells[a, b].PutValue(valueOfTable, true);

                        // set style cho number
                        style.Custom = "#,##0.00";

                        // set border
                        sheetReport.Cells[a, b].SetStyle(style);

                        // Cột tổng cộng
                        if (b.Equals(totalCol - 1))
                        {
                            sheetReport.Cells[a, b].PutValue(valueOfTable, true, true);
                            style.Font.IsBold = true;
                            sheetReport.Cells[a, b].SetStyle(style);
                        }

                        // Trường hợp thuộc 2 dòng cuối
                        if (a.Equals(totalRow - 1))
                        {
                            sheetReport.Cells[a, b].PutValue(valueOfTable, true, true);
                            style.Font.IsBold = true;
                            sheetReport.Cells[a, b].SetStyle(style);
                        }

                        // Set lại giá trị mặt định
                        style.Font.IsBold = false;
                        // Tăng cột theo dòng của table
                        stepColumn++;
                    }
                    // Tăng dòng của table lên
                    stepRow++;

                    // Set lại color cho dòng hiện tại 
                    style.Font.Color = Color.Black;
                }
            }
            else
            {
                sheetReport.Cells["D10"].PutValue("Không có dữ liệu");
            }


            // Tăng/Giảm so với cùng kì năm trước
            listDataGradation = new ReportBL().ReportDetailtMTGradationCompareForOneConvert(year, int.Parse(gradationID), reportTypeID, marketID);

            if (marketID.Contains("005"))
            {
                List<string> listMarket = new List<string>();
                List<ReportDetailtForTotalMoneyType> listDataGradationConvert = new List<ReportDetailtForTotalMoneyType>();

                foreach (ReportDetailtForTotalMoneyType item in listDataGradation)
                {
                    if (!listMarket.Contains(item.MarketName))
                    {
                        listMarket.Add(item.MarketName);
                    }
                }

                foreach (string item in listMarket)
                {
                    List<ReportDetailtForTotalMoneyType> listDataItemYear = listDataGradation.Where(x => x.MarketName == item && x.Year == year.ToString()).ToList();
                    List<ReportDetailtForTotalMoneyType> listDataItemLastYear = listDataGradation.Where(x => x.MarketName == item && x.Year == (year - 1).ToString()).ToList();

                    // Year
                    listDataGradationConvert.Add(
                        new ReportDetailtForTotalMoneyType()
                        {
                            PartnerName = item,
                            MarketName = "Thị trường Châu Á",
                            VND = listDataItemYear.Sum(x => x.VND),
                            USD = listDataItemYear.Sum(x => x.USD),
                            EUR = listDataItemYear.Sum(x => x.EUR),
                            CAD = listDataItemYear.Sum(x => x.CAD),
                            AUD = listDataItemYear.Sum(x => x.AUD),
                            GBP = listDataItemYear.Sum(x => x.GBP),
                            Year = year.ToString()
                        }
                    );

                    // Last Year
                    listDataGradationConvert.Add(
                        new ReportDetailtForTotalMoneyType()
                        {
                            PartnerName = item,
                            MarketName = "Thị trường Châu Á",
                            VND = listDataItemLastYear.Sum(x => x.VND),
                            USD = listDataItemLastYear.Sum(x => x.USD),
                            EUR = listDataItemLastYear.Sum(x => x.EUR),
                            CAD = listDataItemLastYear.Sum(x => x.CAD),
                            AUD = listDataItemLastYear.Sum(x => x.AUD),
                            GBP = listDataItemLastYear.Sum(x => x.GBP),
                            Year = (year - 1).ToString()
                        }
                    );
                }

                if (listDataGradationConvert.Count > 0)
                {
                    listDataGradation = new List<ReportDetailtForTotalMoneyType>(listDataGradationConvert);
                }
            }

            foreach (ReportDetailtForTotalMoneyType item in listDataGradation)
            {
                item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;
            }

            // Khởi tạo datatable
            table = new DataTable();
            // Tạo các cột cho datatable
            table.Columns.Add("PartnerName", typeof(String));
            table.Columns.Add("VND1", typeof(double));
            table.Columns.Add("USD1", typeof(double));
            table.Columns.Add("EUR1", typeof(double));
            table.Columns.Add("CAD1", typeof(double));
            table.Columns.Add("AUD1", typeof(double));
            table.Columns.Add("GBP1", typeof(double));
            table.Columns.Add("TDS1", typeof(double));

            table.Columns.Add("MarketName", typeof(string));

            foreach (ReportDetailtForTotalMoneyType item in listDataGradation)
            {

                // Cùng kì
                ReportDetailtForTotalMoneyType dataItemLastYear = listDataGradation.Find(x => x.PartnerName == item.PartnerName && x.Year == (year - 1).ToString());
                ReportDetailtForTotalMoneyType dataItemYear = listDataGradation.Find(x => x.PartnerName == item.PartnerName && x.Year == year.ToString());

                // Trường hợp năm trước có đối tác và năm nay không có
                if (dataItemLastYear != null && dataItemYear == null)
                {
                    dataItemYear = new ReportDetailtForTotalMoneyType();
                    dataItemYear.PartnerName = dataItemLastYear.PartnerName;
                    dataItemYear.Year = year.ToString();
                }

                // Trường hợp năm trước không có đối tác và năm nay có đối tác
                if (dataItemYear != null && dataItemLastYear == null)
                {
                    dataItemLastYear = new ReportDetailtForTotalMoneyType();
                    dataItemLastYear.PartnerName = dataItemYear.PartnerName;
                    dataItemLastYear.Year = (year - 1).ToString();
                }

                double sumVND = dataItemYear.VND - dataItemLastYear.VND;
                double sumUSD = dataItemYear.USD - dataItemLastYear.USD;
                double sumEUR = dataItemYear.EUR - dataItemLastYear.EUR;
                double sumCAD = dataItemYear.CAD - dataItemLastYear.CAD;
                double sumAUD = dataItemYear.AUD - dataItemLastYear.AUD;
                double sumGBP = dataItemYear.GBP - dataItemLastYear.GBP;
                double sumTongDS = sumVND + sumUSD + sumEUR + sumCAD + sumAUD + sumGBP;

                // Check tồn tại của item
                string value = string.Format("PartnerName='{0}'", item.PartnerName);
                DataRow[] foundRows = table.Select(value);
                if (dataItemLastYear != null && dataItemYear != null && foundRows.Count() == 0)
                {
                    // add item vào table
                    table.Rows.Add(item.PartnerName
                        , Math.Round(sumVND, 2, MidpointRounding.ToEven)
                        , Math.Round(sumUSD, 2, MidpointRounding.ToEven)
                        , Math.Round(sumEUR, 2, MidpointRounding.ToEven)
                        , Math.Round(sumCAD, 2, MidpointRounding.ToEven)
                        , Math.Round(sumAUD, 2, MidpointRounding.ToEven)
                        , Math.Round(sumGBP, 2, MidpointRounding.ToEven)
                        , Math.Round(sumTongDS, 2, MidpointRounding.ToEven)
                        , item.MarketName
                        );
                }
            }

            row = table.NewRow();
            row["PartnerName"] = "Tổng";
            row["VND1"] = table.Compute("Sum(VND1)", "");
            row["USD1"] = table.Compute("Sum(USD1)", "");
            row["EUR1"] = table.Compute("Sum(EUR1)", "");
            row["CAD1"] = table.Compute("Sum(CAD1)", "");
            row["AUD1"] = table.Compute("Sum(AUD1)", "");
            row["GBP1"] = table.Compute("Sum(GBP1)", "");
            row["TDS1"] = table.Compute("Sum(TDS1)", "");
            row["MarketName"] = "";

            table.Rows.Add(row);

            // Tạo hearder
            title = "Đối tác";
            CreateTitle(string.Format("A{0}", totalRowTable1 + 3), string.Format("A{0}", totalRowTable1 + 4), sheetReport, title, 12, true);

            title = string.Format("Tăng/Giảm so với cùng kì năm {0}", year - 1);
            CreateTitle(string.Format("B{0}", totalRowTable1 + 3), string.Format("H{0}", totalRowTable1 + 3), sheetReport, title, 12, true);

            // Tạo cột hearder cho table3
            // VND
            title = "VND";
            CreateTitle(string.Format("B{0}", totalRowTable1 + 4), string.Format("B{0}", totalRowTable1 + 4), sheetReport, title, 12, true);

            // USD
            title = "USD";
            CreateTitle(string.Format("C{0}", totalRowTable1 + 4), string.Format("C{0}", totalRowTable1 + 4), sheetReport, title, 12, true);

            //EUR
            title = "EUR";
            CreateTitle(string.Format("D{0}", totalRowTable1 + 4), string.Format("D{0}", totalRowTable1 + 4), sheetReport, title, 12, true);

            //CAD
            title = "CAD";
            CreateTitle(string.Format("E{0}", totalRowTable1 + 4), string.Format("E{0}", totalRowTable1 + 4), sheetReport, title, 12, true);

            // AUD
            title = "AUD";
            CreateTitle(string.Format("F{0}", totalRowTable1 + 4), string.Format("F{0}", totalRowTable1 + 4), sheetReport, title, 12, true);

            //GBP
            title = "GBP";
            CreateTitle(string.Format("G{0}", totalRowTable1 + 4), string.Format("G{0}", totalRowTable1 + 4), sheetReport, title, 12, true);

            // Tổng
            title = "Tổng";
            CreateTitle(string.Format("H{0}", totalRowTable1 + 4), string.Format("H{0}", totalRowTable1 + 4), sheetReport, title, 12, true);

            // Set border
            style = new CellsFactory().CreateStyle();
            style.SetBorder(BorderType.BottomBorder, CellBorderType.Thin, Color.Black);
            style.SetBorder(BorderType.LeftBorder, CellBorderType.Thin, Color.Black);
            style.SetBorder(BorderType.RightBorder, CellBorderType.Thin, Color.Black);
            style.SetBorder(BorderType.TopBorder, CellBorderType.Thin, Color.Black);

            int totalRowTable2 = 0;

            if (table.Rows.Count > 0)
            {
                int stepRow = 0;
                // total row = row start + số row hiện có
                int totalRow = table.Rows.Count + totalRowTable1 + 4;
                totalRowTable2 = totalRow;
                // Số dòng của row
                for (int a = totalRowTable1 + 4; a < totalRow; a++)
                {
                    int stepColumn = 0;
                    // Số cột trong báo cáo cần hiển thị
                    // Tổng số cột hiển thị = Số cột hiển thị bắt đầu + tổng số cột cần hiển thị
                    int totalCol = 0 + 8;
                    for (int b = 0; b < totalCol; b++)
                    {
                        // Giá trị của value trong table
                        string valueOfTable = table.Rows[stepRow][stepColumn].ToString();

                        // Tô màu cho các dòng có giá trị tăng giảm
                        if (b > 0)
                        {
                            decimal tryParseValue = 0;
                            decimal.TryParse(valueOfTable, out tryParseValue);
                            style.Font.Color = Color.Green;

                            if (tryParseValue < 0)
                            {
                                style.Font.Color = Color.Red;
                            }
                        }

                        // Insert vào dòng cột xác định trong Excel
                        sheetReport.Cells[a, b].PutValue(valueOfTable, true);

                        // set style cho number
                        style.Custom = "#,##0.00";

                        // set border
                        sheetReport.Cells[a, b].SetStyle(style);

                        // Cột tổng cộng
                        if (b.Equals(totalCol - 1))
                        {
                            sheetReport.Cells[a, b].PutValue(valueOfTable, true, true);
                            style.Font.IsBold = true;
                            sheetReport.Cells[a, b].SetStyle(style);
                        }

                        // Trường hợp thuộc 2 dòng cuối
                        if (a.Equals(totalRow - 1))
                        {
                            sheetReport.Cells[a, b].PutValue(valueOfTable, true, true);
                            style.Font.IsBold = true;
                            sheetReport.Cells[a, b].SetStyle(style);
                        }

                        // Set lại giá trị mặt định
                        style.Font.IsBold = false;
                        // Tăng cột theo dòng của table
                        stepColumn++;
                    }
                    // Tăng dòng của table lên
                    stepRow++;

                    // Set lại color cho dòng hiện tại 
                    style.Font.Color = Color.Black;
                }
            }
            else
            {
                sheetReport.Cells["D10"].PutValue("Không có dữ liệu");
            }

            // Vẽ biểu đồ cột cho Excel
            //Chart reference
            Aspose.Cells.Charts.Chart leadSourceColumnVND;
            //Add Pie Chart
            // VND
            int chartIndex = sheetReport.Charts.Add(ChartType.Column3DClustered, 6, 0, 19, 6);
            leadSourceColumnVND = sheetReport.Charts[chartIndex];

            //Chart title
            leadSourceColumnVND.Title.Text = string.Format("Doanh số dịch vụ từng loại tiền - VND \n Giai đoạn: {0} \n", text);
            leadSourceColumnVND.Title.Font.Color = Color.Silver;

            // Adding SeriesCollection (chart data source) to the chart ranging from "A1" cell to "B3"
            string totalRowData = string.Format("B83:C{0}", 83 + table.Rows.Count - 2);
            leadSourceColumnVND.NSeries.Add(totalRowData, true);

            // Set the category data covering the range A2:A5.
            // Tổng số dòng cần hiển thị là số dòng hiện tại  + số dòng trong listData -2 (trong đó 1 là dòng tổng cuối cùng)
            string categoryData = "A83:A88";
            leadSourceColumnVND.NSeries.CategoryData = categoryData;

            // Set the names of the chart series taken from cells.
            leadSourceColumnVND.NSeries[0].Name = "=B82";
            leadSourceColumnVND.NSeries[1].Name = "=C82";

            // Set the 1st series fill color.
            leadSourceColumnVND.NSeries[0].Area.ForegroundColor = Color.Orange;
            leadSourceColumnVND.NSeries[0].Area.Formatting = FormattingType.Custom;

            // Set the 2nd series fill color.
            leadSourceColumnVND.NSeries[1].Area.ForegroundColor = Color.Green;
            leadSourceColumnVND.NSeries[1].Area.Formatting = FormattingType.Custom;


            // Set plot area formatting as none and hide its border.
            leadSourceColumnVND.PlotArea.Area.FillFormat.FillType = FillType.None;
            leadSourceColumnVND.PlotArea.Border.IsVisible = false;

            // Set value axis major tick mark as none and hide axis line. 
            // Also set the color of value axis major grid lines.
            leadSourceColumnVND.ValueAxis.MajorTickMark = TickMarkType.None;
            leadSourceColumnVND.ValueAxis.AxisLine.IsVisible = false;
            //leadSourceColumnChiNha.ValueAxis.IsAutomaticMajorUnit = false;
            //leadSourceColumnChiNha.ValueAxis.MajorUnit = 10000000;
            leadSourceColumnVND.ValueAxis.MajorGridLines.Color = Color.FromArgb(217, 217, 217);

            //Chart reference
            Aspose.Cells.Charts.Chart leadSourceColumnUSD;
            // USD
            chartIndex = sheetReport.Charts.Add(ChartType.Column3DClustered, 7, 7, 19, 13);
            leadSourceColumnUSD = sheetReport.Charts[chartIndex];

            //Chart title
            leadSourceColumnUSD.Title.Text = string.Format("Doanh số dịch vụ từng loại tiền - USD \n Giai đoạn: {0} \n", text);
            leadSourceColumnUSD.Title.Font.Color = Color.Silver;

            // Adding SeriesCollection (chart data source) to the chart ranging from "A1" cell to "B3"
            totalRowData = string.Format("D83:E{0}", 83 + table.Rows.Count - 2);
            leadSourceColumnUSD.NSeries.Add(totalRowData, true);

            // Set the category data covering the range A2:A5.
            // Tổng số dòng cần hiển thị là số dòng hiện tại  + số dòng trong listData -2 (trong đó 1 là dòng tổng cuối cùng)
            categoryData = "A83:A88";
            leadSourceColumnUSD.NSeries.CategoryData = categoryData;

            // Set the names of the chart series taken from cells.
            leadSourceColumnUSD.NSeries[0].Name = "=B82";
            leadSourceColumnUSD.NSeries[1].Name = "=C82";


            // Set the 1st series fill color.
            leadSourceColumnUSD.NSeries[0].Area.ForegroundColor = Color.Orange;
            leadSourceColumnUSD.NSeries[0].Area.Formatting = FormattingType.Custom;

            // Set the 2nd series fill color.
            leadSourceColumnUSD.NSeries[1].Area.ForegroundColor = Color.Green;
            leadSourceColumnUSD.NSeries[1].Area.Formatting = FormattingType.Custom;


            // Set plot area formatting as none and hide its border.
            leadSourceColumnUSD.PlotArea.Area.FillFormat.FillType = FillType.None;
            leadSourceColumnUSD.PlotArea.Border.IsVisible = false;

            // Set value axis major tick mark as none and hide axis line. 
            // Also set the color of value axis major grid lines.
            leadSourceColumnUSD.ValueAxis.MajorTickMark = TickMarkType.None;
            leadSourceColumnUSD.ValueAxis.AxisLine.IsVisible = false;
            leadSourceColumnUSD.ValueAxis.MajorGridLines.Color = Color.FromArgb(217, 217, 217);


            //Chart reference
            Aspose.Cells.Charts.Chart leadSourceColumnEUR;
            // EUR
            chartIndex = sheetReport.Charts.Add(ChartType.Column3DClustered, 20, 0, 33, 6);
            leadSourceColumnEUR = sheetReport.Charts[chartIndex];

            //Chart title
            leadSourceColumnEUR.Title.Text = string.Format("Doanh số dịch vụ từng loại tiền - EUR \n Giai đoạn: {0} \n", text);
            leadSourceColumnEUR.Title.Font.Color = Color.Silver;

            // Adding SeriesCollection (chart data source) to the chart ranging from "A1" cell to "B3"
            totalRowData = string.Format("F83:G{0}", 83 + table.Rows.Count - 2);
            leadSourceColumnEUR.NSeries.Add(totalRowData, true);

            // Set the category data covering the range A2:A5.
            // Tổng số dòng cần hiển thị là số dòng hiện tại  + số dòng trong listData -2 (trong đó 1 là dòng tổng cuối cùng)
            categoryData = "A83:A88";
            leadSourceColumnEUR.NSeries.CategoryData = categoryData;

            // Set the names of the chart series taken from cells.
            leadSourceColumnEUR.NSeries[0].Name = "=B82";
            leadSourceColumnEUR.NSeries[1].Name = "=C82";


            // Set the 1st series fill color.
            leadSourceColumnEUR.NSeries[0].Area.ForegroundColor = Color.Orange;
            leadSourceColumnEUR.NSeries[0].Area.Formatting = FormattingType.Custom;

            // Set the 2nd series fill color.
            leadSourceColumnEUR.NSeries[1].Area.ForegroundColor = Color.Green;
            leadSourceColumnEUR.NSeries[1].Area.Formatting = FormattingType.Custom;


            // Set plot area formatting as none and hide its border.
            leadSourceColumnEUR.PlotArea.Area.FillFormat.FillType = FillType.None;
            leadSourceColumnEUR.PlotArea.Border.IsVisible = false;

            // Set value axis major tick mark as none and hide axis line. 
            // Also set the color of value axis major grid lines.
            leadSourceColumnEUR.ValueAxis.MajorTickMark = TickMarkType.None;
            leadSourceColumnEUR.ValueAxis.AxisLine.IsVisible = false;
            leadSourceColumnEUR.ValueAxis.MajorGridLines.Color = Color.FromArgb(217, 217, 217);


            //Chart reference
            Aspose.Cells.Charts.Chart leadSourceColumnCAD;
            // CAD
            chartIndex = sheetReport.Charts.Add(ChartType.Column3DClustered, 20, 7, 33, 13);
            leadSourceColumnCAD = sheetReport.Charts[chartIndex];

            //Chart title
            leadSourceColumnCAD.Title.Text = string.Format("Doanh số dịch vụ từng loại tiền - CAD \n Giai đoạn: {0} \n", text);
            leadSourceColumnCAD.Title.Font.Color = Color.Silver;

            // Adding SeriesCollection (chart data source) to the chart ranging from "A1" cell to "B3"
            totalRowData = string.Format("H83:I{0}", 83 + table.Rows.Count - 2);
            leadSourceColumnCAD.NSeries.Add(totalRowData, true);

            // Set the category data covering the range A2:A5.
            // Tổng số dòng cần hiển thị là số dòng hiện tại  + số dòng trong listData -2 (trong đó 1 là dòng tổng cuối cùng)
            categoryData = "A83:A88";
            leadSourceColumnCAD.NSeries.CategoryData = categoryData;

            // Set the names of the chart series taken from cells.
            leadSourceColumnCAD.NSeries[0].Name = "=B82";
            leadSourceColumnCAD.NSeries[1].Name = "=C82";


            // Set the 1st series fill color.
            leadSourceColumnCAD.NSeries[0].Area.ForegroundColor = Color.Orange;
            leadSourceColumnCAD.NSeries[0].Area.Formatting = FormattingType.Custom;

            // Set the 2nd series fill color.
            leadSourceColumnCAD.NSeries[1].Area.ForegroundColor = Color.Green;
            leadSourceColumnCAD.NSeries[1].Area.Formatting = FormattingType.Custom;


            // Set plot area formatting as none and hide its border.
            leadSourceColumnCAD.PlotArea.Area.FillFormat.FillType = FillType.None;
            leadSourceColumnCAD.PlotArea.Border.IsVisible = false;

            // Set value axis major tick mark as none and hide axis line. 
            // Also set the color of value axis major grid lines.
            leadSourceColumnCAD.ValueAxis.MajorTickMark = TickMarkType.None;
            leadSourceColumnCAD.ValueAxis.AxisLine.IsVisible = false;
            leadSourceColumnCAD.ValueAxis.MajorGridLines.Color = Color.FromArgb(217, 217, 217);


            //Chart reference
            Aspose.Cells.Charts.Chart leadSourceColumnAUD;
            // AUD
            chartIndex = sheetReport.Charts.Add(ChartType.Column3DClustered, 34, 0, 47, 6);
            leadSourceColumnAUD = sheetReport.Charts[chartIndex];

            //Chart title
            leadSourceColumnAUD.Title.Text = string.Format("Doanh số dịch vụ từng loại tiền - AUD \n Giai đoạn: {0} \n", text);
            leadSourceColumnAUD.Title.Font.Color = Color.Silver;

            // Adding SeriesCollection (chart data source) to the chart ranging from "A1" cell to "B3"
            totalRowData = string.Format("J83:K{0}", 83 + table.Rows.Count - 2);
            leadSourceColumnAUD.NSeries.Add(totalRowData, true);

            // Set the category data covering the range A2:A5.
            // Tổng số dòng cần hiển thị là số dòng hiện tại  + số dòng trong listData -2 (trong đó 1 là dòng tổng cuối cùng)
            categoryData = "A83:A88";
            leadSourceColumnAUD.NSeries.CategoryData = categoryData;

            // Set the names of the chart series taken from cells.
            leadSourceColumnAUD.NSeries[0].Name = "=B82";
            leadSourceColumnAUD.NSeries[1].Name = "=C82";


            // Set the 1st series fill color.
            leadSourceColumnAUD.NSeries[0].Area.ForegroundColor = Color.Orange;
            leadSourceColumnAUD.NSeries[0].Area.Formatting = FormattingType.Custom;

            // Set the 2nd series fill color.
            leadSourceColumnAUD.NSeries[1].Area.ForegroundColor = Color.Green;
            leadSourceColumnAUD.NSeries[1].Area.Formatting = FormattingType.Custom;


            // Set plot area formatting as none and hide its border.
            leadSourceColumnAUD.PlotArea.Area.FillFormat.FillType = FillType.None;
            leadSourceColumnAUD.PlotArea.Border.IsVisible = false;

            // Set value axis major tick mark as none and hide axis line. 
            // Also set the color of value axis major grid lines.
            leadSourceColumnAUD.ValueAxis.MajorTickMark = TickMarkType.None;
            leadSourceColumnAUD.ValueAxis.AxisLine.IsVisible = false;
            leadSourceColumnAUD.ValueAxis.MajorGridLines.Color = Color.FromArgb(217, 217, 217);


            //Chart reference
            Aspose.Cells.Charts.Chart leadSourceColumnGBP;
            // AUD
            chartIndex = sheetReport.Charts.Add(ChartType.Column3DClustered, 34, 7, 47, 13);
            leadSourceColumnGBP = sheetReport.Charts[chartIndex];

            //Chart title
            leadSourceColumnGBP.Title.Text = string.Format("Doanh số dịch vụ từng loại tiền - GBP \n Giai đoạn: {0} \n", text);
            leadSourceColumnGBP.Title.Font.Color = Color.Silver;

            // Adding SeriesCollection (chart data source) to the chart ranging from "A1" cell to "B3"
            totalRowData = string.Format("L83:M{0}", 83 + table.Rows.Count - 2);
            leadSourceColumnGBP.NSeries.Add(totalRowData, true);

            // Set the category data covering the range A2:A5.
            // Tổng số dòng cần hiển thị là số dòng hiện tại  + số dòng trong listData -2 (trong đó 1 là dòng tổng cuối cùng)
            categoryData = "A83:A88";
            leadSourceColumnGBP.NSeries.CategoryData = categoryData;

            // Set the names of the chart series taken from cells.
            leadSourceColumnGBP.NSeries[0].Name = "=B82";
            leadSourceColumnGBP.NSeries[1].Name = "=C82";


            // Set the 1st series fill color.
            leadSourceColumnGBP.NSeries[0].Area.ForegroundColor = Color.Orange;
            leadSourceColumnGBP.NSeries[0].Area.Formatting = FormattingType.Custom;

            // Set the 2nd series fill color.
            leadSourceColumnGBP.NSeries[1].Area.ForegroundColor = Color.Green;
            leadSourceColumnGBP.NSeries[1].Area.Formatting = FormattingType.Custom;


            // Set plot area formatting as none and hide its border.
            leadSourceColumnGBP.PlotArea.Area.FillFormat.FillType = FillType.None;
            leadSourceColumnGBP.PlotArea.Border.IsVisible = false;
            

            // Set value axis major tick mark as none and hide axis line. 
            // Also set the color of value axis major grid lines.
            leadSourceColumnGBP.ValueAxis.MajorTickMark = TickMarkType.None;
            leadSourceColumnGBP.ValueAxis.AxisLine.IsVisible = false;
            leadSourceColumnGBP.ValueAxis.MajorGridLines.Color = Color.FromArgb(217, 217, 217);

            // Tạo chart cột tỉ trọng cho các thị trường
            //Add Pie Chart
            //Chart reference
            Aspose.Cells.Charts.Chart leadSourceColumn;
            chartIndex = sheetReport.Charts.Add(ChartType.Column3DStacked, 50, 0, 77, 13);
            leadSourceColumn = sheetReport.Charts[chartIndex];


            //Chart title
            leadSourceColumn.Title.Text = string.Format("Tỉ trọng từng thị trường \n Giai đoạn: {0} \n", text);
            leadSourceColumn.Title.Font.Color = Color.Silver;

            // count cho 3 loại: chi Quầy, Chi nhà, Chuyển khoản
            // list thị trường

            listDataGradation = new ReportBL().ReportDetailtMTGradationCompareForOnePercent(year, int.Parse(gradationID), reportTypeID, marketID);
            List<ReportDetailtForTotalMoneyType> listDataGradationClone = new List<ReportDetailtForTotalMoneyType>(listDataGradation);

            foreach (ReportDetailtForTotalMoneyType item in listDataGradationClone)
            {
                ReportDetailtForTotalMoneyType dataItemYear = listDataGradation.Find(x => x.PartnerCode == item.PartnerCode && x.Year == year.ToString());
                ReportDetailtForTotalMoneyType dataItemLastYear = listDataGradation.Find(x => x.PartnerCode == item.PartnerCode && x.Year == (year - 1).ToString());

                // Trường hợp năm trước có đối tác và năm nay không có
                if (dataItemLastYear != null && dataItemYear == null)
                {
                    dataItemYear = new ReportDetailtForTotalMoneyType();
                    dataItemYear.PartnerCode = dataItemLastYear.PartnerCode;
                    dataItemYear.PartnerName = dataItemLastYear.PartnerName;
                    dataItemYear.Year = year.ToString();

                    // Add item thiếu vào
                    listDataGradation.Add(dataItemYear);
                }

                // Trường hợp năm trước không có đối tác và năm nay có đối tác
                if (dataItemYear != null && dataItemLastYear == null)
                {
                    dataItemLastYear = new ReportDetailtForTotalMoneyType();
                    dataItemLastYear.PartnerCode = dataItemYear.PartnerCode;
                    dataItemLastYear.PartnerName = dataItemYear.PartnerName;
                    dataItemLastYear.Year = (year - 1).ToString();

                    // Add item thiếu vào
                    listDataGradation.Add(dataItemLastYear);
                }
            }

            List<string> listPartnerCurrent = new List<string>();
            foreach (ReportDetailtForTotalMoneyType item in listDataGradation)
            {
                if (!listPartnerCurrent.Contains(item.PartnerName))
                {
                    listPartnerCurrent.Add(item.PartnerName);
                }
            }

            // List dữ liệu dataRow
            string[] listTotalRowData = new string[listPartnerCurrent.Count];
            int i = 0;
            foreach (string item in listPartnerCurrent)
            {
                // Năm trước
                List<ReportDetailtForTotalMoneyType> dataItemLastYear = listDataGradation.Where(x => x.PartnerName == item && x.Year == (year - 1).ToString()).ToList();
                List<ReportDetailtForTotalMoneyType> dataItemYear = listDataGradation.Where(x => x.PartnerName == item && x.Year == year.ToString()).ToList();


                listTotalRowData[i++] = string.Concat("{"
                    , string.Format("{0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9}, {10}, {11}"
                    , dataItemYear.Sum(x => x.VND), dataItemLastYear.Sum(x => x.VND)
                    , dataItemYear.Sum(x => x.USD), dataItemLastYear.Sum(x => x.USD)
                    , dataItemYear.Sum(x => x.EUR), dataItemLastYear.Sum(x => x.EUR)
                    , dataItemYear.Sum(x => x.CAD), dataItemLastYear.Sum(x => x.CAD)
                    , dataItemYear.Sum(x => x.AUD), dataItemLastYear.Sum(x => x.AUD)
                    , dataItemYear.Sum(x => x.GBP), dataItemLastYear.Sum(x => x.GBP))
                    , "}");
            }

            int count = 0;
            string yearCurent = string.Format("Năm {0}", year);
            string lastYear = string.Format("Năm {0}", year - 1);
            foreach (string item in listTotalRowData)
            {
                totalRowData = item;
                leadSourceColumn.NSeries.Add(totalRowData, true);

                categoryData = string.Concat("{", string.Format("VND {0}, VND {1}, USD {0}, USD {1}, EUR {0},  EUR {1}, CAD {0}, CAD {1}, AUD {0}, AUD {1}, GBP {0}, GBP {1}", yearCurent, lastYear), "}");
                leadSourceColumn.NSeries.CategoryData = categoryData;

                leadSourceColumn.NSeries[count].Name = listPartnerCurrent[count];

                //switch (count)
                //{
                //    case 0:
                //        leadSourceColumn.NSeries[count].Area.ForegroundColor = Color.Orange;
                //        leadSourceColumn.NSeries[count].Area.Formatting = FormattingType.Custom;
                //        break;
                //    case 1:
                //        leadSourceColumn.NSeries[count].Area.ForegroundColor = Color.Green;
                //        leadSourceColumn.NSeries[count].Area.Formatting = FormattingType.Custom;
                //        break;
                //    case 2:
                //        leadSourceColumn.NSeries[count].Area.ForegroundColor = Color.Blue;
                //        leadSourceColumn.NSeries[count].Area.Formatting = FormattingType.Custom;
                //        break;
                //    case 3:
                //        leadSourceColumn.NSeries[count].Area.ForegroundColor = Color.Yellow;
                //        leadSourceColumn.NSeries[count].Area.Formatting = FormattingType.Custom;
                //        break;
                //    case 4:
                //        leadSourceColumn.NSeries[count].Area.ForegroundColor = Color.Brown;
                //        leadSourceColumn.NSeries[count].Area.Formatting = FormattingType.Custom;
                //        break;
                //    case 5:
                //        leadSourceColumn.NSeries[count].Area.ForegroundColor = Color.Olive;
                //        leadSourceColumn.NSeries[count].Area.Formatting = FormattingType.Custom;
                //        break;
                //    default:
                //        leadSourceColumn.NSeries[count].Area.ForegroundColor = Color.Pink;
                //        leadSourceColumn.NSeries[count].Area.Formatting = FormattingType.Custom;
                //        break;
                //}
                count++;
            }

            // Set plot area formatting as none and hide its border.
            leadSourceColumn.PlotArea.Area.FillFormat.FillType = FillType.None;
            leadSourceColumn.PlotArea.Border.IsVisible = false;

            // Set value axis major tick mark as none and hide axis line. 
            // Also set the color of value axis major grid lines.
            leadSourceColumn.ValueAxis.MajorTickMark = TickMarkType.None;
            leadSourceColumn.ValueAxis.MaxValue = 100;
            leadSourceColumn.ValueAxis.AxisLine.IsVisible = false;
            leadSourceColumn.ValueAxis.MajorGridLines.Color = Color.FromArgb(217, 217, 217);

            // Chạy process
            designer.Process();
            return ExportReport("ReportGradationCompare", designer);
        }

        /// <summary>
        /// Hàm tạo tile cho xuất Excell
        /// </summary>
        /// <param name="upperLeft"></param>
        /// <param name="upperRight"></param>
        /// <param name="sheetReport"></param>
        /// <param name="Title"></param>
        private void CreateTitle(string upperLeft, string upperRight, Worksheet sheetReport, string Title, int size, bool checkStyle = false)
        {
            //Add range title
            Range rangeTitle = sheetReport.Cells.CreateRange(upperLeft, upperRight);
            //Merage title
            rangeTitle.UnMerge();
            rangeTitle.Merge();
            //Put text to title
            rangeTitle.PutValue(Title, false, true);
            //Set style range title
            Style styleTitle = new CellsFactory().CreateStyle();
            styleTitle.Font.IsBold = true;
            styleTitle.Font.Color = Color.Black;
            styleTitle.Font.Size = size;
            if (checkStyle)
            {
                styleTitle.SetBorder(BorderType.BottomBorder, CellBorderType.Thin, Color.Black);
                styleTitle.SetBorder(BorderType.LeftBorder, CellBorderType.Thin, Color.Black);
                styleTitle.SetBorder(BorderType.RightBorder, CellBorderType.Thin, Color.Black);
                styleTitle.SetBorder(BorderType.TopBorder, CellBorderType.Thin, Color.Black);
                styleTitle.VerticalAlignment = TextAlignmentType.Center;
            }

            //styleTitle.Font.Name = "Times New Roman";
            styleTitle.Font.Name = "Calibri";
            styleTitle.HorizontalAlignment = TextAlignmentType.Center;
            rangeTitle.SetStyle(styleTitle);
        }

        /// <summary>
        /// Tạo cột cho datatable
        /// </summary>
        /// <returns></returns>
        private DataTable CreateDataTableFormart(bool typeID = false)
        {
            DataTable db = new DataTable();
            
            db.Columns.Add("ReportID", typeof(string));
            db.Columns.Add("VND", typeof(double));
            db.Columns.Add("USD", typeof(double));
            db.Columns.Add("EUR", typeof(double));
            db.Columns.Add("CAD", typeof(double));
            db.Columns.Add("AUD", typeof(double));
            db.Columns.Add("GBP", typeof(double));
            if (typeID)
            {
                db.Columns.Add("TongDS", typeof(double));
            }
            return db;
        }

        /// <summary>
        /// convert từ list object to dataset
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="data"></param>
        /// <returns></returns>
        public DataSet ConvertListObjectToDataSet<T>(IList<T> data)
        {
            DataTable dataTable = ConvertToDataTable(data);

            DataSet ds = new DataSet();
            if (dataTable.Rows.Count > 0)
            {
                ds.Tables.Add(dataTable);
            }
            return ds;
        }

        /// <summary>
        /// Add row cho datatable
        /// </summary>
        /// <param name="mother"></param>
        /// <param name="fill"></param>
        private void FillData(DataTable mother, DataTable fill, bool typeID = false)
        {
            int stt = 1;
            foreach (DataRow dr in mother.Rows)
            {
                var row = fill.NewRow();
                
                row["ReportID"] = string.IsNullOrEmpty(dr["ReportID"].ToString()) ? "" : (string)dr["ReportID"];
                row["VND"] = (double)dr["VND"];
                row["USD"] = (double)dr["USD"];
                row["EUR"] = (double)dr["EUR"];
                row["CAD"] = (double)dr["CAD"];
                row["AUD"] = (double)dr["AUD"];
                row["GBP"] = (double)dr["GBP"];
                if (typeID)
                {
                    row["TongDS"] = (double)dr["TongDS"];
                }
                
                fill.Rows.Add(row);
                stt++;
            }
        }

        /// <summary>
        /// Convert to listObject to datatable
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="data"></param>
        /// <returns></returns>
        public DataTable ConvertToDataTable<T>(IList<T> data)
        {
            PropertyDescriptorCollection properties = TypeDescriptor.GetProperties(typeof(T));
            DataTable table = new DataTable();
            foreach (PropertyDescriptor prop in properties)
                table.Columns.Add(prop.Name, Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType);
            foreach (T item in data)
            {
                DataRow row = table.NewRow();
                foreach (PropertyDescriptor prop in properties)
                    row[prop.Name] = prop.GetValue(item) ?? DBNull.Value;
                table.Rows.Add(row);
            }
            return table;
        }

        public FileStreamResult ExportReport(string ReportID, WorkbookDesigner designer)
        {
            try
            {
                string LData = WebConfigurationManager.AppSettings["LData"];
                Stream stream = new MemoryStream(Convert.FromBase64String(LData));

                stream.Seek(0, SeekOrigin.Begin);
                new Aspose.Cells.License().SetLicense(stream);

                if (designer != null)
                {
                    stream = new DongA.Core.DongAExcel().SaveToStream(designer.Workbook);
                }

                // Return excel
                return File(stream, XLSX,
                    string.Format("{0}.{1}", ReportID, "xlsx"));
            }
            catch (Exception ex)
            {
                MemoryStream memStream = new MemoryStream();
                return File(memStream, PDF);
            }
        }
    }
}
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
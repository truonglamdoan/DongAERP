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
    public class ReportExcelDetailtByMarketController : Controller
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

        // GET: Admin/ReportExcelDetailtByMarket
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
            string templatePath = "~/Content/Report/ReportDetailtByMakets.xlsx";
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
            List<ReportDetailtSTMarket> listReportData = new List<ReportDetailtSTMarket>();

            switch (typeID)
            {
                // Theo ngày
                case 1:
                    listReportData = new ReportBL().SearchMarketForTotalForDay(fromDate, toDate, reportTypeID, marketID);

                    if (!string.IsNullOrEmpty(marketID))
                    {
                        int count = 1;
                        if (marketID == "0")
                        {
                            foreach (ReportDetailtSTMarket item in listReportData)
                            {
                                item.ReportID = item.MarketName;
                                item.TongDS = item.DSChiNha + item.DSChiQuay + item.DSCK;
                            }
                        }
                        else
                        {
                            List<string> listAsianChild = new List<string>();
                            List<ReportDetailtSTMarket> listDataConvert = new List<ReportDetailtSTMarket>();

                            foreach (ReportDetailtSTMarket item in listReportData)
                            {
                                item.STT = (count++).ToString();
                                item.ReportID = item.PartnerName;
                                item.TongDS = item.DSChiNha + item.DSChiQuay + item.DSCK;

                                if (!listAsianChild.Contains(item.MarketName))
                                {
                                    listAsianChild.Add(item.MarketName);
                                }
                            }

                            foreach (string item in listAsianChild)
                            {
                                List<ReportDetailtSTMarket> listDataAsianChild = listReportData.Where(x => x.MarketName == item).ToList();

                                listDataConvert.Add(
                                    new ReportDetailtSTMarket()
                                    {
                                        STT = (count++).ToString(),
                                        MarketName = "Châu Á",
                                        ReportID = item,
                                        DSChiQuay = listDataAsianChild.Sum(x => x.DSChiQuay),
                                        DSChiNha = listDataAsianChild.Sum(x => x.DSChiNha),
                                        DSCK = listDataAsianChild.Sum(x => x.DSCK),
                                        TongDS = listDataAsianChild.Sum(x => x.TongDS),

                                    }
                                );
                            }

                            if (listDataConvert.Count > 0)
                            {
                                listReportData = new List<ReportDetailtSTMarket>(listDataConvert);
                            }
                        }
                    }

                    // Set from day and to day
                    sheetReport.Cells["E4"].PutValue(fromDate.ToString("dd/MM/yyyy"));
                    sheetReport.Cells["H4"].PutValue(toDate.ToString("dd/MM/yyyy"));
                    break;
                // Theo tháng
                case 2:

                    listReportData = new ReportBL().SearchMarketForTotalForMonth(fromDate, toDate, reportTypeID, marketID);

                    if (!string.IsNullOrEmpty(marketID))
                    {
                        int count = 1;
                        if (marketID == "0")
                        {
                            foreach (ReportDetailtSTMarket item in listReportData)
                            {
                                item.STT = (count++).ToString();
                                item.ReportID = item.MarketName;
                                item.TongDS = item.DSChiNha + item.DSChiQuay + item.DSCK;
                            }
                        }
                        else
                        {
                            List<string> listAsianChild = new List<string>();
                            List<ReportDetailtSTMarket> listDataConvert = new List<ReportDetailtSTMarket>();

                            foreach (ReportDetailtSTMarket item in listReportData)
                            {
                                item.ReportID = item.PartnerName;
                                item.TongDS = item.DSChiNha + item.DSChiQuay + item.DSCK;

                                if (!listAsianChild.Contains(item.MarketName))
                                {
                                    listAsianChild.Add(item.MarketName);
                                }
                            }

                            foreach (string item in listAsianChild)
                            {
                                List<ReportDetailtSTMarket> listDataAsianChild = listReportData.Where(x => x.MarketName == item).ToList();

                                listDataConvert.Add(
                                    new ReportDetailtSTMarket()
                                    {
                                        STT = (count++).ToString(),
                                        MarketName = "Châu Á",
                                        ReportID = item,
                                        DSChiQuay = listDataAsianChild.Sum(x => x.DSChiQuay),
                                        DSChiNha = listDataAsianChild.Sum(x => x.DSChiNha),
                                        DSCK = listDataAsianChild.Sum(x => x.DSCK),
                                        TongDS = listDataAsianChild.Sum(x => x.TongDS),

                                    }
                                );
                            }

                            if (listDataConvert.Count > 0)
                            {
                                listReportData = new List<ReportDetailtSTMarket>(listDataConvert);
                            }
                        }
                    }
                    // Set from day and to day
                    sheetReport.Cells["E4"].PutValue(string.Format("{0}/{1}", fromDate.Month, fromDate.Year));
                    sheetReport.Cells["H4"].PutValue(string.Format("{0}/{1}", toDate.Month, toDate.Year));
                    break;
                // Theo năm
                default:

                    listReportData = new ReportBL().SearchMarketForTotalForYear(fromDate, toDate, reportTypeID, marketID);

                    if (!string.IsNullOrEmpty(marketID))
                    {
                        int count = 1;
                        if (marketID == "0")
                        {
                            foreach (ReportDetailtSTMarket item in listReportData)
                            {
                                item.STT = (count++).ToString();
                                item.ReportID = item.MarketName;
                                item.TongDS = item.DSChiNha + item.DSChiQuay + item.DSCK;
                            }
                        }
                        else
                        {
                            List<string> listAsianChild = new List<string>();
                            List<ReportDetailtSTMarket> listDataConvert = new List<ReportDetailtSTMarket>();

                            foreach (ReportDetailtSTMarket item in listReportData)
                            {
                                item.ReportID = item.PartnerName;
                                item.TongDS = item.DSChiNha + item.DSChiQuay + item.DSCK;

                                if (!listAsianChild.Contains(item.MarketName))
                                {
                                    listAsianChild.Add(item.MarketName);
                                }
                            }

                            foreach (string item in listAsianChild)
                            {
                                List<ReportDetailtSTMarket> listDataAsianChild = listReportData.Where(x => x.MarketName == item).ToList();

                                listDataConvert.Add(
                                    new ReportDetailtSTMarket()
                                    {
                                        STT = (count++).ToString(),
                                        MarketName = "Châu Á",
                                        ReportID = item,
                                        DSChiQuay = listDataAsianChild.Sum(x => x.DSChiQuay),
                                        DSChiNha = listDataAsianChild.Sum(x => x.DSChiNha),
                                        DSCK = listDataAsianChild.Sum(x => x.DSCK),
                                        TongDS = listDataAsianChild.Sum(x => x.TongDS),

                                    }
                                );
                            }

                            if (listDataConvert.Count > 0)
                            {
                                listReportData = new List<ReportDetailtSTMarket>(listDataConvert);
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
                ReportDetailtSTMarket reportForMaket = new ReportDetailtSTMarket()
                {
                    STT = "",
                    ReportID = "Tổng",
                    DSChiQuay = listReportData.Sum(item => item.DSChiQuay),
                    DSChiNha = listReportData.Sum(item => item.DSChiNha),
                    DSCK = listReportData.Sum(item => item.DSCK),
                    TongDS = listReportData.Sum(item => item.TongDS)
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
            style.SetBorder(BorderType.BottomBorder, CellBorderType.Thin, Color.Black);
            style.SetBorder(BorderType.LeftBorder, CellBorderType.Thin, Color.Black);
            style.SetBorder(BorderType.RightBorder, CellBorderType.Thin, Color.Black);

            if (dataTable.Rows.Count > 0)
            {
                int stepRow = 0;
                // total row = row start + số row hiện có
                int totalRow = dataTable.Rows.Count + 7;
                // Số dòng của row
                for (int a = 7; a < totalRow; a++)
                {
                    int stepColumn = 0;
                    // Số cột trong báo cáo cần hiển thị
                    // Tổng số cột hiển thị = Số cột hiển thị bắt đầu + tổng số cột cần hiển thị
                    // Số 6 là cột marketName
                    int totalCol = 1 + 6;
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

                //// Group 
                //// Create a cellarea i.e.., B3:C19
                //CellArea ca = new CellArea();
                //ca.StartRow = 7;
                //ca.StartColumn = 2;
                //ca.EndRow = dataTable.Rows.Count + 7;
                //ca.EndColumn = 7;

                //// Second column (C) in the list
                //// Group by column index
                //// số cột được tí từ start column
                //sheetReport.Cells.Subtotal(ca, 5, ConsolidationFunction.Sum, new int[] { 1,2,3,4 });
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
            string templatePath = "~/Content/Report/ReportDetailtByMakets.xlsx";
            // Get đường dẫn
            templatePath = System.Web.HttpContext.Current.Server.MapPath(templatePath);

            designer.Workbook = new DongA.Core.DongAExcel().OpenExcelFile(templatePath);
            designer.Workbook.CalculateFormula();

            // Set the glorbalization setting to change subtotal and grand total names
            GlobalizationSettings gsi = new GlobalizationSettingsImp();
            designer.Workbook.Settings.GlobalizationSettings = gsi;

            WorksheetCollection workSheets = designer.Workbook.Worksheets;
            Worksheet sheetReport = designer.Workbook.Worksheets[0];

            // Thay đổi title cột thị trường thành đối tác
            sheetReport.Cells["C7"].PutValue("Đối tác");
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
            List<ReportDetailtServiceType> listReportData = new List<ReportDetailtServiceType>();
            // Get danh sách thị trường
            List<string> listMarket = new List<string>();

            switch (typeID)
            {
                // Theo ngày
                case 1:
                    listReportData = new ReportBL().SearchMarketForOne(fromDate, toDate, reportTypeID, marketID);
                    // Tính tổng doanh số
                    foreach (ReportDetailtServiceType item in listReportData)
                    {
                        item.TongDS = item.DSChiNha + item.DSChiQuay + item.DSCK;
                    }

                    if (!string.IsNullOrEmpty(marketID))
                    {
                        int count = 1;
                        if (marketID != "005")
                        {
                            foreach (ReportDetailtServiceType item in listReportData)
                            {
                                // Add thêm thị trường vào list
                                if (!listMarket.Contains(item.MarketName))
                                {
                                    listMarket.Add(item.MarketName);
                                }
                                item.STT = (count++).ToString();
                                item.ReportID = item.PartnerName;
                                item.TongDS = item.DSChiNha + item.DSChiQuay + item.DSCK;
                            }

                            listReportData.Add(
                                new ReportDetailtServiceType()
                                {
                                    STT = "",
                                    ReportID = "Tổng",
                                    DSChiQuay = listReportData.Sum(x => x.DSChiQuay),
                                    DSChiNha = listReportData.Sum(x => x.DSChiNha),
                                    DSCK = listReportData.Sum(x => x.DSCK),
                                    TongDS = listReportData.Sum(x => x.TongDS),
                                }
                            );
                        }
                        else
                        {
                            List<ReportDetailtServiceType> listDataConvert = new List<ReportDetailtServiceType>();
                            string marketName = string.Empty;
                            foreach (ReportDetailtServiceType item in listReportData)
                            {
                                // Add thêm thị trường vào list
                                if (!listMarket.Contains(item.MarketName))
                                {
                                    // Dòng đầu tiên
                                    // Và chỉ có 1 đối tác của 1 thị trường
                                    if (count == 1)
                                    {
                                        listDataConvert.Add(
                                            new ReportDetailtServiceType()
                                            {
                                                STT = (count++).ToString(),
                                                MarketCode = item.MarketCode,
                                                MarketName = item.MarketName,
                                                PartnerCode = item.PartnerCode,
                                                PartnerName = item.PartnerName,
                                                ReportID = item.PartnerName,
                                                DSChiQuay = item.DSChiQuay,
                                                DSChiNha = item.DSChiNha,
                                                DSCK = item.DSCK,
                                                TongDS = item.DSChiNha + item.DSChiQuay + item.DSCK
                                            }
                                        );
                                    }
                                    else
                                    {
                                        // Trường hợp thuộc dòng cuối mỗi thị trường
                                        // Add dòng tổng cho các đôi tác thị trường trước
                                        listDataConvert.Add(
                                            new ReportDetailtServiceType()
                                            {
                                                STT = "",
                                                PartnerCode = "",
                                                PartnerName = "Tổng",
                                                ReportID = "Tổng",
                                                MarketName = marketName,
                                                DSChiQuay = listReportData.Where(x => x.MarketName == marketName).Sum(x => x.DSChiQuay),
                                                DSChiNha = listReportData.Where(x => x.MarketName == marketName).Sum(x => x.DSChiNha),
                                                DSCK = listReportData.Where(x => x.MarketName == marketName).Sum(x => x.DSCK),
                                                TongDS = listReportData.Where(x => x.MarketName == marketName).Sum(x => x.TongDS),
                                            }
                                        );
                                        // Set stt lại bằng 1
                                        count = 1;

                                        // Số dòng của Thị trường
                                        int countList = listReportData.Where(x => x.MarketCode == item.MarketCode).Count();
                                        if (countList == 1)
                                        {
                                            listDataConvert.Add(
                                                new ReportDetailtServiceType()
                                                {
                                                    STT = (count++).ToString(),
                                                    MarketCode = item.MarketCode,
                                                    MarketName = item.MarketName,
                                                    PartnerCode = item.PartnerCode,
                                                    PartnerName = item.PartnerName,
                                                    ReportID = item.PartnerName,
                                                    DSChiQuay = item.DSChiQuay,
                                                    DSChiNha = item.DSChiNha,
                                                    DSCK = item.DSCK,
                                                    TongDS = item.DSChiNha + item.DSChiQuay + item.DSCK
                                                }
                                            );

                                            // Add dòng tổng vào
                                            listDataConvert.Add(
                                                new ReportDetailtServiceType()
                                                {
                                                    STT = "",
                                                    PartnerCode = "",
                                                    PartnerName = "Tổng",
                                                    ReportID = "Tổng",
                                                    MarketName = item.MarketName,
                                                    DSChiQuay = listReportData.Where(x => x.MarketName == item.MarketName).Sum(x => x.DSChiQuay),
                                                    DSChiNha = listReportData.Where(x => x.MarketName == item.MarketName).Sum(x => x.DSChiNha),
                                                    DSCK = listReportData.Where(x => x.MarketName == item.MarketName).Sum(x => x.DSCK),
                                                    TongDS = listReportData.Where(x => x.MarketName == item.MarketName).Sum(x => x.TongDS),
                                                }
                                            );
                                        }
                                        else
                                        {
                                            listDataConvert.Add(
                                                new ReportDetailtServiceType()
                                                {
                                                    STT = (count++).ToString(),
                                                    MarketCode = item.MarketCode,
                                                    MarketName = item.MarketName,
                                                    PartnerCode = item.PartnerCode,
                                                    PartnerName = item.PartnerName,
                                                    ReportID = item.PartnerName,
                                                    DSChiQuay = item.DSChiQuay,
                                                    DSChiNha = item.DSChiNha,
                                                    DSCK = item.DSCK,
                                                    TongDS = item.DSChiNha + item.DSChiQuay + item.DSCK
                                                }
                                            );
                                        }
                                    }
                                    listMarket.Add(item.MarketName);
                                    marketName = item.MarketName;

                                }
                                else
                                {
                                    ReportDetailtServiceType dataItemLast = listReportData[listReportData.Count - 1];
                                    // Trường hợp khi dòng cuối cần tính tổng
                                    if (dataItemLast.PartnerCode == item.PartnerCode)
                                    {
                                        // Add item cuối vào
                                        listDataConvert.Add(
                                            new ReportDetailtServiceType()
                                            {
                                                STT = (count++).ToString(),
                                                MarketCode = item.MarketCode,
                                                MarketName = item.MarketName,
                                                PartnerCode = item.PartnerCode,
                                                PartnerName = item.PartnerName,
                                                ReportID = item.PartnerName,
                                                DSChiQuay = item.DSChiQuay,
                                                DSChiNha = item.DSChiNha,
                                                DSCK = item.DSCK,
                                                TongDS = item.DSChiNha + item.DSChiQuay + item.DSCK
                                            }
                                        );

                                        // Add dòng tổng vào
                                        listDataConvert.Add(
                                            new ReportDetailtServiceType()
                                            {
                                                STT = "",
                                                PartnerCode = "",
                                                PartnerName = "Tổng",
                                                ReportID = "Tổng",
                                                MarketName = item.MarketName,
                                                DSChiQuay = listReportData.Where(x => x.MarketName == item.MarketName).Sum(x => x.DSChiQuay),
                                                DSChiNha = listReportData.Where(x => x.MarketName == item.MarketName).Sum(x => x.DSChiNha),
                                                DSCK = listReportData.Where(x => x.MarketName == item.MarketName).Sum(x => x.DSCK),
                                                TongDS = listReportData.Where(x => x.MarketName == item.MarketName).Sum(x => x.TongDS),
                                            }
                                        );
                                    }
                                    else
                                    {
                                        // Chưa phải dòng cuối chỉ add item vào
                                        listDataConvert.Add(
                                            new ReportDetailtServiceType()
                                            {
                                                STT = (count++).ToString(),
                                                MarketCode = item.MarketCode,
                                                MarketName = item.MarketName,
                                                PartnerCode = item.PartnerCode,
                                                PartnerName = item.PartnerName,
                                                ReportID = item.PartnerName,
                                                DSChiQuay = item.DSChiQuay,
                                                DSChiNha = item.DSChiNha,
                                                DSCK = item.DSCK,
                                                TongDS = item.DSChiNha + item.DSChiQuay + item.DSCK
                                            }
                                        );
                                    }
                                }
                            }
                            listReportData = new List<ReportDetailtServiceType>(listDataConvert);
                        }
                    }

                    // Set from day and to day
                    sheetReport.Cells["E4"].PutValue(fromDate.ToString("dd/MM/yyyy"));
                    sheetReport.Cells["H4"].PutValue(toDate.ToString("dd/MM/yyyy"));
                    break;
                // Theo tháng
                case 2:

                    listReportData = new ReportBL().SearchMarketForOneForMonth(fromDate, toDate, reportTypeID, marketID);

                    // Tính tổng doanh số
                    foreach (ReportDetailtServiceType item in listReportData)
                    {
                        item.TongDS = item.DSChiNha + item.DSChiQuay + item.DSCK;
                    }

                    if (!string.IsNullOrEmpty(marketID))
                    {
                        int count = 1;
                        if (marketID != "005")
                        {
                            foreach (ReportDetailtServiceType item in listReportData)
                            {
                                // Add thêm thị trường vào list
                                if (!listMarket.Contains(item.MarketName))
                                {
                                    listMarket.Add(item.MarketName);
                                }
                                item.STT = (count++).ToString();
                                item.ReportID = item.PartnerName;
                                item.TongDS = item.DSChiNha + item.DSChiQuay + item.DSCK;
                            }

                            listReportData.Add(
                                new ReportDetailtServiceType()
                                {
                                    STT = "",
                                    ReportID = "Tổng",
                                    DSChiQuay = listReportData.Sum(x => x.DSChiQuay),
                                    DSChiNha = listReportData.Sum(x => x.DSChiNha),
                                    DSCK = listReportData.Sum(x => x.DSCK),
                                    TongDS = listReportData.Sum(x => x.TongDS),
                                }
                            );
                        }
                        else
                        {
                            List<ReportDetailtServiceType> listDataConvert = new List<ReportDetailtServiceType>();
                            string marketName = string.Empty;
                            foreach (ReportDetailtServiceType item in listReportData)
                            {
                                // Add thêm thị trường vào list
                                if (!listMarket.Contains(item.MarketName))
                                {
                                    // Dòng đầu tiên
                                    // Và chỉ có 1 đối tác của 1 thị trường
                                    if (count == 1)
                                    {
                                        listDataConvert.Add(
                                            new ReportDetailtServiceType()
                                            {
                                                STT = (count++).ToString(),
                                                MarketCode = item.MarketCode,
                                                MarketName = item.MarketName,
                                                PartnerCode = item.PartnerCode,
                                                PartnerName = item.PartnerName,
                                                ReportID = item.PartnerName,
                                                DSChiQuay = item.DSChiQuay,
                                                DSChiNha = item.DSChiNha,
                                                DSCK = item.DSCK,
                                                TongDS = item.DSChiNha + item.DSChiQuay + item.DSCK
                                            }
                                        );
                                    }
                                    else
                                    {
                                        // Trường hợp thuộc dòng cuối mỗi thị trường
                                        // Add dòng tổng cho các đôi tác thị trường trước
                                        listDataConvert.Add(
                                            new ReportDetailtServiceType()
                                            {
                                                STT = "",
                                                PartnerCode = "",
                                                PartnerName = "Tổng",
                                                ReportID = "Tổng",
                                                MarketName = marketName,
                                                DSChiQuay = listReportData.Where(x => x.MarketName == marketName).Sum(x => x.DSChiQuay),
                                                DSChiNha = listReportData.Where(x => x.MarketName == marketName).Sum(x => x.DSChiNha),
                                                DSCK = listReportData.Where(x => x.MarketName == marketName).Sum(x => x.DSCK),
                                                TongDS = listReportData.Where(x => x.MarketName == marketName).Sum(x => x.TongDS),
                                            }
                                        );
                                        // Set stt lại bằng 1
                                        count = 1;

                                        // Số dòng của Thị trường
                                        int countList = listReportData.Where(x => x.MarketCode == item.MarketCode).Count();
                                        if (countList == 1)
                                        {
                                            listDataConvert.Add(
                                                new ReportDetailtServiceType()
                                                {
                                                    STT = (count++).ToString(),
                                                    MarketCode = item.MarketCode,
                                                    MarketName = item.MarketName,
                                                    PartnerCode = item.PartnerCode,
                                                    PartnerName = item.PartnerName,
                                                    ReportID = item.PartnerName,
                                                    DSChiQuay = item.DSChiQuay,
                                                    DSChiNha = item.DSChiNha,
                                                    DSCK = item.DSCK,
                                                    TongDS = item.DSChiNha + item.DSChiQuay + item.DSCK
                                                }
                                            );

                                            // Add dòng tổng vào
                                            listDataConvert.Add(
                                                new ReportDetailtServiceType()
                                                {
                                                    STT = "",
                                                    PartnerCode = "",
                                                    PartnerName = "Tổng",
                                                    ReportID = "Tổng",
                                                    MarketName = item.MarketName,
                                                    DSChiQuay = listReportData.Where(x => x.MarketName == item.MarketName).Sum(x => x.DSChiQuay),
                                                    DSChiNha = listReportData.Where(x => x.MarketName == item.MarketName).Sum(x => x.DSChiNha),
                                                    DSCK = listReportData.Where(x => x.MarketName == item.MarketName).Sum(x => x.DSCK),
                                                    TongDS = listReportData.Where(x => x.MarketName == item.MarketName).Sum(x => x.TongDS),
                                                }
                                            );
                                        }
                                        else
                                        {
                                            listDataConvert.Add(
                                                new ReportDetailtServiceType()
                                                {
                                                    STT = (count++).ToString(),
                                                    MarketCode = item.MarketCode,
                                                    MarketName = item.MarketName,
                                                    PartnerCode = item.PartnerCode,
                                                    PartnerName = item.PartnerName,
                                                    ReportID = item.PartnerName,
                                                    DSChiQuay = item.DSChiQuay,
                                                    DSChiNha = item.DSChiNha,
                                                    DSCK = item.DSCK,
                                                    TongDS = item.DSChiNha + item.DSChiQuay + item.DSCK
                                                }
                                            );
                                        }
                                    }
                                    listMarket.Add(item.MarketName);
                                    marketName = item.MarketName;

                                }
                                else
                                {
                                    ReportDetailtServiceType dataItemLast = listReportData[listReportData.Count - 1];
                                    // Trường hợp khi dòng cuối cần tính tổng
                                    if (dataItemLast.PartnerCode == item.PartnerCode)
                                    {
                                        // Add item cuối vào
                                        listDataConvert.Add(
                                            new ReportDetailtServiceType()
                                            {
                                                STT = (count++).ToString(),
                                                MarketCode = item.MarketCode,
                                                MarketName = item.MarketName,
                                                PartnerCode = item.PartnerCode,
                                                PartnerName = item.PartnerName,
                                                ReportID = item.PartnerName,
                                                DSChiQuay = item.DSChiQuay,
                                                DSChiNha = item.DSChiNha,
                                                DSCK = item.DSCK,
                                                TongDS = item.DSChiNha + item.DSChiQuay + item.DSCK
                                            }
                                        );

                                        // Add dòng tổng vào
                                        listDataConvert.Add(
                                            new ReportDetailtServiceType()
                                            {
                                                STT = "",
                                                PartnerCode = "",
                                                PartnerName = "Tổng",
                                                ReportID = "Tổng",
                                                MarketName = item.MarketName,
                                                DSChiQuay = listReportData.Where(x => x.MarketName == item.MarketName).Sum(x => x.DSChiQuay),
                                                DSChiNha = listReportData.Where(x => x.MarketName == item.MarketName).Sum(x => x.DSChiNha),
                                                DSCK = listReportData.Where(x => x.MarketName == item.MarketName).Sum(x => x.DSCK),
                                                TongDS = listReportData.Where(x => x.MarketName == item.MarketName).Sum(x => x.TongDS),
                                            }
                                        );
                                    }
                                    else
                                    {
                                        // Chưa phải dòng cuối chỉ add item vào
                                        listDataConvert.Add(
                                            new ReportDetailtServiceType()
                                            {
                                                STT = (count++).ToString(),
                                                MarketCode = item.MarketCode,
                                                MarketName = item.MarketName,
                                                PartnerCode = item.PartnerCode,
                                                PartnerName = item.PartnerName,
                                                ReportID = item.PartnerName,
                                                DSChiQuay = item.DSChiQuay,
                                                DSChiNha = item.DSChiNha,
                                                DSCK = item.DSCK,
                                                TongDS = item.DSChiNha + item.DSChiQuay + item.DSCK
                                            }
                                        );
                                    }
                                }
                            }
                            listReportData = new List<ReportDetailtServiceType>(listDataConvert);
                        }
                    }
                    // Set from day and to day
                    sheetReport.Cells["E4"].PutValue(string.Format("{0}/{1}", fromDate.Month, fromDate.Year));
                    sheetReport.Cells["H4"].PutValue(string.Format("{0}/{1}", toDate.Month, toDate.Year));
                    break;
                // Theo năm
                default:

                    listReportData = new ReportBL().SearchMarketForOneForYear(fromDate, toDate, reportTypeID, marketID);

                    // Tính tổng doanh số
                    foreach (ReportDetailtServiceType item in listReportData)
                    {
                        item.TongDS = item.DSChiNha + item.DSChiQuay + item.DSCK;
                    }

                    if (!string.IsNullOrEmpty(marketID))
                    {
                        int count = 1;
                        if (marketID != "005")
                        {
                            foreach (ReportDetailtServiceType item in listReportData)
                            {
                                // Add thêm thị trường vào list
                                if (!listMarket.Contains(item.MarketName))
                                {
                                    listMarket.Add(item.MarketName);
                                }
                                item.STT = (count++).ToString();
                                item.ReportID = item.PartnerName;
                                item.TongDS = item.DSChiNha + item.DSChiQuay + item.DSCK;
                            }

                            listReportData.Add(
                                new ReportDetailtServiceType()
                                {
                                    STT = "",
                                    ReportID = "Tổng",
                                    DSChiQuay = listReportData.Sum(x => x.DSChiQuay),
                                    DSChiNha = listReportData.Sum(x => x.DSChiNha),
                                    DSCK = listReportData.Sum(x => x.DSCK),
                                    TongDS = listReportData.Sum(x => x.TongDS),
                                }
                            );
                        }
                        else
                        {
                            List<ReportDetailtServiceType> listDataConvert = new List<ReportDetailtServiceType>();
                            string marketName = string.Empty;
                            foreach (ReportDetailtServiceType item in listReportData)
                            {
                                // Add thêm thị trường vào list
                                if (!listMarket.Contains(item.MarketName))
                                {
                                    // Dòng đầu tiên
                                    // Và chỉ có 1 đối tác của 1 thị trường
                                    if (count == 1)
                                    {
                                        listDataConvert.Add(
                                            new ReportDetailtServiceType()
                                            {
                                                STT = (count++).ToString(),
                                                MarketCode = item.MarketCode,
                                                MarketName = item.MarketName,
                                                PartnerCode = item.PartnerCode,
                                                PartnerName = item.PartnerName,
                                                ReportID = item.PartnerName,
                                                DSChiQuay = item.DSChiQuay,
                                                DSChiNha = item.DSChiNha,
                                                DSCK = item.DSCK,
                                                TongDS = item.DSChiNha + item.DSChiQuay + item.DSCK
                                            }
                                        );
                                    }
                                    else
                                    {
                                        // Trường hợp thuộc dòng cuối mỗi thị trường
                                        // Add dòng tổng cho các đôi tác thị trường trước
                                        listDataConvert.Add(
                                            new ReportDetailtServiceType()
                                            {
                                                STT = "",
                                                PartnerCode = "",
                                                PartnerName = "Tổng",
                                                ReportID = "Tổng",
                                                MarketName = marketName,
                                                DSChiQuay = listReportData.Where(x => x.MarketName == marketName).Sum(x => x.DSChiQuay),
                                                DSChiNha = listReportData.Where(x => x.MarketName == marketName).Sum(x => x.DSChiNha),
                                                DSCK = listReportData.Where(x => x.MarketName == marketName).Sum(x => x.DSCK),
                                                TongDS = listReportData.Where(x => x.MarketName == marketName).Sum(x => x.TongDS),
                                            }
                                        );
                                        // Set stt lại bằng 1
                                        count = 1;

                                        // Số dòng của Thị trường
                                        int countList = listReportData.Where(x => x.MarketCode == item.MarketCode).Count();
                                        if (countList == 1)
                                        {
                                            listDataConvert.Add(
                                                new ReportDetailtServiceType()
                                                {
                                                    STT = (count++).ToString(),
                                                    MarketCode = item.MarketCode,
                                                    MarketName = item.MarketName,
                                                    PartnerCode = item.PartnerCode,
                                                    PartnerName = item.PartnerName,
                                                    ReportID = item.PartnerName,
                                                    DSChiQuay = item.DSChiQuay,
                                                    DSChiNha = item.DSChiNha,
                                                    DSCK = item.DSCK,
                                                    TongDS = item.DSChiNha + item.DSChiQuay + item.DSCK
                                                }
                                            );

                                            // Add dòng tổng vào
                                            listDataConvert.Add(
                                                new ReportDetailtServiceType()
                                                {
                                                    STT = "",
                                                    PartnerCode = "",
                                                    PartnerName = "Tổng",
                                                    ReportID = "Tổng",
                                                    MarketName = item.MarketName,
                                                    DSChiQuay = listReportData.Where(x => x.MarketName == item.MarketName).Sum(x => x.DSChiQuay),
                                                    DSChiNha = listReportData.Where(x => x.MarketName == item.MarketName).Sum(x => x.DSChiNha),
                                                    DSCK = listReportData.Where(x => x.MarketName == item.MarketName).Sum(x => x.DSCK),
                                                    TongDS = listReportData.Where(x => x.MarketName == item.MarketName).Sum(x => x.TongDS),
                                                }
                                            );
                                        }
                                        else
                                        {
                                            listDataConvert.Add(
                                                new ReportDetailtServiceType()
                                                {
                                                    STT = (count++).ToString(),
                                                    MarketCode = item.MarketCode,
                                                    MarketName = item.MarketName,
                                                    PartnerCode = item.PartnerCode,
                                                    PartnerName = item.PartnerName,
                                                    ReportID = item.PartnerName,
                                                    DSChiQuay = item.DSChiQuay,
                                                    DSChiNha = item.DSChiNha,
                                                    DSCK = item.DSCK,
                                                    TongDS = item.DSChiNha + item.DSChiQuay + item.DSCK
                                                }
                                            );
                                        }
                                    }
                                    listMarket.Add(item.MarketName);
                                    marketName = item.MarketName;

                                }
                                else
                                {
                                    ReportDetailtServiceType dataItemLast = listReportData[listReportData.Count - 1];
                                    // Trường hợp khi dòng cuối cần tính tổng
                                    if (dataItemLast.PartnerCode == item.PartnerCode)
                                    {
                                        // Add item cuối vào
                                        listDataConvert.Add(
                                            new ReportDetailtServiceType()
                                            {
                                                STT = (count++).ToString(),
                                                MarketCode = item.MarketCode,
                                                MarketName = item.MarketName,
                                                PartnerCode = item.PartnerCode,
                                                PartnerName = item.PartnerName,
                                                ReportID = item.PartnerName,
                                                DSChiQuay = item.DSChiQuay,
                                                DSChiNha = item.DSChiNha,
                                                DSCK = item.DSCK,
                                                TongDS = item.DSChiNha + item.DSChiQuay + item.DSCK
                                            }
                                        );

                                        // Add dòng tổng vào
                                        listDataConvert.Add(
                                            new ReportDetailtServiceType()
                                            {
                                                STT = "",
                                                PartnerCode = "",
                                                PartnerName = "Tổng",
                                                ReportID = "Tổng",
                                                MarketName = item.MarketName,
                                                DSChiQuay = listReportData.Where(x => x.MarketName == item.MarketName).Sum(x => x.DSChiQuay),
                                                DSChiNha = listReportData.Where(x => x.MarketName == item.MarketName).Sum(x => x.DSChiNha),
                                                DSCK = listReportData.Where(x => x.MarketName == item.MarketName).Sum(x => x.DSCK),
                                                TongDS = listReportData.Where(x => x.MarketName == item.MarketName).Sum(x => x.TongDS),
                                            }
                                        );
                                    }
                                    else
                                    {
                                        // Chưa phải dòng cuối chỉ add item vào
                                        listDataConvert.Add(
                                            new ReportDetailtServiceType()
                                            {
                                                STT = (count++).ToString(),
                                                MarketCode = item.MarketCode,
                                                MarketName = item.MarketName,
                                                PartnerCode = item.PartnerCode,
                                                PartnerName = item.PartnerName,
                                                ReportID = item.PartnerName,
                                                DSChiQuay = item.DSChiQuay,
                                                DSChiNha = item.DSChiNha,
                                                DSCK = item.DSCK,
                                                TongDS = item.DSChiNha + item.DSChiQuay + item.DSCK
                                            }
                                        );
                                    }
                                }
                            }
                            listReportData = new List<ReportDetailtServiceType>(listDataConvert);
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
                ReportDetailtServiceType reportForMaket = new ReportDetailtServiceType()
                {
                    STT = "",
                    ReportID = "Tổng tất cả",
                    DSChiQuay = listReportData.Where(x => x.STT == "").Sum(item => item.DSChiQuay),
                    DSChiNha = listReportData.Where(x => x.STT == "").Sum(item => item.DSChiNha),
                    DSCK = listReportData.Where(x => x.STT == "").Sum(item => item.DSCK),
                    TongDS = listReportData.Where(x => x.STT == "").Sum(item => item.TongDS)
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
                // check trùng
                List<string> listDuplicate = new List<string>();
                // Số dòng của row
                for (int a = 7; a < totalRow; a++)
                {
                    string valueOfTable = string.Empty;

                    if (listMarket.Count.Equals(1) && valueCheck)
                    {
                        // Cho phép chạy qua 1 lần
                        valueCheck = false;
                        valueOfTable = listMarket[stepRow].ToString();
                        // Insert vào dòng cột xác định trong Excel
                        sheetReport.Cells[a, 1].PutValue(valueOfTable, true);
                        // Merge cell
                        Cells cells = sheetReport.Cells;
                        cells.Merge(a, 1, 1, 6);
                        Style style1 = sheetReport.Cells[a, 1].GetStyle();
                        style1.Font.IsBold = true;
                        style1.SetBorder(BorderType.TopBorder, CellBorderType.Thin, Color.Black);
                        style1.SetBorder(BorderType.BottomBorder, CellBorderType.Thin, Color.Black);
                        style1.SetBorder(BorderType.LeftBorder, CellBorderType.Thin, Color.Black);
                        style1.SetBorder(BorderType.RightBorder, CellBorderType.Thin, Color.Black);
                        // Apply the Style to C6 Cell.
                        cells[a, 1].SetStyle(style1);

                        totalRow++;
                        continue;
                    }

                    // Trường hợp nhiều thị trường
                    valueOfTable = dataTable.Rows[stepRow][0].ToString();

                    if (listMarket.Count > 1 && valueOfTable == "1")
                    {
                        if (valueCheck && count < listMarket.Count)
                        {
                            string valueMarket = listMarket[count].ToString();
                            // Insert vào dòng cột xác định trong Excel
                            sheetReport.Cells[a, 1].PutValue(valueMarket, true);
                            // Merge cell
                            Cells cells = sheetReport.Cells;
                            cells.Merge(a, 1, 1, 6);

                            Style style1 = sheetReport.Cells[a, 1].GetStyle();
                            style1.Font.IsBold = true;
                            style1.SetBorder(BorderType.TopBorder, CellBorderType.Thin, Color.Black);
                            style1.SetBorder(BorderType.BottomBorder, CellBorderType.Thin, Color.Black);
                            style1.SetBorder(BorderType.LeftBorder, CellBorderType.Thin, Color.Black);
                            style1.SetBorder(BorderType.RightBorder, CellBorderType.Thin, Color.Black);
                            // Apply the Style to C6 Cell.
                            cells[a, 1].SetStyle(style1);
                            // Tăng tổng số dòng
                            // Get giá trị trước
                            // Tăng tổng số dòng
                            totalRow++;
                            count++;
                            valueCheck = false;
                            continue;
                        }
                        else
                        {
                            // Thay đổi giá trị của value để có thể truy xuất value 1 lần 2
                            valueCheck = true;
                        }
                    }

                    int stepColumn = 0;
                    // Số cột trong báo cáo cần hiển thị
                    // Tổng số cột hiển thị = Số cột hiển thị bắt đầu + tổng số cột cần hiển thị
                    // Số 6 là cột marketName
                    int totalCol = 1 + 6;

                    for (int b = 1; b < totalCol; b++)
                    {
                        // Giá trị của value trong table
                        valueOfTable = dataTable.Rows[stepRow][stepColumn].ToString();

                        // Insert vào dòng cột xác định trong Excel
                        sheetReport.Cells[a, b].PutValue(valueOfTable, true);
                        // set style cho number
                        if (b == 1)
                        {
                            style.Custom = "";
                            style.HorizontalAlignment = TextAlignmentType.Center;
                        }
                        else
                        {
                            style.Custom = "#,##0.00";
                            style.HorizontalAlignment = TextAlignmentType.General;
                        }

                        // Trường hợp với cột tổng của mỗi thị trường
                        if (string.IsNullOrEmpty(dataTable.Rows[stepRow][0].ToString()))
                        {
                            style.Font.IsBold = true;
                        }
                        else
                        {
                            style.Font.IsBold = false;
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

                //// Group 
                //// Create a cellarea i.e.., B3:C19
                //CellArea ca = new CellArea();
                //ca.StartRow = 7;
                //ca.StartColumn = 2;
                //ca.EndRow = dataTable.Rows.Count + 7;
                //ca.EndColumn = 7;

                //// Second column (C) in the list
                //// Group by column index
                //// số cột được tí từ start column
                //sheetReport.Cells.Subtotal(ca, 5, ConsolidationFunction.Sum, new int[] { 1,2,3,4 });
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
            string templatePath = "~/Content/Report/ReportDetailtGradationByMakets.xlsx";
            // Get đường dẫn
            templatePath = System.Web.HttpContext.Current.Server.MapPath(templatePath);

            designer.Workbook = new DongA.Core.DongAExcel().OpenExcelFile(templatePath);
            designer.Workbook.CalculateFormula();

            WorksheetCollection workSheets = designer.Workbook.Worksheets;
            Worksheet sheetReport = designer.Workbook.Worksheets[0];

            // Tạo title
            string typeReport = "So sánh - Theo giai đoạn - Tất cả";

            string text = string.Format( " tháng năm {0}", year);
            switch (gradationID)
            {
                case "1":
                    text = string.Concat("3", text);
                    break;
                case "2":
                    text = string.Concat("6", text);
                    break;
                case "3":
                    text = string.Concat("9", text);
                    break;
                default:
                    text = string.Concat("12", text);
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

            List<ReportDetailtServiceType> listReportData = new List<ReportDetailtServiceType>();

            List<string> listMarket = new List<string>();

            // Trường hợp tất cả thị trường
            if (marketID.Equals("0"))
            {
                listReportData = new ReportBL().ReportDetailtGradationCompareForAll(year, int.Parse(gradationID), reportTypeID, marketID);
                foreach (ReportDetailtServiceType item in listReportData)
                {
                    if (!listMarket.Contains(item.MarketName))
                    {
                        listMarket.Add(item.MarketName);
                    }
                }
            }
            else
            {
                listReportData = new ReportBL().ReportDetailtGradationCompareForAll(year, int.Parse(gradationID), reportTypeID, marketID);
                List<ReportDetailtServiceType> listReportDataConvert = new List<ReportDetailtServiceType>();
                foreach(ReportDetailtServiceType item in listReportData)
                {
                    if(!listMarket.Contains(item.MarketName))
                    {
                        listMarket.Add(item.MarketName);
                    }
                }

                foreach(string item in listMarket)
                {
                    List<ReportDetailtServiceType> listDataItemYear = listReportData.Where(x => x.MarketName == item && x.Year == year.ToString()).ToList();
                    List<ReportDetailtServiceType> listDataItemLastYear = listReportData.Where(x => x.MarketName == item && x.Year == (year -1).ToString()).ToList();
                    if(listDataItemYear.Count > 0)
                    {
                        listReportDataConvert.Add(
                            new ReportDetailtServiceType()
                            {
                                MarketCode = listDataItemYear[0].MarketCode,
                                MarketName = item,
                                DSChiQuay = listDataItemYear.Sum(x => x.DSChiQuay),
                                DSChiNha = listDataItemYear.Sum(x => x.DSChiNha),
                                DSCK = listDataItemYear.Sum(x => x.DSCK),
                                Year = year.ToString()
                            }
                        );
                    }
                    // Last year
                    if (listDataItemLastYear.Count > 0)
                    {
                        listReportDataConvert.Add(
                            new ReportDetailtServiceType()
                            {
                                MarketCode = listDataItemLastYear[0].MarketCode,
                                MarketName = item,
                                DSChiQuay = listDataItemLastYear.Sum(x => x.DSChiQuay),
                                DSChiNha = listDataItemLastYear.Sum(x => x.DSChiNha),
                                DSCK = listDataItemLastYear.Sum(x => x.DSCK),
                                Year = (year - 1).ToString()
                            }
                        );
                    }
                }
                if (listReportDataConvert.Count > 0)
                {
                    listReportData = new List<ReportDetailtServiceType>(listReportDataConvert);
                }
            }

            List<ReportDetailtServiceType> listReportDataClone = new List<ReportDetailtServiceType>(listReportData);
            // Khởi tạo datatable
            DataTable table = new DataTable();

            foreach (ReportDetailtServiceType item in listReportData)
            {
                item.ReportID = item.MarketName;
                item.TongDS = item.DSChiQuay + item.DSChiNha + item.DSCK;
                item.MarketName = "Tất cả thị trường";
            }

            // Khởi tạo datatable
            table = new DataTable();
            // Tạo các cột cho datatable
            table.Columns.Add("ReportID", typeof(String));
            table.Columns.Add("CQ1", typeof(double));
            table.Columns.Add("CQ2", typeof(double));

            table.Columns.Add("CN1", typeof(double));
            table.Columns.Add("CN2", typeof(double));

            table.Columns.Add("CK1", typeof(double));
            table.Columns.Add("CK2", typeof(double));

            table.Columns.Add("TDS1", typeof(double));
            table.Columns.Add("TDS2", typeof(double));
            table.Columns.Add("MarketName", typeof(String));
            
            string reportID = string.Empty;

            if (listReportData.Count() > 0)
            {
                foreach (string item in listMarket)
                {
                    // Cùng kì
                    ReportDetailtServiceType dataItemLastYear = listReportData.Find(x => x.ReportID == item && x.Year == (year - 1).ToString());
                    ReportDetailtServiceType dataItemYear = listReportData.Find(x => x.ReportID == item && x.Year == year.ToString());

                    if (dataItemYear != null && dataItemLastYear != null)
                    {
                        reportID = dataItemYear.ReportID;
                    }

                    // Trường hợp năm không có đối tác
                    if (dataItemLastYear == null && dataItemYear != null)
                    {
                        dataItemLastYear = new ReportDetailtServiceType();
                        dataItemLastYear.MarketName = dataItemYear.MarketName;
                        dataItemLastYear.DSChiQuay = 0;
                        dataItemLastYear.DSChiNha = 0;
                        dataItemLastYear.DSCK = 0;
                        dataItemLastYear.Year = (year - 1).ToString();

                        reportID = dataItemYear.ReportID;
                    }

                    // Trường hợp năm hiện tại không có đối tác
                    if (dataItemYear == null && dataItemLastYear != null)
                    {
                        dataItemYear = new ReportDetailtServiceType();
                        dataItemYear.MarketName = dataItemLastYear.MarketName;
                        dataItemYear.DSChiQuay = 0;
                        dataItemYear.DSChiNha = 0;
                        dataItemYear.DSCK = 0;
                        dataItemYear.Year = year.ToString();

                        reportID = dataItemLastYear.ReportID;
                    }

                    // add item vào table
                    table.Rows.Add(reportID
                        , dataItemLastYear.DSChiQuay, dataItemYear.DSChiQuay
                        , dataItemLastYear.DSChiNha, dataItemYear.DSChiNha
                        , dataItemLastYear.DSCK, dataItemYear.DSCK
                        , dataItemLastYear.TongDS, dataItemYear.TongDS
                        , dataItemLastYear.MarketName);
                }

                DataRow row = table.NewRow();
                row["ReportID"] = "Tổng";
                row["CQ1"] = table.Compute("Sum(CQ1)", "");
                row["CQ2"] = table.Compute("Sum(CQ2)", "");

                row["CN1"] = table.Compute("Sum(CN1)", "");
                row["CN2"] = table.Compute("Sum(CN2)", "");

                row["CK1"] = table.Compute("Sum(CK1)", "");
                row["CK2"] = table.Compute("Sum(CK2)", "");

                row["TDS1"] = table.Compute("Sum(TDS1)", "");
                row["TDS2"] = table.Compute("Sum(TDS2)", "");

                row["MarketName"] = "";
                table.Rows.Add(row);

            }
            
            // Set border
            Style style = new CellsFactory().CreateStyle();
            style.SetBorder(BorderType.BottomBorder, CellBorderType.Thin, Color.Black);
            style.SetBorder(BorderType.LeftBorder, CellBorderType.Thin, Color.Black);
            style.SetBorder(BorderType.RightBorder, CellBorderType.Thin, Color.Black);

            if (table.Rows.Count > 0)
            {
                int stepRow = 0;
                // total row = row start + số row hiện có
                int totalRow = table.Rows.Count + 62;
                // Số dòng của row
                for (int a = 62; a < totalRow; a++)
                {
                    int stepColumn = 0;
                    // Số cột trong báo cáo cần hiển thị
                    // Tổng số cột hiển thị = Số cột hiển thị bắt đầu + tổng số cột cần hiển thị
                    int totalCol = 1 + 9;
                    for (int b = 1; b < totalCol; b++)
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

            // Vẽ biểu đồ cột cho Excel
            //Chart reference
            Aspose.Cells.Charts.Chart leadSourceColumnChiNha;
            //Add Pie Chart
            // Chi Nhà
            int chartIndex = sheetReport.Charts.Add(ChartType.Bar, 25, 0, 40, 7);
            leadSourceColumnChiNha = sheetReport.Charts[chartIndex];

            //Chart title
            leadSourceColumnChiNha.Title.Text = string.Format("Doanh số dịch vụ chi quầy từng thị trường \n Giai đoạn: {0} \n Chi Nhà", text);
            leadSourceColumnChiNha.Title.Font.Color = Color.Silver;

            // Adding SeriesCollection (chart data source) to the chart ranging from "A1" cell to "B3"
            string totalRowData = string.Format("C63:D{0}", 63 + table.Rows.Count - 2);
            leadSourceColumnChiNha.NSeries.Add(totalRowData, true);

            // Set the category data covering the range A2:A5.
            // Tổng số dòng cần hiển thị là số dòng hiện tại  + số dòng trong listData -2 (trong đó 1 là dòng tổng cuối cùng)
            string categoryData = "B63:B68";
            leadSourceColumnChiNha.NSeries.CategoryData = categoryData;

            // Set the names of the chart series taken from cells.
            leadSourceColumnChiNha.NSeries[0].Name = "=C62";
            leadSourceColumnChiNha.NSeries[1].Name = "=D62";

            // Set the 1st series fill color.
            leadSourceColumnChiNha.NSeries[0].Area.ForegroundColor = Color.Orange;
            leadSourceColumnChiNha.NSeries[0].Area.Formatting = FormattingType.Custom;

            // Set the 2nd series fill color.
            leadSourceColumnChiNha.NSeries[1].Area.ForegroundColor = Color.Green;
            leadSourceColumnChiNha.NSeries[1].Area.Formatting = FormattingType.Custom;


            // Set plot area formatting as none and hide its border.
            leadSourceColumnChiNha.PlotArea.Area.FillFormat.FillType = FillType.None;
            leadSourceColumnChiNha.PlotArea.Border.IsVisible = false;

            // Set value axis major tick mark as none and hide axis line. 
            // Also set the color of value axis major grid lines.
            leadSourceColumnChiNha.ValueAxis.MajorTickMark = TickMarkType.None;
            leadSourceColumnChiNha.ValueAxis.AxisLine.IsVisible = false;
            //leadSourceColumnChiNha.ValueAxis.IsAutomaticMajorUnit = false;
            //leadSourceColumnChiNha.ValueAxis.MajorUnit = 10000000;
            leadSourceColumnChiNha.ValueAxis.MajorGridLines.Color = Color.FromArgb(217, 217, 217);

            //Chart reference
            Aspose.Cells.Charts.Chart leadSourceColumnChiQuay;
            // Chi Quầy
            chartIndex = sheetReport.Charts.Add(ChartType.Bar, 25, 8, 40, 15);
            leadSourceColumnChiQuay = sheetReport.Charts[chartIndex];

            //Chart title
            leadSourceColumnChiQuay.Title.Text = string.Format("Doanh số dịch vụ chi quầy từng thị trường \n Giai đoạn: {0} \n Chi Quầy", text);
            leadSourceColumnChiQuay.Title.Font.Color = Color.Silver;

            // Adding SeriesCollection (chart data source) to the chart ranging from "A1" cell to "B3"
            totalRowData = string.Format("E63:F{0}", 63 + table.Rows.Count - 2);
            leadSourceColumnChiQuay.NSeries.Add(totalRowData, true);

            // Set the category data covering the range A2:A5.
            // Tổng số dòng cần hiển thị là số dòng hiện tại  + số dòng trong listData -2 (trong đó 1 là dòng tổng cuối cùng)
            categoryData = "B63:B68";
            leadSourceColumnChiQuay.NSeries.CategoryData = categoryData;

            // Set the names of the chart series taken from cells.
            leadSourceColumnChiQuay.NSeries[0].Name = "=C62";
            leadSourceColumnChiQuay.NSeries[1].Name = "=D62";

            // Set the 1st series fill color.
            leadSourceColumnChiQuay.NSeries[0].Area.ForegroundColor = Color.Orange;
            leadSourceColumnChiQuay.NSeries[0].Area.Formatting = FormattingType.Custom;

            // Set the 2nd series fill color.
            leadSourceColumnChiQuay.NSeries[1].Area.ForegroundColor = Color.Green;
            leadSourceColumnChiQuay.NSeries[1].Area.Formatting = FormattingType.Custom;


            // Set plot area formatting as none and hide its border.
            leadSourceColumnChiQuay.PlotArea.Area.FillFormat.FillType = FillType.None;
            leadSourceColumnChiQuay.PlotArea.Border.IsVisible = false;

            // Set value axis major tick mark as none and hide axis line. 
            // Also set the color of value axis major grid lines.
            leadSourceColumnChiQuay.ValueAxis.MajorTickMark = TickMarkType.None;
            leadSourceColumnChiQuay.ValueAxis.AxisLine.IsVisible = false;
            leadSourceColumnChiQuay.ValueAxis.MajorGridLines.Color = Color.FromArgb(217, 217, 217);


            //Chart reference
            Aspose.Cells.Charts.Chart leadSourceColumnChuyenKhoan;
            // Chuyển khoản
            chartIndex = sheetReport.Charts.Add(ChartType.Bar, 41, 0, 56, 7);
            leadSourceColumnChuyenKhoan = sheetReport.Charts[chartIndex];

            //Chart title
            leadSourceColumnChuyenKhoan.Title.Text = string.Format("Doanh số dịch vụ chi quầy từng thị trường \n Giai đoạn: {0} \n Chuyển khoản", text);
            leadSourceColumnChuyenKhoan.Title.Font.Color = Color.Silver;

            // Adding SeriesCollection (chart data source) to the chart ranging from "A1" cell to "B3"
            totalRowData = string.Format("G63:H{0}", 63 + table.Rows.Count - 2);
            leadSourceColumnChuyenKhoan.NSeries.Add(totalRowData, true);

            // Set the category data covering the range A2:A5.
            // Tổng số dòng cần hiển thị là số dòng hiện tại  + số dòng trong listData -2 (trong đó 1 là dòng tổng cuối cùng)
            categoryData = "B63:B68";
            leadSourceColumnChuyenKhoan.NSeries.CategoryData = categoryData;

            // Set the names of the chart series taken from cells.
            leadSourceColumnChuyenKhoan.NSeries[0].Name = "=C62";
            leadSourceColumnChuyenKhoan.NSeries[1].Name = "=D62";

            // Set the 1st series fill color.
            leadSourceColumnChuyenKhoan.NSeries[0].Area.ForegroundColor = Color.Orange;
            leadSourceColumnChuyenKhoan.NSeries[0].Area.Formatting = FormattingType.Custom;

            // Set the 2nd series fill color.
            leadSourceColumnChuyenKhoan.NSeries[1].Area.ForegroundColor = Color.Green;
            leadSourceColumnChuyenKhoan.NSeries[1].Area.Formatting = FormattingType.Custom;


            // Set plot area formatting as none and hide its border.
            leadSourceColumnChuyenKhoan.PlotArea.Area.FillFormat.FillType = FillType.None;
            leadSourceColumnChuyenKhoan.PlotArea.Border.IsVisible = false;

            // Set value axis major tick mark as none and hide axis line. 
            // Also set the color of value axis major grid lines.
            leadSourceColumnChuyenKhoan.ValueAxis.MajorTickMark = TickMarkType.None;
            leadSourceColumnChuyenKhoan.ValueAxis.AxisLine.IsVisible = false;
            leadSourceColumnChuyenKhoan.ValueAxis.MajorGridLines.Color = Color.FromArgb(217, 217, 217);


            // Tạo chart cột tỉ trọng cho các thị trường
            //Add Pie Chart
            //Chart reference
            Aspose.Cells.Charts.Chart leadSourceColumn;
            chartIndex = sheetReport.Charts.Add(ChartType.Column3DStacked, 6, 0, 24, 13);
            leadSourceColumn = sheetReport.Charts[chartIndex];


            //Chart title
            leadSourceColumn.Title.Text = string.Format("Tỉ trọng từng thị trường \n Giai đoạn: {0} \n", text);
            leadSourceColumn.Title.Font.Color = Color.Silver;

            // count cho 3 loại: chi Quầy, Chi nhà, Chuyển khoản
            // list thị trường

            List<ReportDetailtServiceType> listDataGradation = new ReportBL().ReportDetailtGradationCompareForAllPercent(year, int.Parse(gradationID), reportTypeID, marketID);

            List<string> listMarketCurrent = new List<string>();
            foreach(ReportDetailtServiceType item in listDataGradation)
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
                List<ReportDetailtServiceType> dataItemLastYear = listDataGradation.Where(x => x.MarketName == item && x.Year == (year - 1).ToString()).ToList();
                List<ReportDetailtServiceType> dataItemYear = listDataGradation.Where(x => x.MarketName == item && x.Year == year.ToString()).ToList();


                listTotalRowData[i++] = string.Concat("{"
                    , string.Format("{0}, {1}, {2}, {3}, {4}, {5}"
                    , dataItemYear.Sum(x =>x.DSChiQuay), dataItemLastYear.Sum(x => x.DSChiQuay)
                    , dataItemYear.Sum(x => x.DSChiNha), dataItemLastYear.Sum(x => x.DSChiNha)
                    , dataItemYear.Sum(x => x.DSCK), dataItemLastYear.Sum(x => x.DSCK))
                    , "}");
            }

            int count = 0;
            string yearCurent = string.Format("Năm {0}", year);
            string lastYear = string.Format("Năm {0}", year - 1);
            foreach (string item in listTotalRowData)
            {
                totalRowData = item;
                leadSourceColumn.NSeries.Add(totalRowData, true);

                categoryData = string.Concat("{",string.Format("Chi Quầy {0}, Chi Quầy {1}, Chi Nhà {0}, Chi Nhà {1}, Chuyển Khoản {0},  Chuyển Khoản {1}", yearCurent, lastYear), "}");
                leadSourceColumn.NSeries.CategoryData = categoryData;

                leadSourceColumn.NSeries[count].Name = listMarketCurrent[count];
                
                switch(count)
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
        public ActionResult CreateExcelGradationCompareForOne(string gradationID, int year, string reportTypeID, string marketID)
        {
            WorkbookDesigner designer = new WorkbookDesigner();
            string templatePath = "~/Content/Report/ReportDetailtGradationByMakets.xlsx";
            // Get đường dẫn
            templatePath = System.Web.HttpContext.Current.Server.MapPath(templatePath);

            designer.Workbook = new DongA.Core.DongAExcel().OpenExcelFile(templatePath);
            designer.Workbook.CalculateFormula();

            WorksheetCollection workSheets = designer.Workbook.Worksheets;
            Worksheet sheetReport = designer.Workbook.Worksheets[0];

            // Tạo title
            string typeReport = "So sánh - Theo giai đoạn - Tất cả";

            string text = string.Format(" tháng năm {0}", year);
            switch (gradationID)
            {
                case "1":
                    text = string.Concat("3", text);
                    break;
                case "2":
                    text = string.Concat("6", text);
                    break;
                case "3":
                    text = string.Concat("9", text);
                    break;
                default:
                    text = string.Concat("12", text);
                    break;
            }

            // Tạo title
            CreateTitle("A2", "K2", sheetReport, typeReport, 14);

            // Tạo title detailt
            string titleDetailt = string.Format("Giai đoạn: {0}", text);
            CreateTitle("A3", "K3", sheetReport, titleDetailt, 12);
            
            titleDetailt = "1. Theo doanh số chi trả theo thị trường - loại hình chi trả";
            CreateTitle("B6", "E6", sheetReport, titleDetailt, 12);

            // Tạo giá trị cho cột dữ liệu của Chi quầy/ Chi nhà/ Chuyển khoản
            sheetReport.Cells["C62"].PutValue(string.Format("Năm {0} ", year - 1));
            sheetReport.Cells["D62"].PutValue(string.Format("Năm {0} ", year));

            sheetReport.Cells["E62"].PutValue(string.Format("Năm {0} ", year - 1));
            sheetReport.Cells["F62"].PutValue(string.Format("Năm {0} ", year));

            sheetReport.Cells["G62"].PutValue(string.Format("Năm {0} ", year - 1));
            sheetReport.Cells["H62"].PutValue(string.Format("Năm {0} ", year));

            sheetReport.Cells["I62"].PutValue(string.Format("Năm {0} ", year - 1));
            sheetReport.Cells["J62"].PutValue(string.Format("Năm {0} ", year));

            List<ReportDetailtServiceType> listReportData = new ReportBL().ReportDetailtGradationCompareForOne(year, int.Parse(gradationID), reportTypeID, marketID);

            List<string> listMarket = new List<string>();

            // Trường hợp tất cả thị trường
            if (marketID.Equals("005"))
            {
                // List data
                List<ReportDetailtServiceType> listReportDataConvert = new List<ReportDetailtServiceType>();
                listMarket = new List<string>();

                foreach (ReportDetailtServiceType item in listReportData)
                {
                    if (!listMarket.Contains(item.MarketName))
                    {
                        listMarket.Add(item.MarketName);
                    }
                }

                foreach (string item in listMarket)
                {
                    List<ReportDetailtServiceType> listDataItemYear = listReportData.Where(x => x.MarketName == item && x.Year == year.ToString()).ToList();
                    List<ReportDetailtServiceType> listDataItemLastYear = listReportData.Where(x => x.MarketName == item && x.Year == (year - 1).ToString()).ToList();

                    // Year
                    listReportDataConvert.Add(
                        new ReportDetailtServiceType()
                        {
                            MarketName = item,
                            PartnerName = item,
                            DSChiQuay = listDataItemYear.Sum(x => x.DSChiQuay),
                            DSChiNha = listDataItemYear.Sum(x => x.DSChiNha),
                            DSCK = listDataItemYear.Sum(x => x.DSCK),
                            TongDS = listDataItemYear.Sum(x => x.TongDS),
                            Year = year.ToString()
                        }
                    );

                    // Last Year
                    listReportDataConvert.Add(
                        new ReportDetailtServiceType()
                        {
                            MarketName = item,
                            PartnerName = item,
                            DSChiQuay = listDataItemLastYear.Sum(x => x.DSChiQuay),
                            DSChiNha = listDataItemLastYear.Sum(x => x.DSChiNha),
                            DSCK = listDataItemLastYear.Sum(x => x.DSCK),
                            TongDS = listDataItemLastYear.Sum(x => x.TongDS),
                            Year = (year - 1).ToString()
                        }
                    );
                }

                if (listReportDataConvert.Count > 0)
                {
                    listReportData = new List<ReportDetailtServiceType>(listReportDataConvert);
                }
            }

            // Khởi tạo datatable
            DataTable table = new DataTable();
            // Tạo các cột cho datatable
            table.Columns.Add("PartnerName", typeof(String));

            table.Columns.Add("CQ1", typeof(double));
            table.Columns.Add("CQ2", typeof(double));

            table.Columns.Add("CN1", typeof(double));
            table.Columns.Add("CN2", typeof(double));

            table.Columns.Add("CK1", typeof(double));
            table.Columns.Add("CK2", typeof(double));

            table.Columns.Add("TDS1", typeof(double));
            table.Columns.Add("TDS2", typeof(double));

            table.Columns.Add("MarketName", typeof(String));

            if (marketID.Equals("005"))
            {
                foreach (ReportDetailtServiceType item in listReportData)
                {
                    item.PartnerName = item.MarketName;
                    item.TongDS = item.DSChiQuay + item.DSChiNha + item.DSCK;
                    item.MarketName = "Thị trường Châu Á";
                }
            }
            else
            {
                foreach (ReportDetailtServiceType item in listReportData)
                {
                    item.TongDS = item.DSChiQuay + item.DSChiNha + item.DSCK;
                }
            }

            string reportID = string.Empty;

            // Table dữ liệu bảng số liệu Doanh số Chi Quầy/Chi Nhà/Chuyển Khoản
            foreach (ReportDetailtServiceType item in listReportData)
            {
                // Cùng kì
                ReportDetailtServiceType dataItemLastYear = listReportData.Find(x => x.PartnerName == item.PartnerName && x.Year == (year - 1).ToString());
                ReportDetailtServiceType dataItemYear = listReportData.Find(x => x.PartnerName == item.PartnerName && x.Year == year.ToString());

                // Trường hợp năm trước có đối tác và năm nay không có
                if (dataItemLastYear != null && dataItemYear == null)
                {
                    dataItemYear = new ReportDetailtServiceType();
                    dataItemYear.PartnerName = dataItemLastYear.PartnerName;
                    dataItemYear.MarketName = dataItemLastYear.MarketName;
                    dataItemYear.DSChiQuay = 0;
                    dataItemYear.DSChiNha = 0;
                    dataItemYear.DSCK = 0;
                    dataItemYear.Year = year.ToString();
                }

                // Trường hợp năm trước không có đối tác và năm nay có đối tác
                if (dataItemYear != null && dataItemLastYear == null)
                {
                    dataItemLastYear = new ReportDetailtServiceType();
                    dataItemLastYear.PartnerName = dataItemYear.PartnerName;
                    dataItemLastYear.MarketName = dataItemYear.MarketName;
                    dataItemLastYear.DSChiQuay = 0;
                    dataItemLastYear.DSChiNha = 0;
                    dataItemLastYear.DSCK = 0;
                    dataItemLastYear.Year = (year - 1).ToString();
                }

                // Check tồn tại của item
                string value = string.Format("PartnerName='{0}'", item.PartnerName);
                DataRow[] foundRows = table.Select(value);
                if (dataItemLastYear != null && dataItemYear != null && foundRows.Count() == 0)
                {
                    // add item vào table
                    table.Rows.Add(dataItemLastYear.PartnerName
                        , dataItemLastYear.DSChiQuay, dataItemYear.DSChiQuay
                        , dataItemLastYear.DSChiNha, dataItemYear.DSChiNha
                        , dataItemLastYear.DSCK, dataItemYear.DSCK
                        , dataItemLastYear.TongDS, dataItemYear.TongDS
                        , dataItemLastYear.MarketName);
                }
            }

            DataRow row = table.NewRow();
            row["PartnerName"] = "Tổng";
            row["CQ1"] = table.Compute("Sum(CQ1)", "");
            row["CQ2"] = table.Compute("Sum(CQ2)", "");

            row["CN1"] = table.Compute("Sum(CN1)", "");
            row["CN2"] = table.Compute("Sum(CN2)", "");

            row["CK1"] = table.Compute("Sum(CK1)", "");
            row["CK2"] = table.Compute("Sum(CK2)", "");

            row["TDS1"] = table.Compute("Sum(TDS1)", "");
            row["TDS2"] = table.Compute("Sum(TDS2)", "");
            table.Rows.Add(row);

            // Set border
            Style style = new CellsFactory().CreateStyle();
            style.SetBorder(BorderType.BottomBorder, CellBorderType.Thin, Color.Black);
            style.SetBorder(BorderType.LeftBorder, CellBorderType.Thin, Color.Black);
            style.SetBorder(BorderType.RightBorder, CellBorderType.Thin, Color.Black);
            style.SetBorder(BorderType.TopBorder, CellBorderType.Thin, Color.Black);

            // Tổng số row theo table1
            int totalRowTable1 = table.Rows.Count + 62;

            // Table dữ liệu bảng số liệu Doanh số Chi Quầy/Chi Nhà/Chuyển Khoản
            if (table.Rows.Count > 0)
            {
                int stepRow = 0;
                // total row = row start + số row hiện có
                int totalRow = table.Rows.Count + 62;
                // Số dòng của row
                for (int a = 62; a < totalRow; a++)
                {
                    int stepColumn = 0;
                    // Số cột trong báo cáo cần hiển thị
                    // Tổng số cột hiển thị = Số cột hiển thị bắt đầu + tổng số cột cần hiển thị
                    int totalCol = 1 + 9;
                    for (int b = 1; b < totalCol; b++)
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


            // Tạo table dữ liệu tăng giảm cho So sánh theo giai đoạn cho từng thị trường
            // Khởi tạo datatable
            table = new DataTable();
            // Tạo các cột cho datatable
            table.Columns.Add("PartnerName", typeof(String));
            table.Columns.Add("CQ1", typeof(double));
            table.Columns.Add("CN1", typeof(double));
            table.Columns.Add("CK1", typeof(double));
            table.Columns.Add("TDS1", typeof(double));

            foreach (ReportDetailtServiceType item in listReportData)
            {
                // Cùng kì
                ReportDetailtServiceType dataItemLastYear = listReportData.Find(x => x.PartnerName == item.PartnerName && x.Year == (year - 1).ToString());
                ReportDetailtServiceType dataItemYear = listReportData.Find(x => x.PartnerName == item.PartnerName && x.Year == year.ToString());

                // Trường hợp năm trước có đối tác và năm nay không có
                if (dataItemLastYear != null && dataItemYear == null)
                {
                    dataItemYear = new ReportDetailtServiceType();
                    dataItemYear.PartnerName = dataItemLastYear.PartnerName;
                    dataItemYear.DSChiQuay = 0;
                    dataItemYear.DSChiNha = 0;
                    dataItemYear.DSCK = 0;
                    dataItemYear.Year = year.ToString();
                }

                // Trường hợp năm trước không có đối tác và năm nay có đối tác
                if (dataItemYear != null && dataItemLastYear == null)
                {
                    dataItemLastYear = new ReportDetailtServiceType();
                    dataItemLastYear.PartnerName = dataItemYear.PartnerName;
                    dataItemLastYear.DSChiQuay = 0;
                    dataItemLastYear.DSChiNha = 0;
                    dataItemLastYear.DSCK = 0;
                    dataItemLastYear.Year = (year - 1).ToString();
                }

                // Check tồn tại của item
                string value = string.Format("PartnerName='{0}'", item.PartnerName);
                DataRow[] foundRows = table.Select(value);

                if (dataItemLastYear != null && dataItemYear != null && foundRows.Count() == 0)
                {
                    // add item vào table
                    table.Rows.Add(dataItemLastYear.PartnerName
                        , dataItemYear.DSChiQuay - dataItemLastYear.DSChiQuay
                        , dataItemYear.DSChiNha - dataItemLastYear.DSChiNha
                        , dataItemYear.DSCK - dataItemLastYear.DSCK
                        , dataItemYear.TongDS - dataItemLastYear.TongDS);
                }
            }

            row = table.NewRow();
            row["PartnerName"] = "Tổng";
            row["CQ1"] = table.Compute("Sum(CQ1)", "");
            row["CN1"] = table.Compute("Sum(CN1)", "");
            row["CK1"] = table.Compute("Sum(CK1)", "");
            row["TDS1"] = table.Compute("Sum(TDS1)", "");
            table.Rows.Add(row);

            // Tổng số row của table2
            // Với 6 là số cách của table1 và table2
            int totalRowTable2 = totalRowTable1 + table.Rows.Count + 6;
            // Tạo title hearder cho table tăng giảm
            // Title cho thị trường
            string title = "Đối tác";
            if (marketID.Contains("005"))
            {
                title = "Thị trường";
            }
            CreateTitle(string.Format("B{0}", totalRowTable1 + 6 - 1), string.Format("B{0}", totalRowTable1 + 6), sheetReport, title, 12, true);
            
            title = string.Format("Tăng/Giảm so với cùng kì năm {0}", year - 1);
            CreateTitle(string.Format("C{0}", totalRowTable1 + 6 - 1), string.Format("F{0}", totalRowTable1 + 6 - 1), sheetReport, title, 12, true);

            title = string.Format("Chi Quầy");
            CreateTitle(string.Format("C{0}", totalRowTable1 + 6), string.Format("C{0}", totalRowTable1 + 6), sheetReport, title, 12, true);

            title = string.Format("Chi Nhà");
            CreateTitle(string.Format("D{0}", totalRowTable1 + 6), string.Format("D{0}", totalRowTable1 + 6), sheetReport, title, 12, true);

            title = string.Format("Chuyển Khoản");
            CreateTitle(string.Format("E{0}", totalRowTable1 + 6), string.Format("E{0}", totalRowTable1 + 6), sheetReport, title, 12, true);

            title = string.Format("Tổng");
            CreateTitle(string.Format("F{0}", totalRowTable1 + 6), string.Format("F{0}", totalRowTable1 + 6), sheetReport, title, 12, true);


            // Table dữ liệu bảng số liệu Doanh số Chi Quầy/Chi Nhà/Chuyển Khoản
            if (table.Rows.Count > 0)
            {
                int stepRow = 0;
                // total row = row start + số row hiện có
                int totalRow = totalRowTable2;
                int rowStart = totalRowTable1 + 6;
                // Số dòng của row
                for (int a = rowStart; a < totalRow; a++)
                {
                    int stepColumn = 0;
                    // Số cột trong báo cáo cần hiển thị
                    // Tổng số cột hiển thị = Số cột hiển thị bắt đầu + tổng số cột cần hiển thị
                    int totalCol = 1 + 5;
                    for (int b = 1; b < totalCol; b++)
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

            // Vẽ biểu đồ cột cho Excel
            //Chart reference
            Aspose.Cells.Charts.Chart leadSourceColumnChiNha;
            //Add Pie Chart
            // Chi Nhà
            int chartIndex = sheetReport.Charts.Add(ChartType.Column3DClustered, 25, 0, 40, 7);
            leadSourceColumnChiNha = sheetReport.Charts[chartIndex];

            //Chart title
            leadSourceColumnChiNha.Title.Text = string.Format("Doanh số dịch vụ chi quầy từng thị trường \n Giai đoạn: {0} \n Chi Nhà", text);
            leadSourceColumnChiNha.Title.Font.Color = Color.Silver;

            // Adding SeriesCollection (chart data source) to the chart ranging from "A1" cell to "B3"
            string totalRowData = string.Format("C63:D{0}", totalRowTable1 - 1);
            leadSourceColumnChiNha.NSeries.Add(totalRowData, true);

            // Set the category data covering the range A2:A5.
            // Tổng số dòng cần hiển thị là số dòng hiện tại  + số dòng trong listData -2 (trong đó 1 là dòng tổng cuối cùng)
            string categoryData = string.Format("B63:B{0}", totalRowTable1 - 1);
            leadSourceColumnChiNha.NSeries.CategoryData = categoryData;

            // Set the names of the chart series taken from cells.
            leadSourceColumnChiNha.NSeries[0].Name = "=C62";
            leadSourceColumnChiNha.NSeries[1].Name = "=D62";

            // Set the 1st series fill color.
            leadSourceColumnChiNha.NSeries[0].Area.ForegroundColor = Color.Orange;
            leadSourceColumnChiNha.NSeries[0].Area.Formatting = FormattingType.Custom;

            // Set the 2nd series fill color.
            leadSourceColumnChiNha.NSeries[1].Area.ForegroundColor = Color.Green;
            leadSourceColumnChiNha.NSeries[1].Area.Formatting = FormattingType.Custom;


            // Set plot area formatting as none and hide its border.
            leadSourceColumnChiNha.PlotArea.Area.FillFormat.FillType = FillType.None;
            leadSourceColumnChiNha.PlotArea.Border.IsVisible = false;

            // Set value axis major tick mark as none and hide axis line. 
            // Also set the color of value axis major grid lines.
            leadSourceColumnChiNha.ValueAxis.MajorTickMark = TickMarkType.None;
            leadSourceColumnChiNha.ValueAxis.AxisLine.IsVisible = false;
            //leadSourceColumnChiNha.ValueAxis.IsAutomaticMajorUnit = false;
            //leadSourceColumnChiNha.ValueAxis.MajorUnit = 10000000;
            leadSourceColumnChiNha.ValueAxis.MajorGridLines.Color = Color.FromArgb(217, 217, 217);

            //Chart reference
            Aspose.Cells.Charts.Chart leadSourceColumnChiQuay;
            // Chi Quầy
            chartIndex = sheetReport.Charts.Add(ChartType.Column3DClustered, 25, 8, 40, 15);
            leadSourceColumnChiQuay = sheetReport.Charts[chartIndex];

            //Chart title
            leadSourceColumnChiQuay.Title.Text = string.Format("Doanh số dịch vụ chi quầy từng thị trường \n Giai đoạn: {0} \n Chi Quầy", text);
            leadSourceColumnChiQuay.Title.Font.Color = Color.Silver;

            // Adding SeriesCollection (chart data source) to the chart ranging from "A1" cell to "B3"
            totalRowData = string.Format("E63:F{0}", totalRowTable1 - 1);
            leadSourceColumnChiQuay.NSeries.Add(totalRowData, true);

            // Set the category data covering the range A2:A5.
            // Tổng số dòng cần hiển thị là số dòng hiện tại  + số dòng trong listData -2 (trong đó 1 là dòng tổng cuối cùng)
            categoryData = string.Format("B63:B{0}", totalRowTable1 - 1);
            leadSourceColumnChiQuay.NSeries.CategoryData = categoryData;

            // Set the names of the chart series taken from cells.
            leadSourceColumnChiQuay.NSeries[0].Name = "=C62";
            leadSourceColumnChiQuay.NSeries[1].Name = "=D62";

            // Set the 1st series fill color.
            leadSourceColumnChiQuay.NSeries[0].Area.ForegroundColor = Color.Orange;
            leadSourceColumnChiQuay.NSeries[0].Area.Formatting = FormattingType.Custom;

            // Set the 2nd series fill color.
            leadSourceColumnChiQuay.NSeries[1].Area.ForegroundColor = Color.Green;
            leadSourceColumnChiQuay.NSeries[1].Area.Formatting = FormattingType.Custom;


            // Set plot area formatting as none and hide its border.
            leadSourceColumnChiQuay.PlotArea.Area.FillFormat.FillType = FillType.None;
            leadSourceColumnChiQuay.PlotArea.Border.IsVisible = false;

            // Set value axis major tick mark as none and hide axis line. 
            // Also set the color of value axis major grid lines.
            leadSourceColumnChiQuay.ValueAxis.MajorTickMark = TickMarkType.None;
            leadSourceColumnChiQuay.ValueAxis.AxisLine.IsVisible = false;
            leadSourceColumnChiQuay.ValueAxis.MajorGridLines.Color = Color.FromArgb(217, 217, 217);


            //Chart reference
            Aspose.Cells.Charts.Chart leadSourceColumnChuyenKhoan;
            // Chuyển khoản
            chartIndex = sheetReport.Charts.Add(ChartType.Column3DClustered, 41, 0, 56, 7);
            leadSourceColumnChuyenKhoan = sheetReport.Charts[chartIndex];

            //Chart title
            leadSourceColumnChuyenKhoan.Title.Text = string.Format("Doanh số dịch vụ chi quầy từng thị trường \n Giai đoạn: {0} \n Chuyển khoản", text);
            leadSourceColumnChuyenKhoan.Title.Font.Color = Color.Silver;

            // Adding SeriesCollection (chart data source) to the chart ranging from "A1" cell to "B3"
            totalRowData = string.Format("G63:H{0}", totalRowTable1 - 1);
            leadSourceColumnChuyenKhoan.NSeries.Add(totalRowData, true);

            // Set the category data covering the range A2:A5.
            // Tổng số dòng cần hiển thị là số dòng hiện tại  + số dòng trong listData -2 (trong đó 1 là dòng tổng cuối cùng)
            categoryData = string.Format("B63:B{0}", totalRowTable1 - 1);
            leadSourceColumnChuyenKhoan.NSeries.CategoryData = categoryData;

            // Set the names of the chart series taken from cells.
            leadSourceColumnChuyenKhoan.NSeries[0].Name = "=C62";
            leadSourceColumnChuyenKhoan.NSeries[1].Name = "=D62";

            // Set the 1st series fill color.
            leadSourceColumnChuyenKhoan.NSeries[0].Area.ForegroundColor = Color.Orange;
            leadSourceColumnChuyenKhoan.NSeries[0].Area.Formatting = FormattingType.Custom;

            // Set the 2nd series fill color.
            leadSourceColumnChuyenKhoan.NSeries[1].Area.ForegroundColor = Color.Green;
            leadSourceColumnChuyenKhoan.NSeries[1].Area.Formatting = FormattingType.Custom;


            // Set plot area formatting as none and hide its border.
            leadSourceColumnChuyenKhoan.PlotArea.Area.FillFormat.FillType = FillType.None;
            leadSourceColumnChuyenKhoan.PlotArea.Border.IsVisible = false;

            // Set value axis major tick mark as none and hide axis line. 
            // Also set the color of value axis major grid lines.
            leadSourceColumnChuyenKhoan.ValueAxis.MajorTickMark = TickMarkType.None;
            leadSourceColumnChuyenKhoan.ValueAxis.AxisLine.IsVisible = false;
            leadSourceColumnChuyenKhoan.ValueAxis.MajorGridLines.Color = Color.FromArgb(217, 217, 217);


            // Tạo chart cột tỉ trọng cho các thị trường
            //Add Pie Chart
            //Chart reference
            Aspose.Cells.Charts.Chart leadSourceColumn;
            chartIndex = sheetReport.Charts.Add(ChartType.Bar, 6, 0, 24, 13);
            leadSourceColumn = sheetReport.Charts[chartIndex];


            //Chart title
            leadSourceColumn.Title.Text = string.Format("Doanh số các đối tác \n Giai đoạn: {0} \n", text);
            leadSourceColumn.Title.Font.Color = Color.Silver;

            // count cho 3 loại: chi Quầy, Chi nhà, Chuyển khoản
            // list thị trường

            string textValue = "T";
            switch (gradationID)
            {
                case "1":
                    textValue = string.Concat("3", textValue);
                    break;
                case "2":
                    textValue = string.Concat("6", textValue);
                    break;
                case "3":
                    textValue = string.Concat("9", textValue);
                    break;
                default:
                    textValue = string.Concat("12", textValue);
                    break;
            }

            
            // Danh sách các đối tác/Thị trường
            List<string> listPartner = new List<string>();

            foreach(ReportDetailtServiceType item in listReportData)
            {
                if (!listPartner.Contains(item.PartnerName))
                {
                    listPartner.Add(item.PartnerName);
                }
            }
            
            int i = 0;
            while (i < 2)
            {
                if (i == 0)
                {
                    List<double> sumYear = new List<double>();

                    List<ReportDetailtServiceType> dataItemYear = listReportData.Where(x => x.Year == year.ToString()).ToList();

                    sumYear = dataItemYear.Select(x => x.TongDS).ToList();

                    totalRowData = string.Concat("{", string.Join(", ", sumYear), "}");
                    leadSourceColumn.NSeries.Add(totalRowData, true);

                    categoryData = string.Concat("{", string.Join(", ", listPartner), "}");
                    leadSourceColumn.NSeries.CategoryData = categoryData;

                    leadSourceColumn.NSeries[i].Name = string.Format("Tổng {0} {1}", textValue, year);

                    leadSourceColumn.NSeries[i].Area.ForegroundColor = Color.Orange;
                    leadSourceColumn.NSeries[i].Area.Formatting = FormattingType.Custom;

                }
                else
                {
                    List<double> sumLastYear = new List<double>();
                    List<ReportDetailtServiceType> dataItemLastYear = listReportData.Where(x => x.Year == (year - 1).ToString()).ToList();

                    sumLastYear = dataItemLastYear.Select(x => x.TongDS).ToList();

                    totalRowData = string.Concat("{", string.Join(", ", sumLastYear), "}");
                    leadSourceColumn.NSeries.Add(totalRowData, true);

                    categoryData = string.Concat("{", string.Join(", ", listPartner), "}");
                    leadSourceColumn.NSeries.CategoryData = categoryData;

                    leadSourceColumn.NSeries[i].Name = string.Format("Tổng {0} {1}", textValue, year - 1);

                    leadSourceColumn.NSeries[i].Area.ForegroundColor = Color.Blue;
                    leadSourceColumn.NSeries[i].Area.Formatting = FormattingType.Custom;
                }

                i++;
            }
            
            // Set plot area formatting as none and hide its border.
            leadSourceColumn.PlotArea.Area.FillFormat.FillType = FillType.None;
            leadSourceColumn.PlotArea.Border.IsVisible = false;

            // Set value axis major tick mark as none and hide axis line. 
            // Also set the color of value axis major grid lines.
            leadSourceColumn.ValueAxis.MajorTickMark = TickMarkType.None;
            //leadSourceColumn.ValueAxis.MaxValue = 100;
            leadSourceColumn.ValueAxis.AxisLine.IsVisible = false;
            leadSourceColumn.ValueAxis.MajorGridLines.Color = Color.FromArgb(217, 217, 217);

            // Phần 2 Theo tỉ trọng doanh số theo thị trường
            title = "2. Theo tỷ trọng doanh số theo thị trường - loại hình chi trả";
            // Số 2 khoảng cách giữa phần 1 và phần 2
            CreateTitle(string.Format("B{0}", totalRowTable2 + 2), string.Format("E{0}", totalRowTable2 + 2), sheetReport, title, 12);

            // Tạo cột hearder cho table3
            title = "STT";
            CreateTitle(string.Format("B{0}", totalRowTable2 + 25), string.Format("B{0}", totalRowTable2 + 25), sheetReport, title, 12, true);

            title = "Đối tác";
            if (marketID.Contains("005"))
            {
                title = "Thị trường";
            }
            CreateTitle(string.Format("C{0}", totalRowTable2 + 25), string.Format("C{0}", totalRowTable2 + 25), sheetReport, title, 12, true);

            title = string.Format("Lũy kế {0} {1}", textValue, year);
            CreateTitle(string.Format("D{0}", totalRowTable2 + 25), string.Format("D{0}", totalRowTable2 + 25), sheetReport, title, 12, true);
            
            title = string.Format("Lũy kế {0} {1}", textValue, year - 1);
            CreateTitle(string.Format("E{0}", totalRowTable2 + 25), string.Format("E{0}", totalRowTable2 + 25), sheetReport, title, 12, true);

            // Tạo bảng cho tỉ trọng theo phần trăm của Loại hình dịch vụ
            List<ReportDetailtServiceType> listDataGradation = new ReportBL().ReportDetailtGradationCompareForOnePercent(year, int.Parse(gradationID), reportTypeID, marketID);

            if (marketID.Contains("005"))
            {
                List<ReportDetailtServiceType> listDataGradationConvert = new List<ReportDetailtServiceType>();
                listMarket = new List<string>();
                foreach (ReportDetailtServiceType item in listDataGradation)
                {
                    if (!listMarket.Contains(item.MarketName))
                    {
                        listMarket.Add(item.MarketName);
                    }
                }

                foreach (string item in listMarket)
                {
                    List<ReportDetailtServiceType> listDataItemYear = listDataGradation.Where(x => x.MarketName == item && x.Year == year.ToString()).ToList();
                    List<ReportDetailtServiceType> listDataItemLastYear = listDataGradation.Where(x => x.MarketName == item && x.Year == (year - 1).ToString()).ToList();

                    // Year
                    listDataGradationConvert.Add(
                        new ReportDetailtServiceType()
                        {
                            MarketName = item,
                            PartnerName = item,
                            DSChiQuay = listDataItemYear.Sum(x => x.DSChiQuay),
                            DSChiNha = listDataItemYear.Sum(x => x.DSChiNha),
                            DSCK = listDataItemYear.Sum(x => x.DSCK),
                            TongDS = listDataItemYear.Sum(x => x.TongDS),
                            Year = year.ToString()
                        }
                    );

                    // Last Year
                    listDataGradationConvert.Add(
                        new ReportDetailtServiceType()
                        {
                            MarketName = item,
                            PartnerName = item,
                            DSChiQuay = listDataItemLastYear.Sum(x => x.DSChiQuay),
                            DSChiNha = listDataItemLastYear.Sum(x => x.DSChiNha),
                            DSCK = listDataItemLastYear.Sum(x => x.DSCK),
                            TongDS = listDataItemLastYear.Sum(x => x.TongDS),
                            Year = (year - 1).ToString()
                        }
                    );
                }

                if (listDataGradationConvert.Count > 0)
                {
                    listDataGradation = new List<ReportDetailtServiceType>(listDataGradationConvert);
                }
            }


            double sumTongDSYear = listDataGradation.Where(x => x.Year == year.ToString()).Sum(x => x.TongDS);
            double sumTongDSLastYear = listDataGradation.Where(x => x.Year == (year - 1).ToString()).Sum(x => x.TongDS);

            List<ReportDetailtServiceType> listDataConvert = new List<ReportDetailtServiceType>();

            // Khởi tạo datatable
            table = new DataTable();
            // Tạo các cột cho datatable
            table.Columns.Add("STT", typeof(String));
            table.Columns.Add("PartnerName", typeof(String));
            table.Columns.Add("LK1", typeof(double));
            table.Columns.Add("LK2", typeof(double));

            int count = 1;
            foreach (ReportDetailtServiceType item in listDataGradation)
            {
                // Get dữ liệu của năm hiện tại
                ReportDetailtServiceType listDataYear = listDataGradation.Find(x => x.Year == year.ToString() && x.PartnerName == item.PartnerName);
                // Get dữ liệu của năm trước
                ReportDetailtServiceType listDataLastYear = listDataGradation.Find(x => x.Year == (year - 1).ToString() && x.PartnerName == item.PartnerName);

                // Check tồn tại của item
                string value = string.Format("PartnerName='{0}'", item.PartnerName);
                DataRow[] foundRows = table.Select(value);

                if (listDataYear != null && listDataLastYear != null && foundRows.Count() == 0)
                {
                    double valueYear = Math.Round(listDataYear.TongDS, 2, MidpointRounding.ToEven);
                    double valueLastYear = Math.Round(listDataLastYear.TongDS, 2, MidpointRounding.ToEven);
                    table.Rows.Add(count, listDataYear.PartnerName, valueYear, valueLastYear);
                }

                count++;
            }

            row = table.NewRow();
            row["STT"] = "";
            row["PartnerName"] = "Tổng";
            row["LK1"] = 100;
            row["LK2"] = 100;
            table.Rows.Add(row);

            // Với 6 là số cách của table1 và table2
            int totalRowTable3 = totalRowTable2 + table.Rows.Count + 25;

            // Table dữ liệu bảng số liệu Doanh số Chi Quầy/Chi Nhà/Chuyển Khoản
            if (table.Rows.Count > 0)
            {
                int stepRow = 0;
                // total row = row start + số row hiện có
                int totalRow = totalRowTable3;
                int rowStart = totalRowTable2 + 25;
                // Số dòng của row
                for (int a = rowStart; a < totalRow; a++)
                {
                    int stepColumn = 0;
                    // Số cột trong báo cáo cần hiển thị
                    // Tổng số cột hiển thị = Số cột hiển thị bắt đầu + tổng số cột cần hiển thị
                    int totalCol = 1 + 4;
                    for (int b = 1; b < totalCol; b++)
                    {
                        // Giá trị của value trong table
                        string valueOfTable = table.Rows[stepRow][stepColumn].ToString();
                        

                        // Insert vào dòng cột xác định trong Excel
                        sheetReport.Cells[a, b].PutValue(valueOfTable, true);

                        // set style cho number
                        if (b == 1)
                        {
                            style.Custom = "";
                            style.HorizontalAlignment = TextAlignmentType.Center;
                        }
                        else
                        {
                            style.Custom = "#,##0.00";
                            style.HorizontalAlignment = TextAlignmentType.General;
                        }
                        // set border
                        sheetReport.Cells[a, b].SetStyle(style);
                        

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

            // Vẻ biểu đồ tròn
            // Year
            List<ReportDetailtServiceType> dataPieYear = listDataGradation.Where(x => x.Year == year.ToString()).ToList();
            // Last Year
            List<ReportDetailtServiceType> dataPieLastYear = listDataGradation.Where(x => x.Year == (year - 1).ToString()).ToList();

            if (dataPieYear.Count > 0)
            {
                //Add Pie Chart
                Aspose.Cells.Charts.Chart leadSourcePie;
                chartIndex = sheetReport.Charts.Add(ChartType.Pie3D, totalRowTable2 + 4, 1, totalRowTable2 + 22, 6);
                leadSourcePie = sheetReport.Charts[chartIndex];

                // Set some properties of chart plot area.
                // To set the fill color and make the border invisible.
                leadSourcePie.PlotArea.Border.IsVisible = false;
                leadSourcePie.Elevation = 45;
                // Set properties of chart title
                leadSourcePie.Title.Text = string.Format("Tỉ trọng các đối tác theo thị trường năm {0}", year);
                leadSourcePie.Title.Font.Color = Color.Silver;
                leadSourcePie.Title.Font.IsBold = true;
                leadSourcePie.Title.Font.Size = 12;

                // Set properties of nseries
                // Danh sách tỉ trọng của các đối tác
                List<double> sumYear = new List<double>();
                sumYear = dataPieYear.Select(x => x.TongDS).ToList();

                // List đối tác/Thị trường
                listPartner = new List<string>();
                listPartner = dataPieYear.Select(x => x.PartnerName).ToList();

                totalRowData = string.Concat("{", string.Join(", ", sumYear), "}");
                leadSourcePie.NSeries.Add(totalRowData, true);

                categoryData = string.Concat("{", string.Join(", ", listPartner), "}");
                leadSourcePie.NSeries.CategoryData = categoryData;

                leadSourcePie.NSeries.IsColorVaried = true;

                // Set the DataLabels in the chart
                Aspose.Cells.Charts.DataLabels datalabels;
                for (i = 0; i < leadSourcePie.NSeries.Count; i++)
                {
                    datalabels = leadSourcePie.NSeries[i].DataLabels;
                    datalabels.Position = Aspose.Cells.Charts.LabelPositionType.InsideBase;
                    datalabels.ShowCategoryName = true;
                    //datalabels.ShowValue = true;
                    datalabels.ShowLegendKey = false;
                    datalabels.NumberFormat = "0.00%";
                    datalabels.ShowPercentage = true;
                }
            }

            // LastYear
            if (dataPieLastYear.Count > 0)
            {
                //Add Pie Chart
                Aspose.Cells.Charts.Chart leadSourcePie;
                chartIndex = sheetReport.Charts.Add(ChartType.Pie3D, totalRowTable2 + 4, 7, totalRowTable2 + 22, 13);
                leadSourcePie = sheetReport.Charts[chartIndex];

                // Set some properties of chart plot area.
                // To set the fill color and make the border invisible.
                leadSourcePie.PlotArea.Border.IsVisible = false;
                leadSourcePie.Elevation = 45;
                // Set properties of chart title
                leadSourcePie.Title.Text = string.Format("Tỉ trọng các đối tác theo thị trường năm {0}", year - 1);
                leadSourcePie.Title.Font.Color = Color.Silver;
                leadSourcePie.Title.Font.IsBold = true;
                leadSourcePie.Title.Font.Size = 12;

                // Set properties of nseries
                // Danh sách tỉ trọng của các đối tác
                List<double> sumYear = new List<double>();
                sumYear = dataPieLastYear.Select(x => x.TongDS).ToList();

                // List đối tác/Thị trường
                listPartner = new List<string>();
                listPartner = dataPieLastYear.Select(x => x.PartnerName).ToList();

                totalRowData = string.Concat("{", string.Join(", ", sumYear), "}");
                leadSourcePie.NSeries.Add(totalRowData, true);

                categoryData = string.Concat("{", string.Join(", ", listPartner), "}");
                leadSourcePie.NSeries.CategoryData = categoryData;

                leadSourcePie.NSeries.IsColorVaried = true;
                
                // Set the DataLabels in the chart
                Aspose.Cells.Charts.DataLabels datalabels;
                for (i = 0; i < leadSourcePie.NSeries.Count; i++)
                {
                    datalabels = leadSourcePie.NSeries[i].DataLabels;
                    datalabels.Position = Aspose.Cells.Charts.LabelPositionType.InsideBase;
                    datalabels.ShowCategoryName = true;
                    //datalabels.ShowValue = true;
                    datalabels.ShowLegendKey = false;
                    datalabels.NumberFormat = "0.00%";
                    datalabels.ShowPercentage = true;
                }
            }



            // Chạy process
            designer.Process();
            return ExportReport("ReportGradationCompare", designer);
        }

        /// <summary>
        /// Tạo mẫu cho Excel cho so sánh theo giai đoạn
        /// </summary>
        /// <param name="gradationID"></param>
        /// <param name="year"></param>
        /// <param name="typeID"></param>
        /// <returns></returns>
        public ActionResult CreateExcelForCompareForMonth(int year, int month, string reportTypeID, string marketID)
        {
            WorkbookDesigner designer = new WorkbookDesigner();
            string templatePath = "~/Content/Report/ReportDetailtCompareMonthByMakets.xlsx";
            // Get đường dẫn
            templatePath = System.Web.HttpContext.Current.Server.MapPath(templatePath);

            designer.Workbook = new DongA.Core.DongAExcel().OpenExcelFile(templatePath);
            designer.Workbook.CalculateFormula();

            WorksheetCollection workSheets = designer.Workbook.Worksheets;
            Worksheet sheetReport = designer.Workbook.Worksheets[0];

            // Tạo title
            string typeReport = "So sánh - Theo tháng - Tất cả";
            // Tạo title
            CreateTitle("A2", "K2", sheetReport, typeReport, 14);

            // Tạo title detailt
            string titleDetailt = string.Format("Tháng {0}/{1}", month, year);
            CreateTitle("A3", "K3", sheetReport, titleDetailt, 12);


            // Tạo giá trị cho cột dữ liệu của Chi quầy/ Chi nhà/ Chuyển khoản
            sheetReport.Cells["C62"].PutValue(string.Format("Tháng {0}/{1} ", month, year));
            sheetReport.Cells["D62"].PutValue(string.Format("Tháng {0}/{1} ", month - 1, year));
            sheetReport.Cells["E62"].PutValue(string.Format("Tháng {0}/{1} ", month, year - 1));
            
            sheetReport.Cells["F62"].PutValue(string.Format("Tháng {0}/{1} ", month, year));
            sheetReport.Cells["G62"].PutValue(string.Format("Tháng {0}/{1} ", month - 1, year));
            sheetReport.Cells["H62"].PutValue(string.Format("Tháng {0}/{1} ", month, year - 1));

            sheetReport.Cells["I62"].PutValue(string.Format("Tháng {0}/{1} ", month, year));
            sheetReport.Cells["J62"].PutValue(string.Format("Tháng {0}/{1} ", month - 1, year));
            sheetReport.Cells["K62"].PutValue(string.Format("Tháng {0}/{1} ", month, year - 1));

            sheetReport.Cells["L62"].PutValue(string.Format("Tháng {0}/{1} ", month, year));
            sheetReport.Cells["M62"].PutValue(string.Format("Tháng {0}/{1} ", month - 1, year));
            sheetReport.Cells["N62"].PutValue(string.Format("Tháng {0}/{1} ", month, year - 1));

            // Tạo table
            DataTable table = new DataTable();
            // Khởi tạo datatable
            table = new DataTable();
            // Tạo các cột cho datatable
            // tháng hiện tại

            table.Columns.Add("ReportID", typeof(String));

            table.Columns.Add("CQ1", typeof(double));
            table.Columns.Add("CQ2", typeof(double));
            table.Columns.Add("CQ3", typeof(double));

            table.Columns.Add("CN1", typeof(double));
            table.Columns.Add("CN2", typeof(double));
            table.Columns.Add("CN3", typeof(double));

            table.Columns.Add("CK1", typeof(double));
            table.Columns.Add("CK2", typeof(double));
            table.Columns.Add("CK3", typeof(double));

            table.Columns.Add("TDS1", typeof(double));
            table.Columns.Add("TDS2", typeof(double));
            table.Columns.Add("TDS3", typeof(double));

            table.Columns.Add("MarketName", typeof(String));
            
            List<ReportDetailtSTMarket> listDataCompareMonth = new ReportBL().ReportDetailtCompareMonthForAll(year, month, reportTypeID, marketID);

            // Trường hợp với tất cả thị trường
            if (marketID.Equals("0"))
            {
                foreach (ReportDetailtSTMarket item in listDataCompareMonth)
                {
                    item.ReportID = item.MarketName;
                    item.TongDS = item.DSChiQuay + item.DSChiNha + item.DSCK;
                    item.MarketName = "Tất cả thị trường";
                }

                // Danh sách mã thị trường của Tất cả
                List<string> listMarket = new List<string>() { "003", "005", "001", "002", "014", "004" };

                foreach (string item in listMarket)
                {
                    // Cùng kì
                    ReportDetailtSTMarket dataItemLastYear = listDataCompareMonth.Find(x => x.MarketCode == item && x.Month == month.ToString() && x.Year == (year - 1).ToString());
                    ReportDetailtSTMarket dataItemYear = listDataCompareMonth.Find(x => x.MarketCode == item && x.Month == month.ToString() && x.Year == year.ToString());
                    ReportDetailtSTMarket dataItemLastMonth = listDataCompareMonth.Find(x => x.MarketCode == item && x.Month == (month - 1).ToString() && x.Year == year.ToString());
                    // Trường hợp tháng 1
                    if (month == 1)
                    {
                        dataItemLastMonth = listDataCompareMonth.Find(x => x.MarketCode == item && x.Month == "12" && x.Year == (year - 1).ToString());
                    }

                    if (dataItemLastYear != null && dataItemYear != null && dataItemLastMonth != null)
                    {
                        // add item vào table
                        table.Rows.Add(dataItemYear.ReportID
                            , dataItemYear.DSChiQuay, dataItemLastMonth.DSChiQuay, dataItemLastYear.DSChiQuay
                            , dataItemYear.DSChiNha, dataItemLastMonth.DSChiNha, dataItemLastYear.DSChiNha
                            , dataItemYear.DSCK, dataItemLastMonth.DSCK, dataItemLastYear.DSCK
                            , dataItemYear.TongDS, dataItemLastMonth.TongDS, dataItemLastYear.TongDS
                            , dataItemYear.MarketName);
                    }
                }

            }
            else
            {
                foreach (ReportDetailtSTMarket item in listDataCompareMonth)
                {
                    item.ReportID = item.PartnerName;
                    item.TongDS = item.DSChiQuay + item.DSChiNha + item.DSCK;
                }

                foreach (ReportDetailtSTMarket item in listDataCompareMonth)
                {
                    // Cùng kì
                    ReportDetailtSTMarket dataItemLastYear = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == month.ToString() && x.Year == (year - 1).ToString());
                    ReportDetailtSTMarket dataItemYear = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == month.ToString() && x.Year == year.ToString());
                    ReportDetailtSTMarket dataItemLastMonth = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == (month - 1).ToString() && x.Year == year.ToString());
                    // Trường hợp tháng 1
                    if (month == 1)
                    {
                        dataItemLastMonth = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == "12" && x.Year == (year - 1).ToString());
                    }

                    // Cùng kì năm trước
                    if (dataItemLastYear == null)
                    {
                        dataItemLastYear = new ReportDetailtSTMarket();
                    }

                    // Hiện tại
                    if (dataItemYear == null)
                    {
                        dataItemYear = new ReportDetailtSTMarket();
                    }

                    // tháng trước
                    if (dataItemYear == null)
                    {
                        dataItemLastMonth = new ReportDetailtSTMarket();

                    }

                    // Check tồn tại của item
                    string value = string.Format("ReportID='{0}'", item.ReportID);

                    DataRow[] foundRows = table.Select(value);

                    if (dataItemLastYear != null && dataItemYear != null && dataItemLastMonth != null && foundRows.Count() == 0)
                    {
                        // add item vào table
                        table.Rows.Add(item.ReportID
                            , dataItemYear.DSChiQuay, dataItemLastMonth.DSChiQuay, dataItemLastYear.DSChiQuay
                            , dataItemYear.DSChiNha, dataItemLastMonth.DSChiNha, dataItemLastYear.DSChiNha
                            , dataItemYear.DSCK, dataItemLastMonth.DSCK, dataItemLastYear.DSCK
                            , dataItemYear.TongDS, dataItemLastMonth.TongDS, dataItemLastYear.TongDS
                            , item.MarketName);
                    }

                }
            }


            DataRow row = table.NewRow();
            row["ReportID"] = "Tổng";
            row["CQ1"] = table.Compute("Sum(CQ1)", "");
            row["CQ2"] = table.Compute("Sum(CQ2)", "");
            row["CQ3"] = table.Compute("Sum(CQ3)", "");

            row["CN1"] = table.Compute("Sum(CN1)", "");
            row["CN2"] = table.Compute("Sum(CN2)", "");
            row["CN3"] = table.Compute("Sum(CN3)", "");

            row["CK1"] = table.Compute("Sum(CK1)", "");
            row["CK2"] = table.Compute("Sum(CK2)", "");
            row["CK3"] = table.Compute("Sum(CK3)", "");

            row["TDS1"] = table.Compute("Sum(TDS1)", "");
            row["TDS2"] = table.Compute("Sum(TDS2)", "");
            row["TDS3"] = table.Compute("Sum(TDS3)", "");

            row["MarketName"] = "";
            table.Rows.Add(row);


            // Tổng số row theo table1
            int totalRowTable1 = table.Rows.Count + 62;

            // Set border
            Style style = new CellsFactory().CreateStyle();
            style.SetBorder(BorderType.BottomBorder, CellBorderType.Thin, Color.Black);
            style.SetBorder(BorderType.LeftBorder, CellBorderType.Thin, Color.Black);
            style.SetBorder(BorderType.RightBorder, CellBorderType.Thin, Color.Black);

            if (table.Rows.Count > 0)
            {
                int stepRow = 0;
                // total row = row start + số row hiện có
                int totalRow = table.Rows.Count + 62;
                // Số dòng của row
                for (int a = 62; a < totalRow; a++)
                {
                    int stepColumn = 0;
                    // Số cột trong báo cáo cần hiển thị
                    // Tổng số cột hiển thị = Số cột hiển thị bắt đầu + tổng số cột cần hiển thị
                    int totalCol = 1 + 13;
                    for (int b = 1; b < totalCol; b++)
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

                        // Cột tổng cộng
                        if (b.Equals(totalCol - 2))
                        {
                            sheetReport.Cells[a, b].PutValue(valueOfTable, true, true);
                            style.Font.IsBold = true;
                            sheetReport.Cells[a, b].SetStyle(style);
                        }

                        // Cột tổng cộng
                        if (b.Equals(totalCol - 3))
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

            // Tạo table cho việc tăng giảm

            // Khởi tạo datatable
            table = new DataTable();
            // Tạo các cột cho datatable

            table.Columns.Add("ReportID", typeof(String));
            // So sánh với tháng trước
            table.Columns.Add("CQ1", typeof(double));
            table.Columns.Add("CQ2", typeof(double));

            table.Columns.Add("CN1", typeof(double));
            table.Columns.Add("CN2", typeof(double));

            table.Columns.Add("CK1", typeof(double));
            table.Columns.Add("CK2", typeof(double));

            table.Columns.Add("TDS1", typeof(double));
            table.Columns.Add("TDS2", typeof(double));

            table.Columns.Add("MarketName", typeof(String));


            // trường hợp tất cả thị trường
            if (marketID.Equals("0"))
            {

                foreach (ReportDetailtSTMarket item in listDataCompareMonth)
                {
                    // Cùng kì
                    ReportDetailtSTMarket dataItemLastYear = listDataCompareMonth.Find(x => x.MarketCode == item.MarketCode && x.Month == month.ToString() && x.Year == (year - 1).ToString());
                    ReportDetailtSTMarket dataItemMonth = listDataCompareMonth.Find(x => x.MarketCode == item.MarketCode && x.Month == month.ToString() && x.Year == year.ToString());
                    ReportDetailtSTMarket dataItemLastMonth = listDataCompareMonth.Find(x => x.MarketCode == item.MarketCode && x.Month == (month - 1).ToString() && x.Year == year.ToString());
                    // Trường hợp tháng 1
                    if (month == 1)
                    {
                        dataItemLastMonth = listDataCompareMonth.Find(x => x.MarketCode == item.MarketCode && x.Month == "12" && x.Year == (year - 1).ToString());
                    }

                    // Check tồn tại của item
                    string value = string.Format("ReportID='{0}'", item.ReportID);

                    DataRow[] foundRows = table.Select(value);

                    if (dataItemLastYear != null && dataItemMonth != null && dataItemLastMonth != null && foundRows.Count() == 0)
                    {
                        // add item vào table
                        table.Rows.Add(item.ReportID
                            , dataItemMonth.DSChiQuay - dataItemLastMonth.DSChiQuay, dataItemMonth.DSChiQuay - dataItemLastYear.DSChiQuay
                            , dataItemMonth.DSChiNha - dataItemLastMonth.DSChiNha, dataItemMonth.DSChiNha - dataItemLastYear.DSChiNha
                            , dataItemMonth.DSCK - dataItemLastMonth.DSCK, dataItemMonth.DSCK - dataItemLastYear.DSCK
                            , dataItemMonth.TongDS - dataItemLastMonth.TongDS, dataItemMonth.TongDS - dataItemLastYear.TongDS
                            , item.MarketName);
                    }
                }
            }
            else
            {
                // Trường hợp các thị trường con của thị trường châu Á
                foreach (ReportDetailtSTMarket item in listDataCompareMonth)
                {
                    item.ReportID = item.PartnerName;
                    item.TongDS = item.DSChiQuay + item.DSChiNha + item.DSCK;
                }

                foreach (ReportDetailtSTMarket item in listDataCompareMonth)
                {
                    // Cùng kì
                    ReportDetailtSTMarket dataItemLastYear = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == month.ToString() && x.Year == (year - 1).ToString());
                    ReportDetailtSTMarket dataItemMonth = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == month.ToString() && x.Year == year.ToString());
                    ReportDetailtSTMarket dataItemLastMonth = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == (month - 1).ToString() && x.Year == year.ToString());
                    // Trường hợp tháng 1
                    if (month == 1)
                    {
                        dataItemLastMonth = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == "12" && x.Year == (year - 1).ToString());
                    }

                    // Cùng kì năm trước
                    if (dataItemLastYear == null)
                    {
                        dataItemLastYear = new ReportDetailtSTMarket();
                    }

                    // Hiện tại
                    if (dataItemMonth == null)
                    {
                        dataItemMonth = new ReportDetailtSTMarket();
                    }

                    // tháng trước
                    {
                        dataItemLastMonth = new ReportDetailtSTMarket();

                    }

                    // Check tồn tại của item
                    string value = string.Format("ReportID='{0}'", item.ReportID);

                    DataRow[] foundRows = table.Select(value);

                    if (dataItemLastYear != null && dataItemMonth != null && dataItemLastMonth != null && foundRows.Count() == 0)
                    {
                        // add item vào table
                        table.Rows.Add(item.ReportID
                            , dataItemMonth.DSChiQuay - dataItemLastMonth.DSChiQuay, dataItemMonth.DSChiQuay - dataItemLastYear.DSChiQuay
                            , dataItemMonth.DSChiNha - dataItemLastMonth.DSChiNha, dataItemMonth.DSChiNha - dataItemLastYear.DSChiNha
                            , dataItemMonth.DSCK - dataItemLastMonth.DSCK, dataItemMonth.DSCK - dataItemLastYear.DSCK
                            , dataItemMonth.TongDS - dataItemLastMonth.TongDS, dataItemMonth.TongDS - dataItemLastYear.TongDS
                            , item.MarketName);
                    }

                }
            }

            row = table.NewRow();
            row["ReportID"] = "Tổng";
            row["CQ1"] = table.Compute("Sum(CQ1)", "");
            row["CQ2"] = table.Compute("Sum(CQ2)", "");

            row["CN1"] = table.Compute("Sum(CN1)", "");
            row["CN2"] = table.Compute("Sum(CN2)", "");
            
            row["CK1"] = table.Compute("Sum(CK1)", "");
            row["CK2"] = table.Compute("Sum(CK2)", "");

            row["TDS1"] = table.Compute("Sum(TDS1)", "");
            row["TDS2"] = table.Compute("Sum(TDS2)", "");

            row["MarketName"] = "";
            table.Rows.Add(row);

            // Tổng số row của table2
            // Với 6 là số cách của table1 và table2
            int totalRowTable2 = totalRowTable1 + table.Rows.Count + 6;

            // Tạo title hearder cho table tăng giảm
            // Title cho thị trường
            string title = "Thị trường";

            CreateTitle(string.Format("B{0}", totalRowTable1 + 6 - 1), string.Format("B{0}", totalRowTable1 + 6), sheetReport, title, 12, true);

            title = "Chi Quầy";
            CreateTitle(string.Format("C{0}", totalRowTable1 + 6 - 1), string.Format("D{0}", totalRowTable1 + 6 - 1), sheetReport, title, 12, true);

            title = "Chi Nhà";
            CreateTitle(string.Format("E{0}", totalRowTable1 + 6 - 1), string.Format("F{0}", totalRowTable1 + 6 - 1), sheetReport, title, 12, true);

            title = "Chuyển Khoản";
            CreateTitle(string.Format("G{0}", totalRowTable1 + 6 - 1), string.Format("H{0}", totalRowTable1 + 6 - 1), sheetReport, title, 12, true);

            title = "Tổng";
            CreateTitle(string.Format("I{0}", totalRowTable1 + 6 - 1), string.Format("J{0}", totalRowTable1 + 6 - 1), sheetReport, title, 12, true);


            title = "So với tháng trước";
            CreateTitle(string.Format("C{0}", totalRowTable1 + 6), string.Format("C{0}", totalRowTable1 + 6), sheetReport, title, 12, true);

            title = string.Format("So với cùng kì");
            CreateTitle(string.Format("D{0}", totalRowTable1 + 6), string.Format("D{0}", totalRowTable1 + 6), sheetReport, title, 12, true);

            title = "So với tháng trước";
            CreateTitle(string.Format("E{0}", totalRowTable1 + 6), string.Format("E{0}", totalRowTable1 + 6), sheetReport, title, 12, true);

            title = string.Format("So với cùng kì");
            CreateTitle(string.Format("F{0}", totalRowTable1 + 6), string.Format("F{0}", totalRowTable1 + 6), sheetReport, title, 12, true);

            title = "So với tháng trước";
            CreateTitle(string.Format("G{0}", totalRowTable1 + 6), string.Format("G{0}", totalRowTable1 + 6), sheetReport, title, 12, true);

            title = string.Format("So với cùng kì");
            CreateTitle(string.Format("H{0}", totalRowTable1 + 6), string.Format("H{0}", totalRowTable1 + 6), sheetReport, title, 12, true);

            title = "So với tháng trước";
            CreateTitle(string.Format("I{0}", totalRowTable1 + 6), string.Format("I{0}", totalRowTable1 + 6), sheetReport, title, 12, true);

            title = string.Format("So với cùng kì");
            CreateTitle(string.Format("J{0}", totalRowTable1 + 6), string.Format("J{0}", totalRowTable1 + 6), sheetReport, title, 12, true);

            // Table dữ liệu bảng số liệu Doanh số Chi Quầy/Chi Nhà/Chuyển Khoản
            if (table.Rows.Count > 0)
            {
                int stepRow = 0;
                // total row = row start + số row hiện có
                int totalRow = totalRowTable2;
                int rowStart = totalRowTable1 + 6;
                // Số dòng của row
                for (int a = rowStart; a < totalRow; a++)
                {
                    int stepColumn = 0;
                    // Số cột trong báo cáo cần hiển thị
                    // Tổng số cột hiển thị = Số cột hiển thị bắt đầu + tổng số cột cần hiển thị
                    int totalCol = 1 + 9;
                    for (int b = 1; b < totalCol; b++)
                    {
                        // Giá trị của value trong table
                        string valueOfTable = table.Rows[stepRow][stepColumn].ToString();
                        
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

            //// Vẽ biểu đồ cột cho Excel
            ////Chart reference
            //Aspose.Cells.Charts.Chart leadSourceColumnChiNha;
            ////Add Pie Chart
            //// Chi Nhà
            //int chartIndex = sheetReport.Charts.Add(ChartType.Bar, 25, 0, 40, 7);
            //leadSourceColumnChiNha = sheetReport.Charts[chartIndex];

            ////Chart title
            //leadSourceColumnChiNha.Title.Text = string.Format("Tỉ trọng từng dịch vụ từng thị trường \n Tháng {0}/{1}", month, year);
            //leadSourceColumnChiNha.Title.Font.Color = Color.Silver;

            //// Adding SeriesCollection (chart data source) to the chart ranging from "A1" cell to "B3"
            //string totalRowData = string.Format("C63:D{0}", 63 + table.Rows.Count - 2);
            //leadSourceColumnChiNha.NSeries.Add(totalRowData, true);

            //// Set the category data covering the range A2:A5.
            //// Tổng số dòng cần hiển thị là số dòng hiện tại  + số dòng trong listData -2 (trong đó 1 là dòng tổng cuối cùng)
            //string categoryData = "B63:B68";
            //leadSourceColumnChiNha.NSeries.CategoryData = categoryData;

            //// Set the names of the chart series taken from cells.
            //leadSourceColumnChiNha.NSeries[0].Name = "=C62";
            //leadSourceColumnChiNha.NSeries[1].Name = "=D62";

            //// Set the 1st series fill color.
            //leadSourceColumnChiNha.NSeries[0].Area.ForegroundColor = Color.Orange;
            //leadSourceColumnChiNha.NSeries[0].Area.Formatting = FormattingType.Custom;

            //// Set the 2nd series fill color.
            //leadSourceColumnChiNha.NSeries[1].Area.ForegroundColor = Color.Green;
            //leadSourceColumnChiNha.NSeries[1].Area.Formatting = FormattingType.Custom;


            //// Set plot area formatting as none and hide its border.
            //leadSourceColumnChiNha.PlotArea.Area.FillFormat.FillType = FillType.None;
            //leadSourceColumnChiNha.PlotArea.Border.IsVisible = false;

            //// Set value axis major tick mark as none and hide axis line. 
            //// Also set the color of value axis major grid lines.
            //leadSourceColumnChiNha.ValueAxis.MajorTickMark = TickMarkType.None;
            //leadSourceColumnChiNha.ValueAxis.AxisLine.IsVisible = false;
            ////leadSourceColumnChiNha.ValueAxis.IsAutomaticMajorUnit = false;
            ////leadSourceColumnChiNha.ValueAxis.MajorUnit = 10000000;
            //leadSourceColumnChiNha.ValueAxis.MajorGridLines.Color = Color.FromArgb(217, 217, 217);

            ////Chart reference
            //Aspose.Cells.Charts.Chart leadSourceColumnChiQuay;
            //// Chi Quầy
            //chartIndex = sheetReport.Charts.Add(ChartType.Bar, 25, 8, 40, 15);
            //leadSourceColumnChiQuay = sheetReport.Charts[chartIndex];

            ////Chart title
            //leadSourceColumnChiQuay.Title.Text = string.Format("Doanh số dịch vụ chi quầy từng thị trường \n Giai đoạn: {0} \n Chi Quầy", text);
            //leadSourceColumnChiQuay.Title.Font.Color = Color.Silver;

            //// Adding SeriesCollection (chart data source) to the chart ranging from "A1" cell to "B3"
            //totalRowData = string.Format("E63:F{0}", 63 + table.Rows.Count - 2);
            //leadSourceColumnChiQuay.NSeries.Add(totalRowData, true);

            //// Set the category data covering the range A2:A5.
            //// Tổng số dòng cần hiển thị là số dòng hiện tại  + số dòng trong listData -2 (trong đó 1 là dòng tổng cuối cùng)
            //categoryData = "B63:B68";
            //leadSourceColumnChiQuay.NSeries.CategoryData = categoryData;

            //// Set the names of the chart series taken from cells.
            //leadSourceColumnChiQuay.NSeries[0].Name = "=C62";
            //leadSourceColumnChiQuay.NSeries[1].Name = "=D62";

            //// Set the 1st series fill color.
            //leadSourceColumnChiQuay.NSeries[0].Area.ForegroundColor = Color.Orange;
            //leadSourceColumnChiQuay.NSeries[0].Area.Formatting = FormattingType.Custom;

            //// Set the 2nd series fill color.
            //leadSourceColumnChiQuay.NSeries[1].Area.ForegroundColor = Color.Green;
            //leadSourceColumnChiQuay.NSeries[1].Area.Formatting = FormattingType.Custom;


            //// Set plot area formatting as none and hide its border.
            //leadSourceColumnChiQuay.PlotArea.Area.FillFormat.FillType = FillType.None;
            //leadSourceColumnChiQuay.PlotArea.Border.IsVisible = false;

            //// Set value axis major tick mark as none and hide axis line. 
            //// Also set the color of value axis major grid lines.
            //leadSourceColumnChiQuay.ValueAxis.MajorTickMark = TickMarkType.None;
            //leadSourceColumnChiQuay.ValueAxis.AxisLine.IsVisible = false;
            //leadSourceColumnChiQuay.ValueAxis.MajorGridLines.Color = Color.FromArgb(217, 217, 217);


            ////Chart reference
            //Aspose.Cells.Charts.Chart leadSourceColumnChuyenKhoan;
            //// Chuyển khoản
            //chartIndex = sheetReport.Charts.Add(ChartType.Bar, 41, 0, 56, 7);
            //leadSourceColumnChuyenKhoan = sheetReport.Charts[chartIndex];

            ////Chart title
            //leadSourceColumnChuyenKhoan.Title.Text = string.Format("Doanh số dịch vụ chi quầy từng thị trường \n Giai đoạn: {0} \n Chuyển khoản", text);
            //leadSourceColumnChuyenKhoan.Title.Font.Color = Color.Silver;

            //// Adding SeriesCollection (chart data source) to the chart ranging from "A1" cell to "B3"
            //totalRowData = string.Format("G63:H{0}", 63 + table.Rows.Count - 2);
            //leadSourceColumnChuyenKhoan.NSeries.Add(totalRowData, true);

            //// Set the category data covering the range A2:A5.
            //// Tổng số dòng cần hiển thị là số dòng hiện tại  + số dòng trong listData -2 (trong đó 1 là dòng tổng cuối cùng)
            //categoryData = "B63:B68";
            //leadSourceColumnChuyenKhoan.NSeries.CategoryData = categoryData;

            //// Set the names of the chart series taken from cells.
            //leadSourceColumnChuyenKhoan.NSeries[0].Name = "=C62";
            //leadSourceColumnChuyenKhoan.NSeries[1].Name = "=D62";

            //// Set the 1st series fill color.
            //leadSourceColumnChuyenKhoan.NSeries[0].Area.ForegroundColor = Color.Orange;
            //leadSourceColumnChuyenKhoan.NSeries[0].Area.Formatting = FormattingType.Custom;

            //// Set the 2nd series fill color.
            //leadSourceColumnChuyenKhoan.NSeries[1].Area.ForegroundColor = Color.Green;
            //leadSourceColumnChuyenKhoan.NSeries[1].Area.Formatting = FormattingType.Custom;


            //// Set plot area formatting as none and hide its border.
            //leadSourceColumnChuyenKhoan.PlotArea.Area.FillFormat.FillType = FillType.None;
            //leadSourceColumnChuyenKhoan.PlotArea.Border.IsVisible = false;

            //// Set value axis major tick mark as none and hide axis line. 
            //// Also set the color of value axis major grid lines.
            //leadSourceColumnChuyenKhoan.ValueAxis.MajorTickMark = TickMarkType.None;
            //leadSourceColumnChuyenKhoan.ValueAxis.AxisLine.IsVisible = false;
            //leadSourceColumnChuyenKhoan.ValueAxis.MajorGridLines.Color = Color.FromArgb(217, 217, 217);


            //// Tạo chart cột tỉ trọng cho các thị trường
            ////Add Pie Chart
            ////Chart reference
            //Aspose.Cells.Charts.Chart leadSourceColumn;
            //chartIndex = sheetReport.Charts.Add(ChartType.Column3DStacked, 6, 0, 24, 13);
            //leadSourceColumn = sheetReport.Charts[chartIndex];


            ////Chart title
            //leadSourceColumn.Title.Text = string.Format("Tỉ trọng từng thị trường \n Giai đoạn: {0} \n", text);
            //leadSourceColumn.Title.Font.Color = Color.Silver;

            //// count cho 3 loại: chi Quầy, Chi nhà, Chuyển khoản
            //// list thị trường

            //List<ReportDetailtServiceType> listDataGradation = new ReportBL().ReportDetailtGradationCompareForAllPercent(year, int.Parse(gradationID), reportTypeID, marketID);

            //List<string> listMarketCurrent = new List<string>();
            //foreach (ReportDetailtServiceType item in listDataGradation)
            //{
            //    if (!listMarketCurrent.Contains(item.MarketName))
            //    {
            //        listMarketCurrent.Add(item.MarketName);
            //    }
            //}

            //// List dữ liệu dataRow
            //string[] listTotalRowData = new string[listMarketCurrent.Count];
            //int i = 0;
            //foreach (string item in listMarketCurrent)
            //{
            //    // Năm trước
            //    List<ReportDetailtServiceType> dataItemLastYear = listDataGradation.Where(x => x.MarketName == item && x.Year == (year - 1).ToString()).ToList();
            //    List<ReportDetailtServiceType> dataItemYear = listDataGradation.Where(x => x.MarketName == item && x.Year == year.ToString()).ToList();


            //    listTotalRowData[i++] = string.Concat("{"
            //        , string.Format("{0}, {1}, {2}, {3}, {4}, {5}"
            //        , dataItemYear.Sum(x => x.DSChiQuay), dataItemLastYear.Sum(x => x.DSChiQuay)
            //        , dataItemYear.Sum(x => x.DSChiNha), dataItemLastYear.Sum(x => x.DSChiNha)
            //        , dataItemYear.Sum(x => x.DSCK), dataItemLastYear.Sum(x => x.DSCK))
            //        , "}");
            //}

            //int count = 0;
            //string yearCurent = string.Format("Năm {0}", year);
            //string lastYear = string.Format("Năm {0}", year - 1);
            //foreach (string item in listTotalRowData)
            //{
            //    totalRowData = item;
            //    leadSourceColumn.NSeries.Add(totalRowData, true);

            //    categoryData = string.Concat("{", string.Format("Chi Quầy {0}, Chi Quầy {1}, Chi Nhà {0}, Chi Nhà {1}, Chuyển Khoản {0},  Chuyển Khoản {1}", yearCurent, lastYear), "}");
            //    leadSourceColumn.NSeries.CategoryData = categoryData;

            //    leadSourceColumn.NSeries[count].Name = listMarketCurrent[count];

            //    switch (count)
            //    {
            //        case 0:
            //            leadSourceColumn.NSeries[count].Area.ForegroundColor = Color.Orange;
            //            leadSourceColumn.NSeries[count].Area.Formatting = FormattingType.Custom;
            //            break;
            //        case 1:
            //            leadSourceColumn.NSeries[count].Area.ForegroundColor = Color.Green;
            //            leadSourceColumn.NSeries[count].Area.Formatting = FormattingType.Custom;
            //            break;
            //        case 2:
            //            leadSourceColumn.NSeries[count].Area.ForegroundColor = Color.Blue;
            //            leadSourceColumn.NSeries[count].Area.Formatting = FormattingType.Custom;
            //            break;
            //        case 3:
            //            leadSourceColumn.NSeries[count].Area.ForegroundColor = Color.Yellow;
            //            leadSourceColumn.NSeries[count].Area.Formatting = FormattingType.Custom;
            //            break;
            //        case 4:
            //            leadSourceColumn.NSeries[count].Area.ForegroundColor = Color.Brown;
            //            leadSourceColumn.NSeries[count].Area.Formatting = FormattingType.Custom;
            //            break;
            //        case 5:
            //            leadSourceColumn.NSeries[count].Area.ForegroundColor = Color.Olive;
            //            leadSourceColumn.NSeries[count].Area.Formatting = FormattingType.Custom;
            //            break;
            //        default:
            //            leadSourceColumn.NSeries[count].Area.ForegroundColor = Color.Pink;
            //            leadSourceColumn.NSeries[count].Area.Formatting = FormattingType.Custom;
            //            break;
            //    }
            //    count++;
            //}

            //// Set plot area formatting as none and hide its border.
            //leadSourceColumn.PlotArea.Area.FillFormat.FillType = FillType.None;
            //leadSourceColumn.PlotArea.Border.IsVisible = false;

            //// Set value axis major tick mark as none and hide axis line. 
            //// Also set the color of value axis major grid lines.
            //leadSourceColumn.ValueAxis.MajorTickMark = TickMarkType.None;
            //leadSourceColumn.ValueAxis.MaxValue = 100;
            //leadSourceColumn.ValueAxis.AxisLine.IsVisible = false;
            //leadSourceColumn.ValueAxis.MajorGridLines.Color = Color.FromArgb(217, 217, 217);

            // Chạy process
            designer.Process();
            return ExportReport("ReportCompareForMonthForAll", designer);
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
        private DataTable CreateDataTableFormart()
        {
            DataTable db = new DataTable();

            db.Columns.Add("STT", typeof(string));
            db.Columns.Add("ReportID", typeof(string));
            db.Columns.Add("DSChiQuay", typeof(double));
            db.Columns.Add("DSChiNha", typeof(double));
            db.Columns.Add("DSCK", typeof(double));
            db.Columns.Add("TongDS", typeof(double));
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
        private void FillData(DataTable mother, DataTable fill)
        {
            int stt = 1;
            foreach (DataRow dr in mother.Rows)
            {
                var row = fill.NewRow();

                row["STT"] = string.IsNullOrEmpty(dr["STT"].ToString()) ? "" : (string)dr["STT"];
                row["ReportID"] = string.IsNullOrEmpty(dr["ReportID"].ToString()) ? "" : (string)dr["ReportID"];
                row["DSChiQuay"] = (double)dr["DSChiQuay"];
                row["DSChiNha"] = (double)dr["DSChiNha"];
                row["DSCK"] = (double)dr["DSCK"];
                row["TongDS"] = (double)dr["TongDS"];
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
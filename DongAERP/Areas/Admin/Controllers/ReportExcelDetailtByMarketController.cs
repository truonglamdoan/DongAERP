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

                    if (!string.IsNullOrEmpty(marketID))
                    {
                        if (marketID == "0")
                        {
                            foreach (ReportDetailtServiceType item in listReportData)
                            {
                                // Add thêm thị trường vào list
                                if (!listMarket.Contains(item.MarketName))
                                {
                                    listMarket.Add(item.MarketName);
                                }

                                item.ReportID = item.MarketName;
                                item.TongDS = item.DSChiNha + item.DSChiQuay + item.DSCK;
                            }
                        }
                        else
                        {
                            List<ReportDetailtServiceType> listDataConvert = new List<ReportDetailtServiceType>();

                            foreach (ReportDetailtServiceType item in listReportData)
                            {
                                // Add thêm thị trường vào list
                                if (!listMarket.Contains(item.MarketName))
                                {
                                    listMarket.Add(item.MarketName);
                                }

                                item.ReportID = item.PartnerName;
                                item.TongDS = item.DSChiNha + item.DSChiQuay + item.DSCK;
                            }

                            foreach (string item in listMarket)
                            {
                                List<ReportDetailtServiceType> listDataAsianChild = listReportData.Where(x => x.MarketName == item).ToList();

                                listDataConvert.Add(
                                    new ReportDetailtServiceType()
                                    {
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
                                listReportData = new List<ReportDetailtServiceType>(listDataConvert);
                            }
                        }
                    }
                    // Set from day and to day
                    sheetReport.Cells["E4"].PutValue(string.Format("{0}/{1}", fromDate.Month, fromDate.Year));
                    sheetReport.Cells["H4"].PutValue(string.Format("{0}/{1}", toDate.Month, toDate.Year));
                    break;
                // Theo năm
                default:

                    listReportData = new ReportBL().SearchMarketForOneForYear(fromDate, toDate, reportTypeID, marketID);

                    if (!string.IsNullOrEmpty(marketID))
                    {
                        if (marketID == "0")
                        {
                            foreach (ReportDetailtServiceType item in listReportData)
                            {
                                // Add thêm thị trường vào list
                                if (!listMarket.Contains(item.MarketName))
                                {
                                    listMarket.Add(item.MarketName);
                                }

                                item.ReportID = item.MarketName;
                                item.TongDS = item.DSChiNha + item.DSChiQuay + item.DSCK;
                            }
                        }
                        else
                        {
                            List<ReportDetailtServiceType> listDataConvert = new List<ReportDetailtServiceType>();

                            foreach (ReportDetailtServiceType item in listReportData)
                            {
                                // Add thêm thị trường vào list
                                if (!listMarket.Contains(item.MarketName))
                                {
                                    listMarket.Add(item.MarketName);
                                }

                                item.ReportID = item.PartnerName;
                                item.TongDS = item.DSChiNha + item.DSChiQuay + item.DSCK;
                            }

                            foreach (string item in listMarket)
                            {
                                List<ReportDetailtServiceType> listDataAsianChild = listReportData.Where(x => x.MarketName == item).ToList();

                                listDataConvert.Add(
                                    new ReportDetailtServiceType()
                                    {
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
                                listReportData = new List<ReportDetailtServiceType>(listDataConvert);
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
        /// Tạo mẫu cho Excel cho so sánh theo giai đoạn
        /// </summary>
        /// <param name="gradationID"></param>
        /// <param name="year"></param>
        /// <param name="typeID"></param>
        /// <returns></returns>
        public ActionResult CreateExcelForGradationCompare(string gradationID, int year, string reportTypeID)
        {
            WorkbookDesigner designer = new WorkbookDesigner();
            string templatePath = "~/Content/Report/ReportMaketForGradationCompare.xlsx";
            // Get đường dẫn
            templatePath = System.Web.HttpContext.Current.Server.MapPath(templatePath);

            designer.Workbook = new DongA.Core.DongAExcel().OpenExcelFile(templatePath);
            designer.Workbook.CalculateFormula();

            WorksheetCollection workSheets = designer.Workbook.Worksheets;
            Worksheet sheetReport = designer.Workbook.Worksheets[0];

            // Tạo title
            string typeReport = "So sánh - Theo giai đoạn";

            string text = " tháng đầu năm";
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
            CreateTitle("A2", "U2", sheetReport, typeReport, 14);

            // Tạo title detailt
            string titleDetailt = string.Format("Giai đoạn: {0}", text);
            CreateTitle("A3", "U3", sheetReport, titleDetailt, 12);

            // Tạo title cho doanh số chi trả theo thị trường
            string titleDS = "1. Theo doanh số chi trả theo thị trường";
            CreateTitle("A5", "E5", sheetReport, titleDS, 12);

            // Tạo title chotỷ trọng chi trả theo thị trường
            string titleTT = "2. Theo tỷ trọng chi trả theo thị trường";
            CreateTitle("A33", "E33", sheetReport, titleTT, 12);

            // Add tên cột cho bảng báo cáo
            // Set width cho column
            // Theo doanh số
            sheetReport.Cells.SetColumnWidthPixel(15, 80);
            sheetReport.Cells["Q7"].PutValue(string.Format("Lũy kế {0} \n {1}", text, year));
            sheetReport.Cells.SetColumnWidthPixel(16, 170);
            sheetReport.Cells["R7"].PutValue(string.Format("Lũy kế {0} \n {1}", text, year - 1));
            sheetReport.Cells.SetColumnWidthPixel(17, 190);
            sheetReport.Cells["S7"].PutValue(string.Format("Tăng giảm so với cùng kì \n năm {0} (%)", year - 1));
            sheetReport.Cells.SetColumnWidthPixel(18, 200);
            sheetReport.Cells["T7"].PutValue(string.Format("Tăng giảm so với cùng kì \n năm {0} (+/-)", year - 1));
            sheetReport.Cells.SetColumnWidthPixel(19, 200);
            sheetReport.Cells.SetRowHeight(6, 40);

            // Tạo cột cho tỉ trọng %
            sheetReport.Cells.SetColumnWidthPixel(15, 200);
            sheetReport.Cells["Q35"].PutValue(string.Format("Lũy kế {0} \n {1}", text, year));
            sheetReport.Cells.SetColumnWidthPixel(16, 200);
            sheetReport.Cells["R35"].PutValue(string.Format("Lũy kế {0} \n {1}", text, year - 1));
            sheetReport.Cells.SetColumnWidthPixel(17, 200);
            sheetReport.Cells["S35"].PutValue(string.Format("Tăng giảm so với cùng \n kỳ năm {0}", year - 1));

            // Set style cho các cột đã thay đổi
            string[] cellArray = { "Q7", "R7", "S7", "T7" };
            foreach (string item in cellArray)
            {
                //Get Cell's Style
                Style style = sheetReport.Cells[item].GetStyle();
                //Set Text Wrap property to true
                style.IsTextWrapped = true;
                //Set Cell's Style
                sheetReport.Cells[item].SetStyle(style);
            }

            // Create Chart column
            //Chart reference
            Aspose.Cells.Charts.Chart leadSourceLine;
            //Add Pie Chart
            int chartIndex = sheetReport.Charts.Add(ChartType.Column3DClustered, 6, 0, 30, 13);
            leadSourceLine = sheetReport.Charts[chartIndex];

            //// Canh hiển thị CategoryAxis nghiên phù hợp
            //leadSourceLine.CategoryAxis.TickLabels.RotationAngle = 45;

            //Chart title
            leadSourceLine.Title.Text = "Doanh số chi trả theo thị trường";
            leadSourceLine.Title.Font.Color = Color.Silver;


            List<ReportForMaket> listReportData = new ReportBL().DataReportMaketForGradationCompare(year, int.Parse(gradationID), reportTypeID);
            // clone Object
            List<ReportForMaket> listReportDataClone = new List<ReportForMaket>(listReportData);

            List<ReportForMaket> listReportDataPercent = new ReportBL().DataReportMaketForGradationComparePercent(year, int.Parse(gradationID), reportTypeID);
            // clone Object
            List<ReportForMaket> listReportDataPercentClone = new List<ReportForMaket>(listReportDataPercent);

            DataTable dataTable = new DataTable();

            string[] str = { "Mỹ", "Châu Á", "Toàn cầu", "Châu Âu", "Canada", "Úc" };
            // Theo doanh số chi trả theo thị trường
            if (listReportData.Count.Equals(2))
            {
                // Tạo các cột cho datatable
                dataTable.Columns.Add("MaketID", typeof(String));
                dataTable.Columns.Add("AccumulateID1", typeof(double));
                dataTable.Columns.Add("AccumulateID2", typeof(double));
                dataTable.Columns.Add("CompareToIDPercent", typeof(double));
                dataTable.Columns.Add("CompareToID", typeof(double));

                double americanCompare = listReportData[0].American - listReportData[1].American;
                double asiaCompare = listReportData[0].Asia - listReportData[1].Asia;
                double globalCompare = listReportData[0].Global - listReportData[1].Global;
                double europeCompare = listReportData[0].Europe - listReportData[1].Europe;
                double canadaCompare = listReportData[0].Canada - listReportData[1].Canada;
                double australiaCompare = listReportData[0].Australia - listReportData[1].Australia;

                // add row vào table
                dataTable.Rows.Add(str[0], listReportData[0].American, listReportData[1].American, Math.Round(americanCompare / listReportData[1].American * 100, 2, MidpointRounding.ToEven), americanCompare);
                dataTable.Rows.Add(str[1], listReportData[0].Asia, listReportData[1].Asia, Math.Round(asiaCompare / listReportData[1].Asia * 100, 2, MidpointRounding.ToEven), asiaCompare);
                dataTable.Rows.Add(str[2], listReportData[0].Global, listReportData[1].Global, Math.Round(globalCompare / listReportData[1].Global * 100, 2, MidpointRounding.ToEven), globalCompare);
                dataTable.Rows.Add(str[3], listReportData[0].Europe, listReportData[1].Europe, Math.Round(europeCompare / listReportData[1].Europe * 100, 2, MidpointRounding.ToEven), europeCompare);
                dataTable.Rows.Add(str[4], listReportData[0].Canada, listReportData[1].Canada, Math.Round(canadaCompare / listReportData[1].Canada * 100, 2, MidpointRounding.ToEven), canadaCompare);
                dataTable.Rows.Add(str[5], listReportData[0].Australia, listReportData[1].Australia, Math.Round(australiaCompare / listReportData[1].Australia * 100, 2, MidpointRounding.ToEven), australiaCompare);

                //////
                DataRow row = dataTable.NewRow();
                row["MaketID"] = "Tổng";
                double sumAccumulateID1 = 0;
                double sumAccumulateID2 = 0;

                if (!string.IsNullOrEmpty(dataTable.Compute("Sum(AccumulateID1)", "").ToString()))
                {
                    sumAccumulateID1 = Convert.ToDouble(dataTable.Compute("Sum(AccumulateID1)", ""));
                }

                if (!string.IsNullOrEmpty(dataTable.Compute("Sum(AccumulateID2)", "").ToString()))
                {
                    sumAccumulateID2 = Convert.ToDouble(dataTable.Compute("Sum(AccumulateID2)", ""));
                }

                row["AccumulateID1"] = sumAccumulateID1;
                row["AccumulateID2"] = sumAccumulateID2;
                row["CompareToID"] = Math.Round(sumAccumulateID1 - sumAccumulateID2, 2, MidpointRounding.ToEven);
                row["CompareToIDPercent"] = sumAccumulateID2 == 0 ? 0 : Math.Round((sumAccumulateID1 - sumAccumulateID2) / sumAccumulateID2 * 100, 2, MidpointRounding.ToEven);
                dataTable.Rows.Add(row);

                int count = 0;
                bool checkColor = true;
                foreach (ReportForMaket item in listReportDataClone)
                {
                    string totalRowData = string.Concat("{", string.Format("{0}, {1}, {2}, {3}, {4}, {5}", item.American, item.Asia, item.Global, item.Europe, item.Canada, item.Australia), "}");
                    leadSourceLine.NSeries.Add(totalRowData, true);

                    string categoryData = string.Concat("{", string.Join(", ", str), "}");
                    leadSourceLine.NSeries.CategoryData = categoryData;

                    leadSourceLine.NSeries[count].Name = string.Concat("Lũy kế ", text, " ", item.Year);

                    // Set the 2nd series fill color.
                    leadSourceLine.NSeries[count].Area.ForegroundColor = Color.Orange;
                    leadSourceLine.NSeries[count].Area.Formatting = FormattingType.Custom;

                    if (!checkColor)
                    {
                        // Set the 1st series fill color.
                        leadSourceLine.NSeries[count].Area.ForegroundColor = Color.Green;
                        leadSourceLine.NSeries[count].Area.Formatting = FormattingType.Custom;
                    }

                    checkColor = false;
                    count++;
                }

                // Set plot area formatting as none and hide its border.
                leadSourceLine.PlotArea.Area.FillFormat.FillType = FillType.None;
                leadSourceLine.PlotArea.Border.IsVisible = false;

                // Set value axis major tick mark as none and hide axis line. 
                // Also set the color of value axis major grid lines.
                leadSourceLine.ValueAxis.MajorTickMark = TickMarkType.None;
                leadSourceLine.ValueAxis.AxisLine.IsVisible = false;
                leadSourceLine.ValueAxis.MajorGridLines.Color = Color.FromArgb(217, 217, 217);

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
                        int totalCol = 15 + 5;
                        for (int b = 15; b < totalCol; b++)
                        {
                            // Giá trị của value trong table
                            string valueOfTable = dataTable.Rows[stepRow][stepColumn].ToString();

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
            }

            // Biểu đồ Doanh số theo loại hình chi trả của năm hiện tại
            if (listReportDataPercentClone.Count.Equals(2))
            {
                bool check = true;
                foreach (ReportForMaket item in listReportDataPercentClone)
                {
                    item.ReportID = string.Concat("Lũy kế ", text, " ", year);
                    if (!check)
                    {
                        item.ReportID = string.Concat("Lũy kế ", text, " ", year - 1);
                    }
                    item.TongDS = item.American + item.Asia + item.Global + item.Europe + item.Canada + item.Australia;
                    item.Type = 0;
                    // Set lại giá trị cho check để lấy giá trị của năm trước
                    check = false;
                }

                ReportForMaket dataPieYear = null;
                ReportForMaket dataPieLastYear = null;
                // Data report năm hiện tại nhập vào
                dataPieYear = listReportDataPercentClone.Find(x => x.Year == year.ToString());
                // Data report năm ngoái so với năm hiện tại nhập vào
                dataPieLastYear = listReportDataPercentClone.Find(x => x.Year == (year - 1).ToString());

                if (dataPieYear != null)
                {
                    //Add Pie Chart
                    Aspose.Cells.Charts.Chart leadSourcePie;
                    chartIndex = sheetReport.Charts.Add(ChartType.Pie3D, 34, 0, 49, 5);
                    leadSourcePie = sheetReport.Charts[chartIndex];

                    // Set some properties of chart plot area.
                    // To set the fill color and make the border invisible.
                    leadSourcePie.PlotArea.Border.IsVisible = false;
                    leadSourcePie.Elevation = 45;
                    // Set properties of chart title
                    leadSourcePie.Title.Text = string.Concat("Lũy kế ", text, " ", dataPieYear.Year);
                    leadSourcePie.Title.Font.Color = Color.Blue;
                    leadSourcePie.Title.Font.IsBold = true;
                    leadSourcePie.Title.Font.Size = 12;

                    // Set properties of nseries
                    string totalRowData = string.Concat("{", string.Format("{0}, {1}, {2}, {3}, {4}, {5}", dataPieYear.American, dataPieYear.Asia, dataPieYear.Global, dataPieYear.Europe, dataPieYear.Canada, dataPieYear.Australia), "}");
                    leadSourcePie.NSeries.Add(totalRowData, true);

                    string categoryData = string.Concat("{", string.Join(", ", str), "}");
                    leadSourcePie.NSeries.CategoryData = categoryData;

                    leadSourcePie.NSeries.IsColorVaried = true;

                    // Set the DataLabels in the chart
                    Aspose.Cells.Charts.DataLabels datalabels;
                    for (int i = 0; i < leadSourcePie.NSeries.Count; i++)
                    {
                        datalabels = leadSourcePie.NSeries[i].DataLabels;
                        datalabels.Position = Aspose.Cells.Charts.LabelPositionType.InsideBase;
                        datalabels.ShowCategoryName = true;
                        datalabels.ShowValue = true;
                        datalabels.ShowPercentage = false;
                        datalabels.ShowLegendKey = false;
                    }
                }

                if (dataPieLastYear != null)
                {
                    //Add Pie Chart
                    Aspose.Cells.Charts.Chart leadSourcePieLasYear;
                    chartIndex = sheetReport.Charts.Add(ChartType.Pie3D, 34, 8, 49, 13);
                    leadSourcePieLasYear = sheetReport.Charts[chartIndex];

                    // Set some properties of chart plot area.
                    // To set the fill color and make the border invisible.
                    leadSourcePieLasYear.PlotArea.Border.IsVisible = false;
                    leadSourcePieLasYear.Elevation = 45;
                    // Set properties of chart title
                    leadSourcePieLasYear.Title.Text = string.Concat("Lũy kế ", text, " ", dataPieLastYear.Year);
                    leadSourcePieLasYear.Title.Font.Color = Color.Blue;
                    leadSourcePieLasYear.Title.Font.IsBold = true;
                    leadSourcePieLasYear.Title.Font.Size = 12;

                    // Set properties of nseries
                    string totalRowData = string.Concat("{", string.Format("{0}, {1}, {2}, {3}, {4}, {5}", dataPieYear.American, dataPieYear.Asia, dataPieYear.Global, dataPieYear.Europe, dataPieYear.Canada, dataPieYear.Australia), "}");
                    leadSourcePieLasYear.NSeries.Add(totalRowData, true);

                    string categoryData = string.Concat("{", string.Join(", ", str), "}");
                    leadSourcePieLasYear.NSeries.CategoryData = categoryData;

                    leadSourcePieLasYear.NSeries.IsColorVaried = true;

                    // Set the DataLabels in the chart
                    Aspose.Cells.Charts.DataLabels datalabels;
                    for (int i = 0; i < leadSourcePieLasYear.NSeries.Count; i++)
                    {
                        datalabels = leadSourcePieLasYear.NSeries[i].DataLabels;
                        datalabels.Position = Aspose.Cells.Charts.LabelPositionType.InsideBase;
                        datalabels.ShowCategoryName = true;
                        datalabels.ShowValue = true;
                        datalabels.ShowPercentage = false;
                        datalabels.ShowLegendKey = false;

                    }
                }

                // ds so với tháng trước
                double americanCompare = listReportDataPercentClone[0].American - listReportDataPercentClone[1].American;
                double asiaCompare = listReportDataPercentClone[0].Asia - listReportDataPercentClone[1].Asia;
                double globalCompare = listReportDataPercentClone[0].Global - listReportDataPercentClone[1].Global;
                double europeCompare = listReportDataPercentClone[0].Europe - listReportDataPercentClone[1].Europe;
                double canadaCompare = listReportDataPercentClone[0].Canada - listReportDataPercentClone[1].Canada;
                double australiaCompare = listReportDataPercentClone[0].Australia - listReportDataPercentClone[1].Australia;

                // Tạo các cột cho datatable
                DataTable dataTablePie = new DataTable();

                dataTablePie.Columns.Add("MaketID", typeof(String));
                dataTablePie.Columns.Add("AccumulateID1", typeof(double));
                dataTablePie.Columns.Add("AccumulateID2", typeof(double));
                dataTablePie.Columns.Add("CompareToIDPercent", typeof(double));

                // add row vào table
                dataTablePie.Rows.Add(str[0], listReportDataPercentClone[0].American, listReportDataPercentClone[1].American, Math.Round(americanCompare, 2, MidpointRounding.ToEven));
                dataTablePie.Rows.Add(str[1], listReportDataPercentClone[0].Asia, listReportDataPercentClone[1].Asia, Math.Round(asiaCompare, 2, MidpointRounding.ToEven));
                dataTablePie.Rows.Add(str[2], listReportDataPercentClone[0].Global, listReportDataPercentClone[1].Global, Math.Round(globalCompare, 2, MidpointRounding.ToEven));
                dataTablePie.Rows.Add(str[3], listReportDataPercentClone[0].Europe, listReportDataPercentClone[1].Europe, Math.Round(europeCompare, 2, MidpointRounding.ToEven));
                dataTablePie.Rows.Add(str[4], listReportDataPercentClone[0].Canada, listReportDataPercentClone[1].Canada, Math.Round(canadaCompare, 2, MidpointRounding.ToEven));
                dataTablePie.Rows.Add(str[5], listReportDataPercentClone[0].Australia, listReportDataPercentClone[1].Australia, Math.Round(australiaCompare, 2, MidpointRounding.ToEven));

                DataRow row = dataTablePie.NewRow();
                row["MaketID"] = "Tổng";
                row["AccumulateID1"] = dataTablePie.Compute("Sum(AccumulateID1)", "");
                row["AccumulateID2"] = dataTablePie.Compute("Sum(AccumulateID2)", "");

                // Sum row tổng compare month
                double sumCompareMonth = (double)row["AccumulateID1"] - (double)row["AccumulateID2"];

                row["CompareToIDPercent"] = Math.Round(sumCompareMonth, 2, MidpointRounding.ToEven);
                dataTablePie.Rows.Add(row);


                // Set border
                Style style = new CellsFactory().CreateStyle();
                style.SetBorder(BorderType.BottomBorder, CellBorderType.Thin, Color.Black);
                style.SetBorder(BorderType.LeftBorder, CellBorderType.Thin, Color.Black);
                style.SetBorder(BorderType.RightBorder, CellBorderType.Thin, Color.Black);

                if (dataTablePie.Rows.Count > 0)
                {
                    int stepRow = 0;
                    // total row = row start + số row hiện có
                    int totalRow = dataTablePie.Rows.Count + 35;
                    // Số dòng của row
                    for (int a = 35; a < totalRow; a++)
                    {
                        int stepColumn = 0;
                        // Số cột trong báo cáo cần hiển thị
                        // Tổng số cột hiển thị = Số cột hiển thị bắt đầu + tổng số cột cần hiển thị
                        int totalCol = 15 + 4;
                        for (int b = 15; b < totalCol; b++)
                        {
                            // Giá trị của value trong table
                            string valueOfTable = dataTablePie.Rows[stepRow][stepColumn].ToString();

                            // Tô màu cho các dòng có giá trị tăng giảm
                            if (b >= 18 && a < totalRow - 1)
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
        public ActionResult CreateExcelForCompareYear(int year, int month, string reportTypeID)
        {
            WorkbookDesigner designer = new WorkbookDesigner();
            string templatePath = "~/Content/Report/ReportMaketForGradationCompare.xlsx";
            // Get đường dẫn
            templatePath = System.Web.HttpContext.Current.Server.MapPath(templatePath);

            designer.Workbook = new DongA.Core.DongAExcel().OpenExcelFile(templatePath);
            designer.Workbook.CalculateFormula();

            WorksheetCollection workSheets = designer.Workbook.Worksheets;
            Worksheet sheetReport = designer.Workbook.Worksheets[0];

            // Tạo title
            string typeReport = "So sánh - Theo tháng";

            string text = string.Format("Tháng: {0}/{1}", month, year);

            // Tạo title
            CreateTitle("A2", "U2", sheetReport, typeReport, 14);

            // Tạo title detailt
            string titleDetailt = text;
            CreateTitle("A3", "U3", sheetReport, titleDetailt, 12);

            // Tạo title cho doanh số chi trả theo thị trường
            string titleDS = "1. Theo doanh số chi trả theo thị trường";
            CreateTitle("A5", "E5", sheetReport, titleDS, 12);

            // Tạo title chotỷ trọng chi trả theo thị trường
            string titleTT = "2. Theo tỷ trọng chi trả theo thị trường";
            CreateTitle("A33", "E33", sheetReport, titleTT, 12);

            // Add tên cột cho bảng báo cáo
            // Set width cho column
            // Theo doanh số
            //sheetReport.Cells.SetColumnWidthPixel(15, 80);
            sheetReport.Cells["S6"].PutValue("");
            sheetReport.Cells["T6"].PutValue("");

            sheetReport.Cells["V6"].PutValue("Đơn vị:");
            Style styleCell = sheetReport.Cells["V6"].GetStyle();
            styleCell.Font.IsBold = true;
            styleCell.HorizontalAlignment = TextAlignmentType.Right;
            sheetReport.Cells["V6"].SetStyle(styleCell);
            // set style cho cell
            sheetReport.Cells["W6"].PutValue("Triệu USD");
            styleCell = sheetReport.Cells["W6"].GetStyle();
            styleCell.Font.IsBold = true;
            styleCell.HorizontalAlignment = TextAlignmentType.Right;
            sheetReport.Cells["W6"].SetStyle(styleCell);

            sheetReport.Cells["Q7"].PutValue(string.Format("Tháng {0}/{1}", month, year));
            //sheetReport.Cells.SetColumnWidthPixel(16, 170);

            // Trường hợp là tháng 1
            sheetReport.Cells["R7"].PutValue(string.Format("Tháng {0}/{1}", month - 1, year));
            if (month - 1 == 0)
            {
                sheetReport.Cells["R7"].PutValue(string.Format("Tháng {0}/{1}", 12, year - 1));
            }
            //sheetReport.Cells.SetColumnWidthPixel(17, 190);
            sheetReport.Cells["S7"].PutValue(string.Format("Tháng {0}/{1}", month, year - 1));
            sheetReport.Cells.SetColumnWidthPixel(19, 200);
            sheetReport.Cells["T7"].PutValue("Tăng giảm so với tháng (+/-)");
            sheetReport.Cells.SetColumnWidthPixel(20, 200);
            sheetReport.Cells["U7"].PutValue("Tăng giảm so với tháng (%)");
            sheetReport.Cells.SetColumnWidthPixel(21, 200);
            sheetReport.Cells["V7"].PutValue("Tăng giảm so với cùng kì \n năm trước (+/-)");
            sheetReport.Cells.SetColumnWidthPixel(22, 200);
            sheetReport.Cells["W7"].PutValue("Tăng giảm so với cùng kì \n năm trước (%)");
            sheetReport.Cells.SetRowHeight(6, 40);

            // Tạo cột cho tỉ trọng %
            //sheetReport.Cells.SetColumnWidthPixel(15, 200);
            sheetReport.Cells["Q35"].PutValue(string.Format("Tháng {0}/{1}", month, year));
            //sheetReport.Cells.SetColumnWidthPixel(16, 200);
            sheetReport.Cells["R35"].PutValue(string.Format("Tháng {0}/{1}", month - 1, year));

            // Trường hợp là tháng 1
            if (month - 1 == 0)
            {
                sheetReport.Cells["R35"].PutValue(string.Format("Tháng {0}/{1}", 12, year - 1));
            }
            //sheetReport.Cells.SetColumnWidthPixel(17, 200);
            sheetReport.Cells["S35"].PutValue(string.Format("Tháng {0}/{1}", month, year - 1));
            sheetReport.Cells.SetColumnWidthPixel(19, 200);
            sheetReport.Cells["T35"].PutValue("Tăng giảm so với tháng (%)");
            sheetReport.Cells.SetColumnWidthPixel(20, 200);
            sheetReport.Cells["U35"].PutValue("Tăng giảm so với cùng kì \n năm trước (%)");
            sheetReport.Cells.SetRowHeight(34, 40);

            //sheetReport.Cells.SetColumnWidthPixel(15, 80);
            sheetReport.Cells["R34"].PutValue("");
            sheetReport.Cells["S34"].PutValue("");

            sheetReport.Cells["T34"].PutValue("Đơn vị:");
            styleCell = sheetReport.Cells["T34"].GetStyle();
            styleCell.Font.IsBold = true;
            styleCell.HorizontalAlignment = TextAlignmentType.Right;
            sheetReport.Cells["T34"].SetStyle(styleCell);
            // set style cho cell
            sheetReport.Cells["U34"].PutValue("%");
            styleCell = sheetReport.Cells["U34"].GetStyle();
            styleCell.Font.IsBold = true;
            styleCell.HorizontalAlignment = TextAlignmentType.Right;
            sheetReport.Cells["U34"].SetStyle(styleCell);

            // Set style cho các cột đã thay đổi
            string[] cellArray = { "Q7", "R7", "S7", "T7", "U7", "V7", "W7", "Q35", "R35", "S35", "T35", "U35" };
            foreach (string item in cellArray)
            {
                //Get Cell's Style
                Style style = sheetReport.Cells[item].GetStyle();
                //Set Text Wrap property to true
                style.IsTextWrapped = true;
                // canh ngang
                style.HorizontalAlignment = TextAlignmentType.Center;
                // Canh dọc
                style.VerticalAlignment = TextAlignmentType.Center;
                // Set border
                style.SetBorder(BorderType.TopBorder, CellBorderType.Thin, Color.Black);
                style.SetBorder(BorderType.RightBorder, CellBorderType.Thin, Color.Black);
                style.SetBorder(BorderType.BottomBorder, CellBorderType.Thin, Color.Black);
                style.Font.IsBold = true;
                //Set Cell's Style
                sheetReport.Cells[item].SetStyle(style);
            }

            // Create Chart Line
            //Chart reference
            Aspose.Cells.Charts.Chart leadSourceLine;
            //Add Pie Chart
            int chartIndex = sheetReport.Charts.Add(ChartType.Column3DClustered, 6, 0, 30, 13);
            leadSourceLine = sheetReport.Charts[chartIndex];

            //// Canh hiển thị CategoryAxis nghiên phù hợp
            //leadSourceLine.CategoryAxis.TickLabels.RotationAngle = 45;

            //Chart title
            leadSourceLine.Title.Text = string.Format("Doanh số theo thị trường tháng {0}/{1} \n so với tháng trước và so với cùng kì năm trước", month, year);
            leadSourceLine.Title.Font.Color = Color.Silver;

            // Set width cho column
            sheetReport.Cells.SetColumnWidthPixel(15, 220);

            List<ReportForMaket> listReportData = new ReportBL().DataReportCompareForMonth(year, month, reportTypeID);
            // clone Object
            List<ReportForMaket> listReportDataClone = new List<ReportForMaket>(listReportData);

            List<ReportForMaket> listReportDataPercent = new ReportBL().DataReportMarketCompareForMonthPercent(year, month, reportTypeID);
            // clone Object
            List<ReportForMaket> listReportDataPercentClone = new List<ReportForMaket>(listReportDataPercent);

            DataTable dataTable = new DataTable();

            string[] str = { "Mỹ", "Châu Á", "Toàn cầu", "Châu Âu", "Canada", "Úc" };

            // Theo doanh số chi trả thị trường
            if (listReportData.Count.Equals(3))
            {
                // Tạo các cột cho datatable
                dataTable.Columns.Add("MaketID", typeof(String));
                dataTable.Columns.Add("AccumulateID1", typeof(double));
                dataTable.Columns.Add("AccumulateID2", typeof(double));
                dataTable.Columns.Add("AccumulateID3", typeof(double));
                dataTable.Columns.Add("CompareToMonth", typeof(double));
                dataTable.Columns.Add("CompareToMonthPercent", typeof(double));
                dataTable.Columns.Add("CompareToMonthLastYear", typeof(double));
                dataTable.Columns.Add("CompareToMonthLastYearPercent", typeof(double));

                // tháng hiện tại
                double americanCompareMonth = listReportData[0].American - listReportData[1].American;
                double asiaCompareMonth = listReportData[0].Asia - listReportData[1].Asia;
                double globalCompareMonth = listReportData[0].Global - listReportData[1].Global;
                double europeCompareMonth = listReportData[0].Europe - listReportData[1].Europe;
                double canadaCompareMonth = listReportData[0].Canada - listReportData[1].Canada;
                double australiaCompareMonth = listReportData[0].Australia - listReportData[1].Australia;

                // tháng cùng kì năm trước
                double americanCompareMonthLastYear = listReportData[0].American - listReportData[2].American;
                double asiaCompareMonthLastYear = listReportData[0].Asia - listReportData[2].Asia;
                double globalCompareMonthLastYear = listReportData[0].Global - listReportData[2].Global;
                double europeCompareMonthLastYear = listReportData[0].Europe - listReportData[2].Europe;
                double canadaCompareMonthLastYear = listReportData[0].Canada - listReportData[2].Canada;
                double australiaCompareMonthLastYear = listReportData[0].Australia - listReportData[2].Australia;

                // add row vào table
                // add row vào table
                dataTable.Rows.Add(str[0], listReportData[0].American, listReportData[1].American, listReportData[2].American
                    , americanCompareMonth
                    , Math.Round(americanCompareMonth / listReportData[1].American * 100, 2, MidpointRounding.ToEven)
                    , americanCompareMonthLastYear
                    , Math.Round(americanCompareMonthLastYear / listReportData[2].American * 100, 2, MidpointRounding.ToEven));

                dataTable.Rows.Add(str[1], listReportData[0].Asia, listReportData[1].Asia, listReportData[2].Asia
                    , asiaCompareMonth
                    , Math.Round(asiaCompareMonth / listReportData[1].Asia * 100, 2, MidpointRounding.ToEven)
                    , asiaCompareMonthLastYear
                    , Math.Round(asiaCompareMonthLastYear / listReportData[2].Asia * 100, 2, MidpointRounding.ToEven));

                dataTable.Rows.Add(str[2], listReportData[0].Global, listReportData[1].Global, listReportData[2].Global
                    , globalCompareMonth
                    , Math.Round(globalCompareMonth / listReportData[1].Global * 100, 2, MidpointRounding.ToEven)
                    , globalCompareMonthLastYear
                    , Math.Round(globalCompareMonthLastYear / listReportData[2].Global * 100, 2, MidpointRounding.ToEven));

                dataTable.Rows.Add(str[3], listReportData[0].Europe, listReportData[1].Europe, listReportData[2].Europe
                    , europeCompareMonth
                    , Math.Round(europeCompareMonth / listReportData[1].Europe * 100, 2, MidpointRounding.ToEven)
                    , europeCompareMonthLastYear
                    , Math.Round(europeCompareMonthLastYear / listReportData[2].Europe * 100, 2, MidpointRounding.ToEven));

                dataTable.Rows.Add(str[4], listReportData[0].Canada, listReportData[1].Canada, listReportData[2].Canada
                    , canadaCompareMonth
                    , Math.Round(canadaCompareMonth / listReportData[1].Canada * 100, 2, MidpointRounding.ToEven)
                    , canadaCompareMonthLastYear
                    , Math.Round(canadaCompareMonthLastYear / listReportData[2].Canada * 100, 2, MidpointRounding.ToEven));

                dataTable.Rows.Add(str[5], listReportData[0].Australia, listReportData[1].Australia, listReportData[2].Australia
                    , australiaCompareMonth
                    , Math.Round(australiaCompareMonth / listReportData[1].Australia * 100, 2, MidpointRounding.ToEven)
                    , australiaCompareMonthLastYear
                    , Math.Round(australiaCompareMonthLastYear / listReportData[2].Australia * 100, 2, MidpointRounding.ToEven));

                DataRow row = dataTable.NewRow();
                row["MaketID"] = "Tổng";
                row["AccumulateID1"] = dataTable.Compute("Sum(AccumulateID1)", "");
                row["AccumulateID2"] = dataTable.Compute("Sum(AccumulateID2)", "");
                row["AccumulateID3"] = dataTable.Compute("Sum(AccumulateID3)", "");

                // Sum row tổng compare month
                double sumCompareMonth = (double)row["AccumulateID1"] - (double)row["AccumulateID2"];
                double sumCompareMonthLastYear = (double)row["AccumulateID1"] - (double)row["AccumulateID3"];
                // tính tăng giảm so với tháng trước theo (+/-) và %
                row["CompareToMonth"] = Math.Round(sumCompareMonth, 2, MidpointRounding.ToEven);
                row["CompareToMonthPercent"] = Math.Round(sumCompareMonth / (double)row["AccumulateID2"] * 100, 2, MidpointRounding.ToEven);
                // tính tăng giảm so với cùng kì năm trước (+/-) và %
                row["CompareToMonthLastYear"] = Math.Round(sumCompareMonthLastYear, 2, MidpointRounding.ToEven);
                row["CompareToMonthLastYearPercent"] = Math.Round(sumCompareMonthLastYear / (double)row["AccumulateID3"] * 100, 2, MidpointRounding.ToEven);

                dataTable.Rows.Add(row);

                int count = 0;
                foreach (ReportForMaket item in listReportDataClone)
                {
                    string totalRowData = string.Concat("{", string.Format("{0}, {1}, {2}, {3}, {4}, {5}", item.American, item.Asia, item.Global, item.Europe, item.Canada, item.Australia), "}");
                    leadSourceLine.NSeries.Add(totalRowData, true);

                    string categoryData = string.Concat("{", string.Join(", ", str), "}");
                    leadSourceLine.NSeries.CategoryData = categoryData;

                    leadSourceLine.NSeries[count].Name = string.Format("Tháng {0}/{1}", item.Month, item.Year);

                    // Set the 2nd series fill color.
                    leadSourceLine.NSeries[count].Area.ForegroundColor = Color.Orange;
                    leadSourceLine.NSeries[count].Area.Formatting = FormattingType.Custom;

                    if (count.Equals(1))
                    {
                        // Set the 1st series fill color.
                        leadSourceLine.NSeries[count].Area.ForegroundColor = Color.Green;
                        leadSourceLine.NSeries[count].Area.Formatting = FormattingType.Custom;
                    }

                    if (count.Equals(2))
                    {
                        // Set the 1st series fill color.
                        leadSourceLine.NSeries[count].Area.ForegroundColor = Color.Blue;
                        leadSourceLine.NSeries[count].Area.Formatting = FormattingType.Custom;
                    }
                    count++;
                }

                // Set plot area formatting as none and hide its border.
                leadSourceLine.PlotArea.Area.FillFormat.FillType = FillType.None;
                leadSourceLine.PlotArea.Border.IsVisible = false;

                // Set value axis major tick mark as none and hide axis line. 
                // Also set the color of value axis major grid lines.
                leadSourceLine.ValueAxis.MajorTickMark = TickMarkType.None;
                leadSourceLine.ValueAxis.AxisLine.IsVisible = false;
                leadSourceLine.ValueAxis.MajorGridLines.Color = Color.FromArgb(217, 217, 217);

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
                        int totalCol = 15 + 8;
                        for (int b = 15; b < totalCol; b++)
                        {
                            // Giá trị của value trong table
                            string valueOfTable = dataTable.Rows[stepRow][stepColumn].ToString();

                            // Tô màu cho các dòng có giá trị tăng giảm
                            if (b >= 19)
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
            }

            // Biểu đồ Doanh số theo loại hình chi trả của năm hiện tại
            if (listReportDataPercentClone.Count.Equals(3))
            {
                foreach (ReportForMaket item in listReportDataPercentClone)
                {
                    item.ReportID = string.Format("Tháng {0}/{1}", item.Month, item.Year);
                }

                ReportForMaket dataPieYear = null;
                ReportForMaket dataPieLastMonth = null;
                ReportForMaket dataPieLastYear = null;
                // Data report năm hiện tại nhập vào
                dataPieYear = listReportDataPercentClone.Find(x => x.Year == year.ToString() && x.Month == month.ToString());
                // Data report năm hiện tại, tháng hiện tại nhập vào

                dataPieLastMonth = listReportDataPercentClone.Find(x => x.Year == year.ToString() && x.Month == (month - 1).ToString());

                // Trường hợp là tháng 1
                if (month - 1 == 0)
                {
                    dataPieLastMonth = listReportDataPercentClone.Find(x => x.Year == (year - 1).ToString() && x.Month == (12).ToString());
                }
                // Data report năm ngoái so với năm hiện tại nhập vào
                dataPieLastYear = listReportDataPercentClone.Find(x => x.Year == (year - 1).ToString() && x.Month == month.ToString());

                if (dataPieYear != null)
                {
                    //Add Pie Chart
                    Aspose.Cells.Charts.Chart leadSourcePie;
                    chartIndex = sheetReport.Charts.Add(ChartType.Pie3D, 34, 0, 49, 6);
                    leadSourcePie = sheetReport.Charts[chartIndex];

                    // Set some properties of chart plot area.
                    // To set the fill color and make the border invisible.
                    leadSourcePie.PlotArea.Border.IsVisible = false;
                    leadSourcePie.Elevation = 45;
                    // Set properties of chart title
                    leadSourcePie.Title.Text = dataPieYear.ReportID;
                    leadSourcePie.Title.Font.Color = Color.Silver;
                    leadSourcePie.Title.Font.IsBold = true;
                    leadSourcePie.Title.Font.Size = 12;

                    // Set properties of nseries
                    string totalRowData = string.Concat("{", string.Format("{0}, {1}, {2}, {3}, {4}, {5}", dataPieYear.American, dataPieYear.Asia, dataPieYear.Global, dataPieYear.Europe, dataPieYear.Canada, dataPieYear.Australia), "}");
                    leadSourcePie.NSeries.Add(totalRowData, true);

                    string categoryData = string.Concat("{", string.Join(", ", str), "}");
                    leadSourcePie.NSeries.CategoryData = categoryData;

                    leadSourcePie.NSeries.IsColorVaried = true;

                    // Set the DataLabels in the chart
                    Aspose.Cells.Charts.DataLabels datalabels;
                    for (int i = 0; i < leadSourcePie.NSeries.Count; i++)
                    {
                        datalabels = leadSourcePie.NSeries[i].DataLabels;
                        datalabels.Position = Aspose.Cells.Charts.LabelPositionType.InsideBase;
                        datalabels.ShowCategoryName = true;
                        datalabels.ShowValue = true;
                        datalabels.ShowPercentage = false;
                        datalabels.ShowLegendKey = false;

                    }
                }

                if (dataPieLastYear != null)
                {
                    //Add Pie Chart
                    Aspose.Cells.Charts.Chart leadSourcePieLasYear;
                    chartIndex = sheetReport.Charts.Add(ChartType.Pie3D, 34, 8, 49, 14);
                    leadSourcePieLasYear = sheetReport.Charts[chartIndex];

                    // Set some properties of chart plot area.
                    // To set the fill color and make the border invisible.
                    leadSourcePieLasYear.PlotArea.Border.IsVisible = false;
                    leadSourcePieLasYear.Elevation = 45;
                    // Set properties of chart title
                    leadSourcePieLasYear.Title.Text = dataPieLastYear.ReportID;
                    leadSourcePieLasYear.Title.Font.Color = Color.Silver;
                    leadSourcePieLasYear.Title.Font.IsBold = true;
                    leadSourcePieLasYear.Title.Font.Size = 12;

                    // Set properties of nseries
                    string totalRowData = string.Concat("{", string.Format("{0}, {1}, {2}, {3}, {4}, {5}", dataPieLastYear.American, dataPieLastYear.Asia, dataPieLastYear.Global, dataPieLastYear.Europe, dataPieLastYear.Canada, dataPieLastYear.Australia), "}");
                    leadSourcePieLasYear.NSeries.Add(totalRowData, true);

                    string categoryData = string.Concat("{", string.Join(", ", str), "}");
                    leadSourcePieLasYear.NSeries.CategoryData = categoryData;

                    leadSourcePieLasYear.NSeries.IsColorVaried = true;

                    // Set the DataLabels in the chart
                    Aspose.Cells.Charts.DataLabels datalabels;
                    for (int i = 0; i < leadSourcePieLasYear.NSeries.Count; i++)
                    {
                        datalabels = leadSourcePieLasYear.NSeries[i].DataLabels;
                        datalabels.Position = Aspose.Cells.Charts.LabelPositionType.InsideBase;
                        datalabels.ShowCategoryName = true;
                        datalabels.ShowValue = true;
                        datalabels.ShowPercentage = false;
                        datalabels.ShowLegendKey = false;

                    }
                }

                if (dataPieLastMonth != null)
                {
                    //Add Pie Chart
                    Aspose.Cells.Charts.Chart leadSourcePie;
                    chartIndex = sheetReport.Charts.Add(ChartType.Pie3D, 51, 0, 65, 6);
                    leadSourcePie = sheetReport.Charts[chartIndex];

                    // Set some properties of chart plot area.
                    // To set the fill color and make the border invisible.
                    leadSourcePie.PlotArea.Border.IsVisible = false;
                    leadSourcePie.Elevation = 45;
                    // Set properties of chart title
                    leadSourcePie.Title.Text = dataPieLastMonth.ReportID;
                    leadSourcePie.Title.Font.Color = Color.Silver;
                    leadSourcePie.Title.Font.IsBold = true;
                    leadSourcePie.Title.Font.Size = 12;

                    // Set properties of nseries
                    string totalRowData = string.Concat("{", string.Format("{0}, {1}, {2}, {3}, {4}, {5}", dataPieLastMonth.American, dataPieLastMonth.Asia, dataPieLastMonth.Global, dataPieLastMonth.Europe, dataPieLastMonth.Canada, dataPieLastMonth.Australia), "}");
                    leadSourcePie.NSeries.Add(totalRowData, true);

                    string categoryData = string.Concat("{", string.Join(", ", str), "}");
                    leadSourcePie.NSeries.CategoryData = categoryData;

                    leadSourcePie.NSeries.IsColorVaried = true;

                    // Set the DataLabels in the chart
                    Aspose.Cells.Charts.DataLabels datalabels;
                    for (int i = 0; i < leadSourcePie.NSeries.Count; i++)
                    {
                        datalabels = leadSourcePie.NSeries[i].DataLabels;
                        datalabels.Position = Aspose.Cells.Charts.LabelPositionType.InsideBase;
                        datalabels.ShowCategoryName = true;
                        datalabels.ShowValue = true;
                        datalabels.ShowPercentage = false;
                        datalabels.ShowLegendKey = false;

                    }
                }

                // ds so với tháng trước
                double americanCompare = listReportDataPercentClone[0].American - listReportDataPercentClone[1].American;
                double asiaCompare = listReportDataPercentClone[0].Asia - listReportDataPercentClone[1].Asia;
                double globalCompare = listReportDataPercentClone[0].Global - listReportDataPercentClone[1].Global;
                double europeCompare = listReportDataPercentClone[0].Europe - listReportDataPercentClone[1].Europe;
                double canadaCompare = listReportDataPercentClone[0].Canada - listReportDataPercentClone[1].Canada;
                double australiaCompare = listReportDataPercentClone[0].Australia - listReportDataPercentClone[1].Australia;

                // tháng cùng kì năm trước
                double americanCompareMonthLastYear = listReportDataPercentClone[0].American - listReportDataPercentClone[2].American;
                double asiaCompareMonthLastYear = listReportDataPercentClone[0].Asia - listReportDataPercentClone[2].Asia;
                double globalCompareMonthLastYear = listReportDataPercentClone[0].Global - listReportDataPercentClone[2].Global;
                double europeCompareMonthLastYear = listReportDataPercentClone[0].Europe - listReportDataPercentClone[2].Europe;
                double canadaCompareMonthLastYear = listReportDataPercentClone[0].Canada - listReportDataPercentClone[2].Canada;
                double australiaCompareMonthLastYear = listReportDataPercentClone[0].Australia - listReportDataPercentClone[2].Australia;

                // Tạo các cột cho datatable
                DataTable dataTablePie = new DataTable();

                dataTablePie.Columns.Add("MaketID", typeof(String));
                dataTablePie.Columns.Add("AccumulateID1", typeof(double));
                dataTablePie.Columns.Add("AccumulateID2", typeof(double));
                dataTablePie.Columns.Add("AccumulateID3", typeof(double));
                dataTablePie.Columns.Add("CompareToMonthPercent", typeof(double));
                dataTablePie.Columns.Add("CompareToMonthLastYearPercent", typeof(double));

                // add row vào table
                dataTablePie.Rows.Add(str[0], listReportDataPercentClone[0].American, listReportDataPercentClone[1].American
                    , listReportDataPercentClone[2].American
                    , Math.Round(americanCompare / listReportDataPercentClone[1].American * 100, 2, MidpointRounding.ToEven)
                    , Math.Round(americanCompareMonthLastYear / listReportDataPercentClone[2].American * 100, 2, MidpointRounding.ToEven));

                dataTablePie.Rows.Add(str[1], listReportDataPercentClone[0].Asia, listReportDataPercentClone[1].Asia
                    , listReportDataPercentClone[2].Asia
                    , Math.Round(asiaCompare / listReportDataPercentClone[1].Asia * 100, 2, MidpointRounding.ToEven)
                    , Math.Round(asiaCompareMonthLastYear / listReportDataPercentClone[2].Asia * 100, 2, MidpointRounding.ToEven));

                dataTablePie.Rows.Add(str[2], listReportDataPercentClone[0].Global, listReportDataPercentClone[1].Global
                    , listReportDataPercentClone[2].Global
                    , Math.Round(globalCompare / listReportDataPercentClone[1].Global * 100, 2, MidpointRounding.ToEven)
                    , Math.Round(globalCompareMonthLastYear / listReportDataPercentClone[2].Global * 100, 2, MidpointRounding.ToEven));

                dataTablePie.Rows.Add(str[3], listReportDataPercentClone[0].Europe, listReportDataPercentClone[1].Europe
                    , listReportDataPercentClone[2].Europe
                    , Math.Round(europeCompare / listReportDataPercentClone[1].Europe * 100, 2, MidpointRounding.ToEven)
                    , Math.Round(europeCompareMonthLastYear / listReportDataPercentClone[2].Europe * 100, 2, MidpointRounding.ToEven));

                dataTablePie.Rows.Add(str[4], listReportDataPercentClone[0].Canada, listReportDataPercentClone[1].Canada
                    , listReportDataPercentClone[2].Canada
                    , Math.Round(canadaCompare / listReportDataPercentClone[1].Canada * 100, 2, MidpointRounding.ToEven)
                    , Math.Round(canadaCompareMonthLastYear / listReportDataPercentClone[2].Canada * 100, 2, MidpointRounding.ToEven));

                dataTablePie.Rows.Add(str[5], listReportDataPercentClone[0].Australia, listReportDataPercentClone[1].Australia
                    , listReportDataPercentClone[2].Australia
                    , Math.Round(australiaCompare / listReportDataPercentClone[1].Australia * 100, 2, MidpointRounding.ToEven)
                    , Math.Round(australiaCompareMonthLastYear / listReportDataPercentClone[2].Australia * 100, 2, MidpointRounding.ToEven));

                DataRow row = dataTablePie.NewRow();
                row["MaketID"] = "Tổng";
                row["AccumulateID1"] = 100;
                row["AccumulateID2"] = 100;
                row["AccumulateID3"] = 100;

                row["CompareToMonthPercent"] = 0;
                row["CompareToMonthLastYearPercent"] = 0;

                dataTablePie.Rows.Add(row);
                // Set border
                Style style = new CellsFactory().CreateStyle();
                style.SetBorder(BorderType.BottomBorder, CellBorderType.Thin, Color.Black);
                style.SetBorder(BorderType.LeftBorder, CellBorderType.Thin, Color.Black);
                style.SetBorder(BorderType.RightBorder, CellBorderType.Thin, Color.Black);

                if (dataTablePie.Rows.Count > 0)
                {
                    int stepRow = 0;
                    // total row = row start + số row hiện có
                    int totalRow = dataTablePie.Rows.Count + 35;
                    // Số dòng của row
                    for (int a = 35; a < totalRow; a++)
                    {
                        int stepColumn = 0;
                        // Số cột trong báo cáo cần hiển thị
                        // Tổng số cột hiển thị = Số cột hiển thị bắt đầu + tổng số cột cần hiển thị
                        int totalCol = 15 + 6;
                        for (int b = 15; b < totalCol; b++)
                        {
                            // Giá trị của value trong table
                            string valueOfTable = dataTablePie.Rows[stepRow][stepColumn].ToString();

                            // Tô màu cho các dòng có giá trị tăng giảm
                            if (b >= 19 && a < totalRow - 1)
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
                            if (b.Equals(totalCol - 1) || b.Equals(totalCol - 2))
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
            }

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
        private void CreateTitle(string upperLeft, string upperRight, Worksheet sheetReport, string Title, int size)
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
﻿using Aspose.Cells;
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
    public class ReportHSDetailtLTMoneyTypeExcelController : Controller
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
            string templatePath = "~/Content/Report/ReportHSDetailt/ReportHSDetailtLTMoneyType.xlsx";
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
                    listReportData = new HSReportBL().SearchReportDetailtMTForAll(fromDate, toDate, reportTypeID, marketID);

                    foreach (ReportDetailtForTotalMoneyType item in listReportData)
                    {
                        item.ReportID = item.MarketName;
                        if (marketID.Equals("0"))
                        {
                            item.MarketName = "";
                        }
                        item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;
                    }
                    
                    // Set from day and to day
                    sheetReport.Cells["E4"].PutValue(fromDate.ToString("dd/MM/yyyy"));
                    sheetReport.Cells["H4"].PutValue(toDate.ToString("dd/MM/yyyy"));
                    break;
                // Theo tháng
                case 2:

                    listReportData = new HSReportBL().SearchReportDetailtMTForAllForMonth(fromDate, toDate, reportTypeID, marketID);

                    foreach (ReportDetailtForTotalMoneyType item in listReportData)
                    {
                        item.ReportID = item.MarketName;
                        if (marketID.Equals("0"))
                        {
                            item.MarketName = "";
                        }
                        item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;
                    }

                    // Set from day and to day
                    sheetReport.Cells["E4"].PutValue(string.Format("{0}/{1}", fromDate.Month, fromDate.Year));
                    sheetReport.Cells["H4"].PutValue(string.Format("{0}/{1}", toDate.Month, toDate.Year));
                    break;
                // Theo năm
                default:

                    listReportData = new HSReportBL().SearchReportDetailtMTForAllForYear(fromDate, toDate, reportTypeID, marketID);

                    foreach (ReportDetailtForTotalMoneyType item in listReportData)
                    {
                        item.ReportID = item.MarketName;
                        if (marketID.Equals("0"))
                        {
                            item.MarketName = "";
                        }
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
                dataTable = CreateDataTableFormart(true);

                // Danh sách dataSet của báo cáo ngày
                DataSet dataReport = ConvertListObjectToDataSet(listReportData);

                // Đổ data vào datatble mới
                FillData(dataReport.Tables[0], dataTable, true);
            }
            
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
                            style.Custom = "#,##0";
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
            string templatePath = "~/Content/Report/ReportHSDetailt/ReportHSDetailtLTMoneyType.xlsx";
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
            string nameMarket = string.Empty;

            switch (typeID)
            {
                // Theo ngày
                case 1:
                    listReportData = new HSReportBL().SearchReportDetailtMTForOne(fromDate, toDate, reportTypeID, marketID);

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
                                    TongDS = listDataItem.Sum(x => x.TongDS),
                                }
                            );
                        }

                        if (listDataConvert.Count > 0)
                        {
                            listReportData = new List<ReportDetailtForTotalMoneyType>(listDataConvert);
                        }
                    }
                    else
                    {

                        foreach (ReportDetailtForTotalMoneyType item in listReportData)
                        {
                            item.ReportID = item.PartnerName;
                            item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;
                        }
                    }

                    // Set from day and to day
                    sheetReport.Cells["E4"].PutValue(fromDate.ToString("dd/MM/yyyy"));
                    sheetReport.Cells["H4"].PutValue(toDate.ToString("dd/MM/yyyy"));
                    break;
                // Theo tháng
                case 2:

                    listReportData = new HSReportBL().SearchReportDetailtMTForOneForMonth(fromDate, toDate, reportTypeID, marketID);

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
                                    TongDS = listDataItem.Sum(x => x.TongDS),
                                }
                            );
                        }

                        if (listDataConvert.Count > 0)
                        {
                            listReportData = new List<ReportDetailtForTotalMoneyType>(listDataConvert);
                        }
                    }
                    else
                    {
                        foreach (ReportDetailtForTotalMoneyType item in listReportData)
                        {
                            item.ReportID = item.PartnerName;
                            item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;
                        }
                    }

                    // Set from day and to day
                    sheetReport.Cells["E4"].PutValue(string.Format("{0}/{1}", fromDate.Month, fromDate.Year));
                    sheetReport.Cells["H4"].PutValue(string.Format("{0}/{1}", toDate.Month, toDate.Year));
                    break;
                // Theo năm
                default:

                    listReportData = new HSReportBL().SearchReportDetailtMTForOneForYear(fromDate, toDate, reportTypeID, marketID);

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
                                    TongDS = listDataItem.Sum(x => x.TongDS),
                                }
                            );
                        }

                        if (listDataConvert.Count > 0)
                        {
                            listReportData = new List<ReportDetailtForTotalMoneyType>(listDataConvert);
                        }
                    }
                    else
                    {
                        foreach (ReportDetailtForTotalMoneyType item in listReportData)
                        {
                            item.ReportID = item.PartnerName;
                            item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;
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
                if (marketID.Contains("005"))
                {
                    nameMarket = listReportData.Count > 0 ? listReportData[0].MarketName : string.Empty;
                    CreateTitle("A3", "K3", sheetReport, string.Format("Thị trường: {0}", nameMarket), 14);
                }
                else
                {
                    nameMarket = listReportData.Count > 0 ? listReportData[0].MarketName : string.Empty;
                    CreateTitle("A3", "K3", sheetReport, string.Format("Thị trường: {0}", nameMarket), 14);
                }
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
                dataTable = CreateDataTableFormart(true);

                // Danh sách dataSet của báo cáo ngày
                DataSet dataReport = ConvertListObjectToDataSet(listReportData);

                // Đổ data vào datatble mới
                FillData(dataReport.Tables[0], dataTable, true);
            }

            // Set border
            Style style = new CellsFactory().CreateStyle();
            style.SetBorder(BorderType.TopBorder, CellBorderType.Thin, Color.Black);
            style.SetBorder(BorderType.BottomBorder, CellBorderType.Thin, Color.Black);
            style.SetBorder(BorderType.LeftBorder, CellBorderType.Thin, Color.Black);
            style.SetBorder(BorderType.RightBorder, CellBorderType.Thin, Color.Black);
            
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
                            style.Custom = "#,##0";
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
            string templatePath = "~/Content/Report/ReportHSDetailt/ReportHSDetailLTGradationForGaration.xlsx";
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

            List<ReportDetailtForTotalMoneyType> listDataGradation = new HSReportBL().ReportDetailtMTGradationCompareForAll(year, int.Parse(gradationID), reportTypeID, marketID);

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


            // Danh sách mã thị trường của Tất cả
            List<string>  listMarket = new List<string>() { "003", "005", "001", "002", "014", "004" };

            // Trường hợp chọn tất cả thị trường
            if (marketID.Equals("0"))
            {
                foreach (ReportDetailtForTotalMoneyType item in listDataGradation)
                {
                    item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;
                }


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
                        style.Custom = "#,##0";

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
            listDataGradation = new HSReportBL().ReportDetailtMTGradationCompareForAll(year, int.Parse(gradationID), reportTypeID, marketID);

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
                        style.Custom = "#,##0";

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
            leadSourceColumnVND.Title.Text = string.Format("Hồ sơ loại tiền từng loại tiền - VND \n Giai đoạn: {0} \n", text);
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
            leadSourceColumnUSD.Title.Text = string.Format("Hồ sơ loại tiền từng loại tiền - USD \n Giai đoạn: {0} \n", text);
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
            leadSourceColumnEUR.Title.Text = string.Format("Hồ sơ loại tiền từng loại tiền - EUR \n Giai đoạn: {0} \n", text);
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
            leadSourceColumnCAD.Title.Text = string.Format("Hồ sơ loại tiền từng loại tiền - CAD \n Giai đoạn: {0} \n", text);
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
            leadSourceColumnAUD.Title.Text = string.Format("Hồ sơ loại tiền từng loại tiền - AUD \n Giai đoạn: {0} \n", text);
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
            leadSourceColumnGBP.Title.Text = string.Format("Hồ sơ loại tiền từng loại tiền - GBP \n Giai đoạn: {0} \n", text);
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

            listDataGradation = new HSReportBL().ReportDetailtMTGradationCompareForAllPercent(year, int.Parse(gradationID), reportTypeID, marketID);

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
            string templatePath = "~/Content/Report/ReportHSDetailt/ReportHSDetailLTGradationForGaration.xlsx";
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

            List<ReportDetailtForTotalMoneyType> listDataGradation = new HSReportBL().ReportDetailtMTGradationCompareForOne(year, int.Parse(gradationID), reportTypeID, marketID);

            string nameMarket = string.Empty;

            if (marketID.Contains("005"))
            {
                nameMarket = "Thị trường Châu Á";
                CreateTitle("A4", "K4", sheetReport, string.Format("Thị trường: {0}", nameMarket), 14);

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

            if(!marketID.Contains("005"))
            {
                nameMarket = listDataGradation.Count > 0 ? listDataGradation[0].MarketName : string.Empty;
                CreateTitle("A4", "K4", sheetReport, string.Format("Thị trường: {0}", nameMarket), 14);
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
                        style.Custom = "#,##0";

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


            listDataGradation = new HSReportBL().ReportDetailtMTGradationCompareForOne(year, int.Parse(gradationID), reportTypeID, marketID);

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
                        style.Custom = "#,##0";

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
            leadSourceColumnVND.Title.Text = string.Format("Hồ sơ loại tiền từng loại tiền - VND \n Giai đoạn: {0} \n", text);
            leadSourceColumnVND.Title.Font.Color = Color.Silver;

            // Adding SeriesCollection (chart data source) to the chart ranging from "A1" cell to "B3"
            string totalRowData = string.Format("B83:C{0}", 83 + table.Rows.Count - 2);
            leadSourceColumnVND.NSeries.Add(totalRowData, true);

            // Set the category data covering the range A2:A5.
            // Tổng số dòng cần hiển thị là số dòng hiện tại  + số dòng trong listData -2 (trong đó 1 là dòng tổng cuối cùng)
            string categoryData = "A83:A89";
            leadSourceColumnVND.NSeries.CategoryData = categoryData;
            // In nghiên cho title
            leadSourceColumnVND.CategoryAxis.TickLabels.RotationAngle = 15;

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
            leadSourceColumnUSD.Title.Text = string.Format("Hồ sơ loại tiền từng loại tiền - USD \n Giai đoạn: {0} \n", text);
            leadSourceColumnUSD.Title.Font.Color = Color.Silver;

            // Adding SeriesCollection (chart data source) to the chart ranging from "A1" cell to "B3"
            totalRowData = string.Format("D83:E{0}", 83 + table.Rows.Count - 2);
            leadSourceColumnUSD.NSeries.Add(totalRowData, true);

            // Set the category data covering the range A2:A5.
            // Tổng số dòng cần hiển thị là số dòng hiện tại  + số dòng trong listData -2 (trong đó 1 là dòng tổng cuối cùng)
            categoryData = "A83:A89";
            leadSourceColumnUSD.NSeries.CategoryData = categoryData;

            // In nghiên cho title
            leadSourceColumnUSD.CategoryAxis.TickLabels.RotationAngle = 15;

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
            leadSourceColumnEUR.Title.Text = string.Format("Hồ sơ loại tiền từng loại tiền - EUR \n Giai đoạn: {0} \n", text);
            leadSourceColumnEUR.Title.Font.Color = Color.Silver;

            // Adding SeriesCollection (chart data source) to the chart ranging from "A1" cell to "B3"
            totalRowData = string.Format("F83:G{0}", 83 + table.Rows.Count - 2);
            leadSourceColumnEUR.NSeries.Add(totalRowData, true);

            // Set the category data covering the range A2:A5.
            // Tổng số dòng cần hiển thị là số dòng hiện tại  + số dòng trong listData -2 (trong đó 1 là dòng tổng cuối cùng)
            categoryData = "A83:A89";
            leadSourceColumnEUR.NSeries.CategoryData = categoryData;

            // In nghiên cho title
            leadSourceColumnEUR.CategoryAxis.TickLabels.RotationAngle = 15;

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
            leadSourceColumnCAD.Title.Text = string.Format("Hồ sơ loại tiền từng loại tiền - CAD \n Giai đoạn: {0} \n", text);
            leadSourceColumnCAD.Title.Font.Color = Color.Silver;

            // Adding SeriesCollection (chart data source) to the chart ranging from "A1" cell to "B3"
            totalRowData = string.Format("H83:I{0}", 83 + table.Rows.Count - 2);
            leadSourceColumnCAD.NSeries.Add(totalRowData, true);

            // Set the category data covering the range A2:A5.
            // Tổng số dòng cần hiển thị là số dòng hiện tại  + số dòng trong listData -2 (trong đó 1 là dòng tổng cuối cùng)
            categoryData = "A83:A89";
            leadSourceColumnCAD.NSeries.CategoryData = categoryData;

            // In nghiên cho title
            leadSourceColumnCAD.CategoryAxis.TickLabels.RotationAngle = 15;

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
            leadSourceColumnAUD.Title.Text = string.Format("Hồ sơ loại tiền từng loại tiền - AUD \n Giai đoạn: {0} \n", text);
            leadSourceColumnAUD.Title.Font.Color = Color.Silver;

            // Adding SeriesCollection (chart data source) to the chart ranging from "A1" cell to "B3"
            totalRowData = string.Format("J83:K{0}", 83 + table.Rows.Count - 2);
            leadSourceColumnAUD.NSeries.Add(totalRowData, true);

            // Set the category data covering the range A2:A5.
            // Tổng số dòng cần hiển thị là số dòng hiện tại  + số dòng trong listData -2 (trong đó 1 là dòng tổng cuối cùng)
            categoryData = "A83:A89";
            leadSourceColumnAUD.NSeries.CategoryData = categoryData;

            // In nghiên cho title
            leadSourceColumnAUD.CategoryAxis.TickLabels.RotationAngle = 15;

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
            leadSourceColumnGBP.Title.Text = string.Format("Hồ sơ loại tiền từng loại tiền - GBP \n Giai đoạn: {0} \n", text);
            leadSourceColumnGBP.Title.Font.Color = Color.Silver;

            // Adding SeriesCollection (chart data source) to the chart ranging from "A1" cell to "B3"
            totalRowData = string.Format("L83:M{0}", 83 + table.Rows.Count - 2);
            leadSourceColumnGBP.NSeries.Add(totalRowData, true);

            // Set the category data covering the range A2:A5.
            // Tổng số dòng cần hiển thị là số dòng hiện tại  + số dòng trong listData -2 (trong đó 1 là dòng tổng cuối cùng)
            categoryData = "A83:A89";
            leadSourceColumnGBP.NSeries.CategoryData = categoryData;

            // In nghiên cho title
            leadSourceColumnGBP.CategoryAxis.TickLabels.RotationAngle = 15;

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

            listDataGradation = new HSReportBL().ReportDetailtMTGradationCompareForOnePercent(year, int.Parse(gradationID), reportTypeID, marketID);
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
            if (marketID == "005")
            {
                foreach (ReportDetailtForTotalMoneyType item in listDataGradation)
                {
                    item.PartnerName = item.MarketName;
                }
            }

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
        /// Tạo mẫu cho Excel cho so sánh theo tháng cho tất cả
        /// </summary>
        /// <param name="gradationID"></param>
        /// <param name="year"></param>
        /// <param name="typeID"></param>
        /// <returns></returns>
        public ActionResult CreateExcelForCompareForMonth(int year, int month, string reportTypeID, string marketID)
        {
            WorkbookDesigner designer = new WorkbookDesigner();
            string templatePath = "~/Content/Report/ReportHSDetailt/ReportHSDetailtLHMoneyTypeCompare.xlsx";
            // Get đường dẫn
            templatePath = System.Web.HttpContext.Current.Server.MapPath(templatePath);

            designer.Workbook = new DongA.Core.DongAExcel().OpenExcelFile(templatePath);
            designer.Workbook.CalculateFormula();

            WorksheetCollection workSheets = designer.Workbook.Worksheets;
            Worksheet sheetReport = designer.Workbook.Worksheets[0];

            // Tạo title
            string typeReport = "So sánh - Theo tháng - Tất cả";
            // Tạo title
            CreateTitle("A2", "P2", sheetReport, typeReport, 14);

            // Tạo title detailt
            string titleDetailt = string.Format("Tháng {0}/{1}", month, year);
            CreateTitle("A3", "P3", sheetReport, titleDetailt, 12);


            // Tạo giá trị cho cột dữ liệu của Chi quầy/ Chi nhà/ Chuyển khoản
            sheetReport.Cells["B102"].PutValue(string.Format("Tháng {0}/{1} ", month, year - 1));
            
            if(month == 1)
            {
                sheetReport.Cells["C102"].PutValue(string.Format("Tháng {0}/{1} ", 12, year - 1));
            }
            sheetReport.Cells["D102"].PutValue(string.Format("Tháng {0}/{1} ", month, year));

            sheetReport.Cells["E102"].PutValue(string.Format("Tháng {0}/{1} ", month, year - 1));
            sheetReport.Cells["F102"].PutValue(string.Format("Tháng {0}/{1} ", month - 1, year));

            if (month == 1)
            {
                sheetReport.Cells["F102"].PutValue(string.Format("Tháng {0}/{1} ", 12, year - 1));
            }
            sheetReport.Cells["G102"].PutValue(string.Format("Tháng {0}/{1} ", month, year));

            sheetReport.Cells["H102"].PutValue(string.Format("Tháng {0}/{1} ", month, year - 1));
            sheetReport.Cells["I102"].PutValue(string.Format("Tháng {0}/{1} ", month - 1, year));

            if (month == 1)
            {
                sheetReport.Cells["I102"].PutValue(string.Format("Tháng {0}/{1} ", 12, year - 1));
            }
            sheetReport.Cells["J102"].PutValue(string.Format("Tháng {0}/{1} ", month, year));

            sheetReport.Cells["K102"].PutValue(string.Format("Tháng {0}/{1} ", month, year - 1));
            sheetReport.Cells["L102"].PutValue(string.Format("Tháng {0}/{1} ", month - 1, year));

            if (month == 1)
            {
                sheetReport.Cells["L102"].PutValue(string.Format("Tháng {0}/{1} ", 12, year - 1));
            }
            sheetReport.Cells["M102"].PutValue(string.Format("Tháng {0}/{1} ", month, year));

            sheetReport.Cells["N102"].PutValue(string.Format("Tháng {0}/{1} ", month, year - 1));
            sheetReport.Cells["O102"].PutValue(string.Format("Tháng {0}/{1} ", month - 1, year));

            if (month == 1)
            {
                sheetReport.Cells["O102"].PutValue(string.Format("Tháng {0}/{1} ", 12, year - 1));
            }
            sheetReport.Cells["P102"].PutValue(string.Format("Tháng {0}/{1} ", month, year));

            sheetReport.Cells["Q102"].PutValue(string.Format("Tháng {0}/{1} ", month, year - 1));
            sheetReport.Cells["R102"].PutValue(string.Format("Tháng {0}/{1} ", month - 1, year));

            if (month == 1)
            {
                sheetReport.Cells["R102"].PutValue(string.Format("Tháng {0}/{1} ", 12, year - 1));
            }
            sheetReport.Cells["S102"].PutValue(string.Format("Tháng {0}/{1} ", month, year));

            sheetReport.Cells["T102"].PutValue(string.Format("Tháng {0}/{1} ", month, year - 1));
            sheetReport.Cells["U102"].PutValue(string.Format("Tháng {0}/{1} ", month - 1, year));

            if (month == 1)
            {
                sheetReport.Cells["U102"].PutValue(string.Format("Tháng {0}/{1} ", 12, year - 1));
            }
            sheetReport.Cells["V102"].PutValue(string.Format("Tháng {0}/{1} ", month, year));

            List<ReportDetailtForTotalMoneyType> listDataCompareMonth = new HSReportBL().ReportDetailtMTCompareMonthForAll(year, month, reportTypeID, marketID);

            foreach (ReportDetailtForTotalMoneyType item in listDataCompareMonth)
            {
                item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;
            }

            // Danh sách mã thị trường của Tất cả
            List<string> listMarket = new List<string>() { "003", "005", "001", "002", "014", "004" };

            // Khởi tạo datatable
            DataTable table = new DataTable();
            // Tạo các cột cho datatable
            table.Columns.Add("MarketName", typeof(String));
            table.Columns.Add("VND1", typeof(double));
            table.Columns.Add("VND2", typeof(double));
            table.Columns.Add("VND3", typeof(double));

            table.Columns.Add("USD1", typeof(double));
            table.Columns.Add("USD2", typeof(double));
            table.Columns.Add("USD3", typeof(double));

            table.Columns.Add("EUR1", typeof(double));
            table.Columns.Add("EUR2", typeof(double));
            table.Columns.Add("EUR3", typeof(double));

            table.Columns.Add("CAD1", typeof(double));
            table.Columns.Add("CAD2", typeof(double));
            table.Columns.Add("CAD3", typeof(double));

            table.Columns.Add("AUD1", typeof(double));
            table.Columns.Add("AUD2", typeof(double));
            table.Columns.Add("AUD3", typeof(double));

            table.Columns.Add("GBP1", typeof(double));
            table.Columns.Add("GBP2", typeof(double));
            table.Columns.Add("GBP3", typeof(double));

            table.Columns.Add("TDS1", typeof(double));
            table.Columns.Add("TDS2", typeof(double));
            table.Columns.Add("TDS3", typeof(double));


            table.Columns.Add("ReportID", typeof(String));

            foreach (ReportDetailtForTotalMoneyType item in listDataCompareMonth)
            {
                ReportDetailtForTotalMoneyType dataItemLastYear = listDataCompareMonth.Find(x => x.MarketCode == item.MarketCode && x.Month == month.ToString() && x.Year == (year - 1).ToString());
                ReportDetailtForTotalMoneyType dataItemYear = listDataCompareMonth.Find(x => x.MarketCode == item.MarketCode && x.Month == month.ToString() && x.Year == year.ToString());
                ReportDetailtForTotalMoneyType dataItemLastMonth = listDataCompareMonth.Find(x => x.MarketCode == item.MarketCode && x.Month == (month - 1).ToString() && x.Year == year.ToString());

                // Trường hợp tháng 1
                if (month == 1)
                {
                    dataItemLastMonth = listDataCompareMonth.Find(x => x.MarketCode == item.MarketCode && x.Month == "12" && x.Year == (year - 1).ToString());
                }

                // Check tồn tại của item
                string value = string.Format("MarketName='{0}'", item.MarketName);
                DataRow[] foundRows = table.Select(value);

                if (dataItemLastYear != null && dataItemYear != null && dataItemLastMonth != null && foundRows.Count() == 0)
                {
                    // add item vào table
                    table.Rows.Add(dataItemYear.MarketName
                        , dataItemLastYear.VND, dataItemLastMonth.VND, dataItemYear.VND
                        , dataItemLastYear.USD, dataItemLastMonth.USD, dataItemYear.USD
                        , dataItemLastYear.EUR, dataItemLastMonth.EUR, dataItemYear.EUR
                        , dataItemLastYear.CAD, dataItemLastMonth.CAD, dataItemYear.CAD
                        , dataItemLastYear.AUD, dataItemLastMonth.AUD, dataItemYear.AUD
                        , dataItemLastYear.GBP, dataItemLastMonth.GBP, dataItemYear.GBP
                        , dataItemLastYear.TongDS, dataItemLastMonth.TongDS, dataItemYear.TongDS
                        );
                }
            }
            // Cùng kì


            DataRow row = table.NewRow();
            row["MarketName"] = "Tổng";
            row["VND1"] = table.Compute("Sum(VND1)", "");
            row["VND2"] = table.Compute("Sum(VND2)", "");
            row["VND3"] = table.Compute("Sum(VND3)", "");

            row["USD1"] = table.Compute("Sum(USD1)", "");
            row["USD2"] = table.Compute("Sum(USD2)", "");
            row["USD3"] = table.Compute("Sum(USD3)", "");

            row["EUR1"] = table.Compute("Sum(EUR1)", "");
            row["EUR2"] = table.Compute("Sum(EUR2)", "");
            row["EUR3"] = table.Compute("Sum(EUR3)", "");

            row["CAD1"] = table.Compute("Sum(CAD1)", "");
            row["CAD2"] = table.Compute("Sum(CAD2)", "");
            row["CAD3"] = table.Compute("Sum(CAD3)", "");

            row["AUD1"] = table.Compute("Sum(AUD1)", "");
            row["AUD2"] = table.Compute("Sum(AUD2)", "");
            row["AUD3"] = table.Compute("Sum(AUD3)", "");

            row["GBP1"] = table.Compute("Sum(GBP1)", "");
            row["GBP2"] = table.Compute("Sum(GBP2)", "");
            row["GBP3"] = table.Compute("Sum(GBP3)", "");

            row["TDS1"] = table.Compute("Sum(TDS1)", "");
            row["TDS2"] = table.Compute("Sum(TDS2)", "");
            row["TDS3"] = table.Compute("Sum(TDS3)", "");

            table.Rows.Add(row);

            // Tổng số row theo table1
            int totalRowTable1 = table.Rows.Count + 102;

            // Set border
            Style style = new CellsFactory().CreateStyle();
            style.SetBorder(BorderType.BottomBorder, CellBorderType.Thin, Color.Black);
            style.SetBorder(BorderType.LeftBorder, CellBorderType.Thin, Color.Black);
            style.SetBorder(BorderType.RightBorder, CellBorderType.Thin, Color.Black);

            if (table.Rows.Count > 0)
            {
                int stepRow = 0;
                // total row = row start + số row hiện có
                int totalRow = table.Rows.Count + 102;
                // Số dòng của row
                for (int a = 102; a < totalRow; a++)
                {
                    int stepColumn = 0;
                    // Số cột trong báo cáo cần hiển thị
                    // Tổng số cột hiển thị = Số cột hiển thị bắt đầu + tổng số cột cần hiển thị
                    int totalCol = 0 + 22;
                    for (int b = 0; b < totalCol; b++)
                    {
                        // Giá trị của value trong table
                        string valueOfTable = table.Rows[stepRow][stepColumn].ToString();

                        //// Tô màu cho các dòng có giá trị tăng giảm
                        //if (b >= 18)
                        //{
                        //    decimal tryParseValue = 0;
                        //    decimal.TryParse(valueOfTable, out tryParseValue);
                        //    style.Font.Color = Color.Green;

                        //    if (tryParseValue < 0)
                        //    {
                        //        style.Font.Color = Color.Red;
                        //    }
                        //}

                        // Insert vào dòng cột xác định trong Excel
                        sheetReport.Cells[a, b].PutValue(valueOfTable, true);

                        // set style cho number
                        style.Custom = "#,##0";

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

            listDataCompareMonth = new HSReportBL().ReportDetailtMTCompareMonthForAll(year, month, reportTypeID, marketID);

            foreach (ReportDetailtForTotalMoneyType item in listDataCompareMonth)
            {
                item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;
            }

            // Khởi tạo datatable
            table = new DataTable();
            // Tạo các cột cho datatable
            table.Columns.Add("MarketName", typeof(String));

            // So sánh với tháng trước
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

            foreach (ReportDetailtForTotalMoneyType item in listDataCompareMonth)
            {
                // Cùng kì
                ReportDetailtForTotalMoneyType dataItemLastYear = listDataCompareMonth.Find(x => x.MarketCode == item.MarketCode && x.Month == month.ToString() && x.Year == (year - 1).ToString());
                ReportDetailtForTotalMoneyType dataItemYear = listDataCompareMonth.Find(x => x.MarketCode == item.MarketCode && x.Month == month.ToString() && x.Year == year.ToString());
                ReportDetailtForTotalMoneyType dataItemLastMonth = listDataCompareMonth.Find(x => x.MarketCode == item.MarketCode && x.Month == (month - 1).ToString() && x.Year == year.ToString());
                // Trường hợp tháng 1
                if (month == 1)
                {
                    dataItemLastMonth = listDataCompareMonth.Find(x => x.MarketCode == item.MarketCode && x.Month == "12" && x.Year == (year - 1).ToString());
                }

                // Cùng kì năm trước
                if (dataItemLastYear == null)
                {
                    dataItemLastYear.MarketCode = item.MarketCode;
                    dataItemLastYear.MarketName = item.MarketName;
                }

                // Tháng hiện tại
                if (dataItemYear == null)
                {
                    dataItemYear.MarketCode = item.MarketCode;
                    dataItemYear.MarketName = item.MarketName;
                }

                // Tháng trước
                if (dataItemLastMonth == null)
                {
                    dataItemLastMonth.MarketCode = item.MarketCode;
                    dataItemLastMonth.MarketName = item.MarketName;
                }

                // So với cùng kì năm trước
                double sumVNDLastYear = Math.Round(dataItemYear.VND - dataItemLastYear.VND, 2, MidpointRounding.ToEven);
                double sumUSDLastYear = Math.Round(dataItemYear.USD - dataItemLastYear.USD, 2, MidpointRounding.ToEven);
                double sumEURLastYear = Math.Round(dataItemYear.EUR - dataItemLastYear.EUR, 2, MidpointRounding.ToEven);
                double sumCADLastYear = Math.Round(dataItemYear.CAD - dataItemLastYear.CAD, 2, MidpointRounding.ToEven);
                double sumAUDLastYear = Math.Round(dataItemYear.AUD - dataItemLastYear.AUD, 2, MidpointRounding.ToEven);
                double sumGBPLastYear = Math.Round(dataItemYear.GBP - dataItemLastYear.GBP, 2, MidpointRounding.ToEven);

                double sumTDSLastYear = sumVNDLastYear + sumUSDLastYear + sumEURLastYear + sumCADLastYear + sumAUDLastYear + sumGBPLastYear;

                // so với tháng trước
                double sumVNDLastMonth = Math.Round(dataItemYear.VND - dataItemLastMonth.VND, 2, MidpointRounding.ToEven);
                double sumUSDLastMonth = Math.Round(dataItemYear.USD - dataItemLastMonth.USD, 2, MidpointRounding.ToEven);
                double sumEURLastMonth = Math.Round(dataItemYear.EUR - dataItemLastMonth.EUR, 2, MidpointRounding.ToEven);
                double sumCADLastMonth = Math.Round(dataItemYear.CAD - dataItemLastMonth.CAD, 2, MidpointRounding.ToEven);
                double sumAUDLastMonth = Math.Round(dataItemYear.AUD - dataItemLastMonth.AUD, 2, MidpointRounding.ToEven);
                double sumGBPLastMonth = Math.Round(dataItemYear.GBP - dataItemLastMonth.GBP, 2, MidpointRounding.ToEven);

                double sumTDSLastMonth = sumVNDLastMonth + sumUSDLastMonth + sumEURLastMonth + sumCADLastMonth + sumAUDLastMonth + sumGBPLastMonth;

                // Check tồn tại của item
                string value = string.Format("MarketName='{0}'", item.MarketName);
                DataRow[] foundRows = table.Select(value);

                if (dataItemLastYear != null && dataItemYear != null && dataItemLastMonth != null && foundRows.Count() == 0)
                {
                    // add item vào table
                    table.Rows.Add(dataItemYear.MarketName
                        , sumVNDLastMonth, sumVNDLastYear
                        , sumUSDLastMonth, sumUSDLastYear
                        , sumEURLastMonth, sumEURLastYear
                        , sumCADLastMonth, sumCADLastYear
                        , sumAUDLastMonth, sumAUDLastYear
                        , sumGBPLastMonth, sumGBPLastYear
                        , Math.Round(sumTDSLastMonth, 2, MidpointRounding.ToEven), Math.Round(sumTDSLastYear, 2, MidpointRounding.ToEven)
                        );
                }
            }

            row = table.NewRow();
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


            // Tổng số row của table2
            // Với 6 là số cách của table1 và table2
            int totalRowTable2 = totalRowTable1 + table.Rows.Count + 6;

            // Tạo title hearder cho table tăng giảm
            // Title cho thị trường
            string title = "Thị trường";

            CreateTitle(string.Format("A{0}", totalRowTable1 + 6 - 1), string.Format("A{0}", totalRowTable1 + 6), sheetReport, title, 12, true);

            title = "VND";
            CreateTitle(string.Format("B{0}", totalRowTable1 + 6 - 1), string.Format("C{0}", totalRowTable1 + 6 - 1), sheetReport, title, 12, true);

            title = "USD";
            CreateTitle(string.Format("D{0}", totalRowTable1 + 6 - 1), string.Format("E{0}", totalRowTable1 + 6 - 1), sheetReport, title, 12, true);

            title = "EUR";
            CreateTitle(string.Format("F{0}", totalRowTable1 + 6 - 1), string.Format("G{0}", totalRowTable1 + 6 - 1), sheetReport, title, 12, true);

            title = "CAD";
            CreateTitle(string.Format("H{0}", totalRowTable1 + 6 - 1), string.Format("I{0}", totalRowTable1 + 6 - 1), sheetReport, title, 12, true);

            title = "AUD";
            CreateTitle(string.Format("J{0}", totalRowTable1 + 6 - 1), string.Format("K{0}", totalRowTable1 + 6 - 1), sheetReport, title, 12, true);

            title = "GBP";
            CreateTitle(string.Format("L{0}", totalRowTable1 + 6 - 1), string.Format("M{0}", totalRowTable1 + 6 - 1), sheetReport, title, 12, true);

            title = "Tổng";
            CreateTitle(string.Format("N{0}", totalRowTable1 + 6 - 1), string.Format("O{0}", totalRowTable1 + 6 - 1), sheetReport, title, 12, true);


            title = "So với tháng trước";
            CreateTitle(string.Format("B{0}", totalRowTable1 + 6), string.Format("B{0}", totalRowTable1 + 6), sheetReport, title, 12, true);

            title = string.Format("So với cùng kì");
            CreateTitle(string.Format("C{0}", totalRowTable1 + 6), string.Format("C{0}", totalRowTable1 + 6), sheetReport, title, 12, true);

            title = "So với tháng trước";
            CreateTitle(string.Format("D{0}", totalRowTable1 + 6), string.Format("D{0}", totalRowTable1 + 6), sheetReport, title, 12, true);

            title = string.Format("So với cùng kì");
            CreateTitle(string.Format("E{0}", totalRowTable1 + 6), string.Format("E{0}", totalRowTable1 + 6), sheetReport, title, 12, true);

            title = "So với tháng trước";
            CreateTitle(string.Format("F{0}", totalRowTable1 + 6), string.Format("F{0}", totalRowTable1 + 6), sheetReport, title, 12, true);

            title = string.Format("So với cùng kì");
            CreateTitle(string.Format("G{0}", totalRowTable1 + 6), string.Format("G{0}", totalRowTable1 + 6), sheetReport, title, 12, true);

            title = "So với tháng trước";
            CreateTitle(string.Format("H{0}", totalRowTable1 + 6), string.Format("H{0}", totalRowTable1 + 6), sheetReport, title, 12, true);

            title = string.Format("So với cùng kì");
            CreateTitle(string.Format("I{0}", totalRowTable1 + 6), string.Format("I{0}", totalRowTable1 + 6), sheetReport, title, 12, true);

            title = "So với tháng trước";
            CreateTitle(string.Format("J{0}", totalRowTable1 + 6), string.Format("J{0}", totalRowTable1 + 6), sheetReport, title, 12, true);

            title = string.Format("So với cùng kì");
            CreateTitle(string.Format("K{0}", totalRowTable1 + 6), string.Format("K{0}", totalRowTable1 + 6), sheetReport, title, 12, true);

            title = "So với tháng trước";
            CreateTitle(string.Format("L{0}", totalRowTable1 + 6), string.Format("L{0}", totalRowTable1 + 6), sheetReport, title, 12, true);

            title = string.Format("So với cùng kì");
            CreateTitle(string.Format("M{0}", totalRowTable1 + 6), string.Format("M{0}", totalRowTable1 + 6), sheetReport, title, 12, true);

            title = "So với tháng trước";
            CreateTitle(string.Format("N{0}", totalRowTable1 + 6), string.Format("N{0}", totalRowTable1 + 6), sheetReport, title, 12, true);

            title = string.Format("So với cùng kì");
            CreateTitle(string.Format("O{0}", totalRowTable1 + 6), string.Format("O{0}", totalRowTable1 + 6), sheetReport, title, 12, true);

            // Table dữ liệu bảng số liệu Hồ sơ Chi Quầy/Chi Nhà/Chuyển Khoản
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
                    int totalCol = 0 + 15;
                    for (int b = 0; b < totalCol; b++)
                    {
                        // Giá trị của value trong table
                        string valueOfTable = table.Rows[stepRow][stepColumn].ToString();

                        // Insert vào dòng cột xác định trong Excel
                        sheetReport.Cells[a, b].PutValue(valueOfTable, true);

                        // Tô màu cho các dòng có giá trị tăng giảm
                        if (b >= 1)
                        {
                            decimal tryParseValue = 0;
                            decimal.TryParse(valueOfTable, out tryParseValue);
                            style.Font.Color = Color.Green;

                            if (tryParseValue < 0)
                            {
                                style.Font.Color = Color.Red;
                            }
                        }

                        // set style cho number
                        style.Custom = "#,##0";

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
            int chartIndex = sheetReport.Charts.Add(ChartType.Bar, 7, 0, 27, 6);
            leadSourceColumnVND = sheetReport.Charts[chartIndex];

            //Chart title
            leadSourceColumnVND.Title.Text = string.Format("Hồ sơ từng loại tiền từng thị trường - VND \n Tháng {0}/{1}", month, year);
            leadSourceColumnVND.Title.Font.Color = Color.Silver;

            // Adding SeriesCollection (chart data source) to the chart ranging from "A1" cell to "B3"
            string totalRowData = string.Format("B103:D{0}", totalRowTable1 - 1);
            leadSourceColumnVND.NSeries.Add(totalRowData, true);

            // Set the category data covering the range A2:A5.
            // Tổng số dòng cần hiển thị là số dòng hiện tại  + số dòng trong listData -2 (trong đó 1 là dòng tổng cuối cùng)
            string categoryData = "A103:A108";
            leadSourceColumnVND.NSeries.CategoryData = categoryData;

            // Set the names of the chart series taken from cells.
            leadSourceColumnVND.NSeries[0].Name = "=B102";
            leadSourceColumnVND.NSeries[1].Name = "=C102";
            leadSourceColumnVND.NSeries[2].Name = "=D102";

            // Set the 1st series fill color.
            leadSourceColumnVND.NSeries[0].Area.ForegroundColor = Color.Orange;
            leadSourceColumnVND.NSeries[0].Area.Formatting = FormattingType.Custom;

            // Set the 2nd series fill color.
            leadSourceColumnVND.NSeries[1].Area.ForegroundColor = Color.Green;
            leadSourceColumnVND.NSeries[1].Area.Formatting = FormattingType.Custom;

            // Set the 2nd series fill color.
            leadSourceColumnVND.NSeries[2].Area.ForegroundColor = Color.Blue;
            leadSourceColumnVND.NSeries[2].Area.Formatting = FormattingType.Custom;


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

            // Vẽ biểu đồ cột cho Excel
            //Chart reference
            Aspose.Cells.Charts.Chart leadSourceColumnUSD;
            //Add Pie Chart
            // USD
            chartIndex = sheetReport.Charts.Add(ChartType.Bar, 7, 7, 27, 13);
            leadSourceColumnUSD = sheetReport.Charts[chartIndex];

            //Chart title
            leadSourceColumnUSD.Title.Text = string.Format("Hồ sơ từng loại tiền từng thị trường - USD \n Tháng {0}/{1}", month, year);
            leadSourceColumnUSD.Title.Font.Color = Color.Silver;

            // Adding SeriesCollection (chart data source) to the chart ranging from "A1" cell to "B3"
            totalRowData = string.Format("E103:G{0}", totalRowTable1 - 1);
            leadSourceColumnUSD.NSeries.Add(totalRowData, true);

            // Set the category data covering the range A2:A5.
            // Tổng số dòng cần hiển thị là số dòng hiện tại  + số dòng trong listData -2 (trong đó 1 là dòng tổng cuối cùng)
            categoryData = "A103:A108";
            leadSourceColumnUSD.NSeries.CategoryData = categoryData;

            // Set the names of the chart series taken from cells.
            leadSourceColumnUSD.NSeries[0].Name = "=B102";
            leadSourceColumnUSD.NSeries[1].Name = "=C102";
            leadSourceColumnUSD.NSeries[2].Name = "=D102";

            // Set the 1st series fill color.
            leadSourceColumnUSD.NSeries[0].Area.ForegroundColor = Color.Orange;
            leadSourceColumnUSD.NSeries[0].Area.Formatting = FormattingType.Custom;

            // Set the 2nd series fill color.
            leadSourceColumnUSD.NSeries[1].Area.ForegroundColor = Color.Green;
            leadSourceColumnUSD.NSeries[1].Area.Formatting = FormattingType.Custom;

            // Set the 2nd series fill color.
            leadSourceColumnUSD.NSeries[2].Area.ForegroundColor = Color.Blue;
            leadSourceColumnUSD.NSeries[2].Area.Formatting = FormattingType.Custom;


            // Set plot area formatting as none and hide its border.
            leadSourceColumnUSD.PlotArea.Area.FillFormat.FillType = FillType.None;
            leadSourceColumnUSD.PlotArea.Border.IsVisible = false;

            // Set value axis major tick mark as none and hide axis line. 
            // Also set the color of value axis major grid lines.
            leadSourceColumnUSD.ValueAxis.MajorTickMark = TickMarkType.None;
            leadSourceColumnUSD.ValueAxis.AxisLine.IsVisible = false;

            leadSourceColumnUSD.ValueAxis.MajorGridLines.Color = Color.FromArgb(217, 217, 217);

            // Vẽ biểu đồ cột cho Excel
            //Chart reference
            Aspose.Cells.Charts.Chart leadSourceColumnEUR;
            //Add Pie Chart
            // EUR
            chartIndex = sheetReport.Charts.Add(ChartType.Bar, 7, 14, 27, 20);
            leadSourceColumnEUR = sheetReport.Charts[chartIndex];

            //Chart title
            leadSourceColumnEUR.Title.Text = string.Format("Hồ sơ từng loại tiền từng thị trường - EUR \n Tháng {0}/{1}", month, year);
            leadSourceColumnEUR.Title.Font.Color = Color.Silver;

            // Adding SeriesCollection (chart data source) to the chart ranging from "A1" cell to "B3"
            totalRowData = string.Format("H103:J{0}", totalRowTable1 - 1);
            leadSourceColumnEUR.NSeries.Add(totalRowData, true);

            // Set the category data covering the range A2:A5.
            // Tổng số dòng cần hiển thị là số dòng hiện tại  + số dòng trong listData -2 (trong đó 1 là dòng tổng cuối cùng)
            categoryData = "A103:A108";
            leadSourceColumnEUR.NSeries.CategoryData = categoryData;

            // Set the names of the chart series taken from cells.
            leadSourceColumnEUR.NSeries[0].Name = "=B102";
            leadSourceColumnEUR.NSeries[1].Name = "=C102";
            leadSourceColumnEUR.NSeries[2].Name = "=D102";

            // Set the 1st series fill color.
            leadSourceColumnEUR.NSeries[0].Area.ForegroundColor = Color.Orange;
            leadSourceColumnEUR.NSeries[0].Area.Formatting = FormattingType.Custom;

            // Set the 2nd series fill color.
            leadSourceColumnEUR.NSeries[1].Area.ForegroundColor = Color.Green;
            leadSourceColumnEUR.NSeries[1].Area.Formatting = FormattingType.Custom;

            // Set the 2nd series fill color.
            leadSourceColumnEUR.NSeries[2].Area.ForegroundColor = Color.Blue;
            leadSourceColumnEUR.NSeries[2].Area.Formatting = FormattingType.Custom;


            // Set plot area formatting as none and hide its border.
            leadSourceColumnEUR.PlotArea.Area.FillFormat.FillType = FillType.None;
            leadSourceColumnEUR.PlotArea.Border.IsVisible = false;

            // Set value axis major tick mark as none and hide axis line. 
            // Also set the color of value axis major grid lines.
            leadSourceColumnEUR.ValueAxis.MajorTickMark = TickMarkType.None;
            leadSourceColumnEUR.ValueAxis.AxisLine.IsVisible = false;

            leadSourceColumnEUR.ValueAxis.MajorGridLines.Color = Color.FromArgb(217, 217, 217);

            // Vẽ biểu đồ cột cho Excel
            //Chart reference
            Aspose.Cells.Charts.Chart leadSourceColumnCAD;
            //Add Pie Chart
            // EUR
            chartIndex = sheetReport.Charts.Add(ChartType.Bar, 29, 0, 49, 6);
            leadSourceColumnCAD = sheetReport.Charts[chartIndex];

            //Chart title
            leadSourceColumnCAD.Title.Text = string.Format("Hồ sơ từng loại tiền từng thị trường - CAD \n Tháng {0}/{1}", month, year);
            leadSourceColumnCAD.Title.Font.Color = Color.Silver;

            // Adding SeriesCollection (chart data source) to the chart ranging from "A1" cell to "B3"
            totalRowData = string.Format("K103:M{0}", totalRowTable1 - 1);
            leadSourceColumnCAD.NSeries.Add(totalRowData, true);

            // Set the category data covering the range A2:A5.
            // Tổng số dòng cần hiển thị là số dòng hiện tại  + số dòng trong listData -2 (trong đó 1 là dòng tổng cuối cùng)
            categoryData = "A103:A108";
            leadSourceColumnCAD.NSeries.CategoryData = categoryData;

            // Set the names of the chart series taken from cells.
            leadSourceColumnCAD.NSeries[0].Name = "=B102";
            leadSourceColumnCAD.NSeries[1].Name = "=C102";
            leadSourceColumnCAD.NSeries[2].Name = "=D102";

            // Set the 1st series fill color.
            leadSourceColumnCAD.NSeries[0].Area.ForegroundColor = Color.Orange;
            leadSourceColumnCAD.NSeries[0].Area.Formatting = FormattingType.Custom;

            // Set the 2nd series fill color.
            leadSourceColumnCAD.NSeries[1].Area.ForegroundColor = Color.Green;
            leadSourceColumnCAD.NSeries[1].Area.Formatting = FormattingType.Custom;

            // Set the 2nd series fill color.
            leadSourceColumnCAD.NSeries[2].Area.ForegroundColor = Color.Blue;
            leadSourceColumnCAD.NSeries[2].Area.Formatting = FormattingType.Custom;


            // Set plot area formatting as none and hide its border.
            leadSourceColumnCAD.PlotArea.Area.FillFormat.FillType = FillType.None;
            leadSourceColumnCAD.PlotArea.Border.IsVisible = false;

            // Set value axis major tick mark as none and hide axis line. 
            // Also set the color of value axis major grid lines.
            leadSourceColumnCAD.ValueAxis.MajorTickMark = TickMarkType.None;
            leadSourceColumnCAD.ValueAxis.AxisLine.IsVisible = false;

            leadSourceColumnCAD.ValueAxis.MajorGridLines.Color = Color.FromArgb(217, 217, 217);

            // Vẽ biểu đồ cột cho Excel
            //Chart reference
            Aspose.Cells.Charts.Chart leadSourceColumnAUD;
            //Add Pie Chart
            // AUD
            chartIndex = sheetReport.Charts.Add(ChartType.Bar, 29, 7, 49, 13);
            leadSourceColumnAUD = sheetReport.Charts[chartIndex];

            //Chart title
            leadSourceColumnAUD.Title.Text = string.Format("Hồ sơ từng loại tiền từng thị trường - AUD \n Tháng {0}/{1}", month, year);
            leadSourceColumnAUD.Title.Font.Color = Color.Silver;

            // Adding SeriesCollection (chart data source) to the chart ranging from "A1" cell to "B3"
            totalRowData = string.Format("N103:P{0}", totalRowTable1 - 1);
            leadSourceColumnAUD.NSeries.Add(totalRowData, true);

            // Set the category data covering the range A2:A5.
            // Tổng số dòng cần hiển thị là số dòng hiện tại  + số dòng trong listData -2 (trong đó 1 là dòng tổng cuối cùng)
            categoryData = "A103:A108";
            leadSourceColumnAUD.NSeries.CategoryData = categoryData;

            // Set the names of the chart series taken from cells.
            leadSourceColumnAUD.NSeries[0].Name = "=B102";
            leadSourceColumnAUD.NSeries[1].Name = "=C102";
            leadSourceColumnAUD.NSeries[2].Name = "=D102";

            // Set the 1st series fill color.
            leadSourceColumnAUD.NSeries[0].Area.ForegroundColor = Color.Orange;
            leadSourceColumnAUD.NSeries[0].Area.Formatting = FormattingType.Custom;

            // Set the 2nd series fill color.
            leadSourceColumnAUD.NSeries[1].Area.ForegroundColor = Color.Green;
            leadSourceColumnAUD.NSeries[1].Area.Formatting = FormattingType.Custom;

            // Set the 2nd series fill color.
            leadSourceColumnAUD.NSeries[2].Area.ForegroundColor = Color.Blue;
            leadSourceColumnAUD.NSeries[2].Area.Formatting = FormattingType.Custom;


            // Set plot area formatting as none and hide its border.
            leadSourceColumnAUD.PlotArea.Area.FillFormat.FillType = FillType.None;
            leadSourceColumnAUD.PlotArea.Border.IsVisible = false;

            // Set value axis major tick mark as none and hide axis line. 
            // Also set the color of value axis major grid lines.
            leadSourceColumnAUD.ValueAxis.MajorTickMark = TickMarkType.None;
            leadSourceColumnAUD.ValueAxis.AxisLine.IsVisible = false;

            leadSourceColumnAUD.ValueAxis.MajorGridLines.Color = Color.FromArgb(217, 217, 217);

            // Vẽ biểu đồ cột cho Excel
            //Chart reference
            Aspose.Cells.Charts.Chart leadSourceColumnGBP;
            //Add Pie Chart
            // GBP
            chartIndex = sheetReport.Charts.Add(ChartType.Bar, 29, 14, 49, 20);
            leadSourceColumnGBP = sheetReport.Charts[chartIndex];

            //Chart title
            leadSourceColumnGBP.Title.Text = string.Format("Hồ sơ từng loại tiền từng thị trường - GBP \n Tháng {0}/{1}", month, year);
            leadSourceColumnGBP.Title.Font.Color = Color.Silver;

            // Adding SeriesCollection (chart data source) to the chart ranging from "A1" cell to "B3"
            totalRowData = string.Format("Q103:S{0}", totalRowTable1 - 1);
            leadSourceColumnGBP.NSeries.Add(totalRowData, true);

            // Set the category data covering the range A2:A5.
            // Tổng số dòng cần hiển thị là số dòng hiện tại  + số dòng trong listData -2 (trong đó 1 là dòng tổng cuối cùng)
            categoryData = "A103:A88";
            leadSourceColumnGBP.NSeries.CategoryData = categoryData;

            // Set the names of the chart series taken from cells.
            leadSourceColumnGBP.NSeries[0].Name = "=B102";
            leadSourceColumnGBP.NSeries[1].Name = "=C102";
            leadSourceColumnGBP.NSeries[2].Name = "=D102";

            // Set the 1st series fill color.
            leadSourceColumnGBP.NSeries[0].Area.ForegroundColor = Color.Orange;
            leadSourceColumnGBP.NSeries[0].Area.Formatting = FormattingType.Custom;

            // Set the 2nd series fill color.
            leadSourceColumnGBP.NSeries[1].Area.ForegroundColor = Color.Green;
            leadSourceColumnGBP.NSeries[1].Area.Formatting = FormattingType.Custom;

            // Set the 2nd series fill color.
            leadSourceColumnGBP.NSeries[2].Area.ForegroundColor = Color.Blue;
            leadSourceColumnGBP.NSeries[2].Area.Formatting = FormattingType.Custom;


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
            chartIndex = sheetReport.Charts.Add(ChartType.Bar3DStacked, 51, 0, 98, 13);
            leadSourceColumn = sheetReport.Charts[chartIndex];


            //Chart title
            leadSourceColumn.Title.Text = string.Format("Tỉ trọng từng loại tiền từng thị trường \n Tháng {0}/{1}", month, year);
            leadSourceColumn.Title.Font.Color = Color.Silver;

            listDataCompareMonth = new HSReportBL().ColumnChartStackDetailtMTCompareMonthForAllPercent(year, month, reportTypeID, marketID);
            List<ReportDetailtForTotalMoneyType> listDataCompareMonthConvert = new List<ReportDetailtForTotalMoneyType>();

            List<string> listMarketCurrent = new List<string>();

            foreach (ReportDetailtForTotalMoneyType item in listDataCompareMonth)
            {
                if (!listMarketCurrent.Contains(item.MarketName))
                {
                    listMarketCurrent.Add(item.MarketName);
                }
                item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;
            }

            if (listDataCompareMonth.Count > 0)
            {
                List<ReportDetailtForTotalMoneyType> listDataYear = listDataCompareMonth.Where(x => x.Month == month.ToString() && x.Year == year.ToString()).ToList();
                List<ReportDetailtForTotalMoneyType> listDataLastMonth = listDataCompareMonth.Where(x => x.Month == (month - 1).ToString() && x.Year == year.ToString()).ToList();
                List<ReportDetailtForTotalMoneyType> listDataLastYear = listDataCompareMonth.Where(x => x.Month == month.ToString() && x.Year == (year - 1).ToString()).ToList();
                // Trường hợp tháng 1
                if (month == 1)
                {
                    listDataLastMonth = listDataCompareMonth.Where(x => x.Month == "12" && x.Year == (year - 1).ToString()).ToList();
                }
                // Trường hợp đủ dữ liệu tháng hiện tại, tháng trước và cùng kì năm trước
                if (listDataYear.Count > 0 && listDataLastMonth.Count > 0 && listDataLastYear.Count > 0)
                {
                    // Tháng hiện tại
                    double sumVNDTotalYear = listDataYear.Sum(x => x.VND);
                    double sumUSDTotalYear = listDataYear.Sum(x => x.USD);
                    double sumEURTotalYear = listDataYear.Sum(x => x.EUR);
                    double sumCADTotalYear = listDataYear.Sum(x => x.CAD);
                    double sumAUDTotalYear = listDataYear.Sum(x => x.AUD);
                    double sumGBPTotalYear = listDataYear.Sum(x => x.GBP);

                    double sumTongDSTotalYear = listDataYear.Sum(x => x.TongDS);

                    // Tháng trước
                    double sumVNDTotalLastMonth = listDataLastMonth.Sum(x => x.VND);
                    double sumUSDTotalLastMonth = listDataLastMonth.Sum(x => x.USD);
                    double sumEURTotalLastMonth = listDataLastMonth.Sum(x => x.EUR);
                    double sumCADTotalLastMonth = listDataLastMonth.Sum(x => x.CAD);
                    double sumAUDTotalLastMonth = listDataLastMonth.Sum(x => x.AUD);
                    double sumGBPTotalLastMonth = listDataLastMonth.Sum(x => x.GBP);

                    double sumTongDSTotalLastMonth = listDataLastMonth.Sum(x => x.TongDS);

                    // Cùng kì năm trước
                    double sumVNDTotalLastYear = listDataLastYear.Sum(x => x.VND);
                    double sumUSDTotalLastYear = listDataLastYear.Sum(x => x.USD);
                    double sumEURTotalLastYear = listDataLastYear.Sum(x => x.EUR);
                    double sumCADTotalLastYear = listDataLastYear.Sum(x => x.CAD);
                    double sumAUDTotalLastYear = listDataLastYear.Sum(x => x.AUD);
                    double sumGBPTotalLastYear = listDataLastYear.Sum(x => x.GBP);

                    double sumTongDSTotalLastYear = listDataLastYear.Sum(x => x.TongDS);

                    foreach (ReportDetailtForTotalMoneyType item in listDataCompareMonth)
                    {
                        // Tháng hiện tại
                        if (item.Month == month.ToString() && item.Year == year.ToString())
                        {
                            listDataCompareMonthConvert.Add(
                                new ReportDetailtForTotalMoneyType()
                                {
                                    MarketCode = item.MarketCode,
                                    MarketName = item.MarketName,
                                    VND = sumVNDTotalYear == 0 ? 0 : Math.Round((item.VND / sumVNDTotalYear) * 100, 2, MidpointRounding.ToEven),
                                    USD = sumUSDTotalYear == 0 ? 0 : Math.Round((item.USD / sumUSDTotalYear) * 100, 2, MidpointRounding.ToEven),
                                    EUR = sumEURTotalYear == 0 ? 0 : Math.Round((item.EUR / sumEURTotalYear) * 100, 2, MidpointRounding.ToEven),
                                    CAD = sumCADTotalYear == 0 ? 0 : Math.Round((item.CAD / sumCADTotalYear) * 100, 2, MidpointRounding.ToEven),
                                    AUD = sumAUDTotalYear == 0 ? 0 : Math.Round((item.AUD / sumAUDTotalYear) * 100, 2, MidpointRounding.ToEven),
                                    GBP = sumGBPTotalYear == 0 ? 0 : Math.Round((item.GBP / sumGBPTotalYear) * 100, 2, MidpointRounding.ToEven),
                                    Month = month.ToString(),
                                    Year = year.ToString()
                                }
                            );
                        }

                        // Tháng trước
                        if (month == 1)
                        {
                            if (item.Month == "12" && item.Year == (year - 1).ToString())
                            {
                                listDataCompareMonthConvert.Add(
                                    new ReportDetailtForTotalMoneyType()
                                    {
                                        MarketCode = item.MarketCode,
                                        MarketName = item.MarketName,
                                        VND = sumVNDTotalLastMonth == 0 ? 0 : Math.Round((item.VND / sumVNDTotalLastMonth) * 100, 2, MidpointRounding.ToEven),
                                        USD = sumUSDTotalLastMonth == 0 ? 0 : Math.Round((item.USD / sumUSDTotalLastMonth) * 100, 2, MidpointRounding.ToEven),
                                        EUR = sumEURTotalLastMonth == 0 ? 0 : Math.Round((item.EUR / sumEURTotalLastMonth) * 100, 2, MidpointRounding.ToEven),
                                        CAD = sumCADTotalLastMonth == 0 ? 0 : Math.Round((item.CAD / sumCADTotalLastMonth) * 100, 2, MidpointRounding.ToEven),
                                        AUD = sumAUDTotalLastMonth == 0 ? 0 : Math.Round((item.AUD / sumAUDTotalLastMonth) * 100, 2, MidpointRounding.ToEven),
                                        GBP = sumGBPTotalLastMonth == 0 ? 0 : Math.Round((item.GBP / sumGBPTotalLastMonth) * 100, 2, MidpointRounding.ToEven),
                                        Month = "12",
                                        Year = (year - 1).ToString()
                                    }
                                );
                            }
                        }
                        else
                        {
                            if (item.Month == (month - 1).ToString() && item.Year == year.ToString())
                            {
                                listDataCompareMonthConvert.Add(
                                    new ReportDetailtForTotalMoneyType()
                                    {
                                        MarketCode = item.MarketCode,
                                        MarketName = item.MarketName,
                                        VND = sumVNDTotalLastMonth == 0 ? 0 : Math.Round((item.VND / sumVNDTotalLastMonth) * 100, 2, MidpointRounding.ToEven),
                                        USD = sumUSDTotalLastMonth == 0 ? 0 : Math.Round((item.USD / sumUSDTotalLastMonth) * 100, 2, MidpointRounding.ToEven),
                                        EUR = sumEURTotalLastMonth == 0 ? 0 : Math.Round((item.EUR / sumEURTotalLastMonth) * 100, 2, MidpointRounding.ToEven),
                                        CAD = sumCADTotalLastMonth == 0 ? 0 : Math.Round((item.CAD / sumCADTotalLastMonth) * 100, 2, MidpointRounding.ToEven),
                                        AUD = sumAUDTotalLastMonth == 0 ? 0 : Math.Round((item.AUD / sumAUDTotalLastMonth) * 100, 2, MidpointRounding.ToEven),
                                        GBP = sumGBPTotalLastMonth == 0 ? 0 : Math.Round((item.GBP / sumGBPTotalLastMonth) * 100, 2, MidpointRounding.ToEven),
                                        Month = (month - 1).ToString(),
                                        Year = year.ToString()
                                    }
                                );
                            }
                        }


                        // Cùng kì năm trước
                        if (item.Month == month.ToString() && item.Year == (year - 1).ToString())
                        {
                            listDataCompareMonthConvert.Add(
                                new ReportDetailtForTotalMoneyType()
                                {
                                    MarketCode = item.MarketCode,
                                    MarketName = item.MarketName,
                                    VND = sumVNDTotalLastYear == 0 ? 0 : Math.Round((item.VND / sumVNDTotalLastYear) * 100, 2, MidpointRounding.ToEven),
                                    USD = sumUSDTotalLastYear == 0 ? 0 : Math.Round((item.USD / sumUSDTotalLastYear) * 100, 2, MidpointRounding.ToEven),
                                    EUR = sumEURTotalLastYear == 0 ? 0 : Math.Round((item.EUR / sumEURTotalLastYear) * 100, 2, MidpointRounding.ToEven),
                                    CAD = sumCADTotalLastYear == 0 ? 0 : Math.Round((item.CAD / sumCADTotalLastYear) * 100, 2, MidpointRounding.ToEven),
                                    AUD = sumAUDTotalLastYear == 0 ? 0 : Math.Round((item.AUD / sumAUDTotalLastYear) * 100, 2, MidpointRounding.ToEven),
                                    GBP = sumGBPTotalLastYear == 0 ? 0 : Math.Round((item.GBP / sumGBPTotalLastYear) * 100, 2, MidpointRounding.ToEven),
                                    Month = month.ToString(),
                                    Year = (year - 1).ToString()
                                }
                            );
                        }
                    }

                    if (listDataCompareMonthConvert.Count > 0)
                    {
                        listDataCompareMonth = new List<ReportDetailtForTotalMoneyType>(listDataCompareMonthConvert);
                    }
                }
            }

            // List dữ liệu dataRow
            string[] listTotalRowData = new string[listMarketCurrent.Count];
            int i = 0;
            foreach (string item in listMarketCurrent)
            {
                // Năm trước
                List<ReportDetailtForTotalMoneyType> listDataYear = listDataCompareMonth.Where(x => x.MarketName == item && x.Month == month.ToString() && x.Year == year.ToString()).ToList();
                List<ReportDetailtForTotalMoneyType> listDataLastMonth = listDataCompareMonth.Where(x => x.MarketName == item && x.Month == (month - 1).ToString() && x.Year == year.ToString()).ToList();
                List<ReportDetailtForTotalMoneyType> listDataMonthLastYear = listDataCompareMonth.Where(x => x.MarketName == item && x.Month == month.ToString() && x.Year == (year - 1).ToString()).ToList();
                // Trường hợp tháng 1
                if (month == 1)
                {
                    listDataLastMonth = listDataCompareMonth.Where(x => x.MarketName == item && x.Month == "12" && x.Year == (year - 1).ToString()).ToList();
                }
                List<ReportDetailtForTotalMoneyType> listDataLastYear = listDataCompareMonth.Where(x => x.MarketName == item && x.Month == month.ToString() && x.Year == (year - 1).ToString()).ToList();

                listTotalRowData[i++] = string.Concat("{"
                    , string.Format("{0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9}, {10}, {11}, {12}, {13}, {14}, {15}, {16}, {17}"
                    , listDataYear.Sum(x => x.VND), listDataLastMonth.Sum(x => x.VND), listDataLastYear.Sum(x => x.VND)
                    , listDataYear.Sum(x => x.USD), listDataLastMonth.Sum(x => x.USD), listDataLastYear.Sum(x => x.USD)
                    , listDataYear.Sum(x => x.EUR), listDataLastMonth.Sum(x => x.EUR), listDataLastYear.Sum(x => x.EUR)
                    , listDataYear.Sum(x => x.CAD), listDataLastMonth.Sum(x => x.CAD), listDataLastYear.Sum(x => x.CAD)
                    , listDataYear.Sum(x => x.AUD), listDataLastMonth.Sum(x => x.AUD), listDataLastYear.Sum(x => x.AUD)
                    , listDataYear.Sum(x => x.GBP), listDataLastMonth.Sum(x => x.GBP), listDataLastYear.Sum(x => x.GBP)
                    )
                    , "}");
            }

            int count = 0;
            string yearCurent = string.Format("Năm {0}", year);
            string lastYear = string.Format("Năm {0}", year - 1);
            foreach (string item in listTotalRowData)
            {
                totalRowData = item;
                leadSourceColumn.NSeries.Add(totalRowData, true);

                categoryData = string.Concat("{"
                    , string.Format("VND {0}/{3}, VND {1}/{2}, VND {0}/{4}, USD {0}/{3}, USD {1}/{2}, USD {0}/{4}, EUR {0}/{3}, EUR {1}/{2}, EUR {0}/{4}, CAD {0}/{3}, CAD {1}/{2}, CAD {0}/{4}, AUD {0}/{3}, AUD {1}/{2}, AUD {0}/{4}, GBP {0}/{3}, GBP {1}/{2}, GBP {0}/{4}"
                    , month
                    , month == 1 ? 12 : month - 1
                    , month == 1 ? year - 1 : year
                    , year
                    , year - 1)
                    , "}");

                leadSourceColumn.NSeries.CategoryData = categoryData;

                leadSourceColumn.NSeries[count].Name = listMarketCurrent[count];

                count++;
            }

            // Set plot area formatting as none and hide its border.
            leadSourceColumn.PlotArea.Area.FillFormat.FillType = FillType.None;
            leadSourceColumn.PlotArea.Border.IsVisible = false;

            // Format chart percent number
            //leadSourceColumn.ValueAxis.TickLabels.NumberFormat = "0.00%";


            // Set value axis major tick mark as none and hide axis line. 
            // Also set the color of value axis major grid lines.
            leadSourceColumn.ValueAxis.MajorTickMark = TickMarkType.None;
            leadSourceColumn.ValueAxis.MaxValue = 100;
            leadSourceColumn.ValueAxis.AxisLine.IsVisible = false;
            leadSourceColumn.ValueAxis.MajorGridLines.Color = Color.FromArgb(217, 217, 217);

            // Chạy process
            designer.Process();
            return ExportReport("ReportCompareForMonthForAll", designer);
        }

        /// <summary>
        /// Tạo mẫu cho Excel cho so sánh theo tháng cho tất cả
        /// </summary>
        /// <param name="gradationID"></param>
        /// <param name="year"></param>
        /// <param name="typeID"></param>
        /// <returns></returns>
        public ActionResult CreateExcelCompareMonthForOne(int year, int month, string reportTypeID, string marketID)
        {
            WorkbookDesigner designer = new WorkbookDesigner();
            string templatePath = "~/Content/Report/ReportHSDetailt/ReportHSDetailtLHMoneyTypeCompare.xlsx";
            // Get đường dẫn
            templatePath = System.Web.HttpContext.Current.Server.MapPath(templatePath);

            designer.Workbook = new DongA.Core.DongAExcel().OpenExcelFile(templatePath);
            designer.Workbook.CalculateFormula();

            WorksheetCollection workSheets = designer.Workbook.Worksheets;
            Worksheet sheetReport = designer.Workbook.Worksheets[0];

            // Tạo title
            string typeReport = "So sánh - Theo tháng - Từng thị trường";
            // Tạo title
            CreateTitle("A2", "P2", sheetReport, typeReport, 14);

            // Tạo title detailt
            string titleDetailt = string.Format("Tháng {0}/{1}", month, year);
            CreateTitle("A3", "P3", sheetReport, titleDetailt, 12);


            // Tạo giá trị cho cột dữ liệu của Chi quầy/ Chi nhà/ Chuyển khoản
            sheetReport.Cells["B102"].PutValue(string.Format("Tháng {0}/{1} ", month, year - 1));
            sheetReport.Cells["C102"].PutValue(string.Format("Tháng {0}/{1} ", month - 1, year));
            if(month == 1)
            {
                sheetReport.Cells["C102"].PutValue(string.Format("Tháng {0}/{1} ", 12, year - 1));
            }
            sheetReport.Cells["D102"].PutValue(string.Format("Tháng {0}/{1} ", month, year));

            sheetReport.Cells["E102"].PutValue(string.Format("Tháng {0}/{1} ", month, year - 1));
            sheetReport.Cells["F102"].PutValue(string.Format("Tháng {0}/{1} ", month - 1, year));
            if (month == 1)
            {
                sheetReport.Cells["F102"].PutValue(string.Format("Tháng {0}/{1} ", 12, year - 1));
            }
            sheetReport.Cells["G102"].PutValue(string.Format("Tháng {0}/{1} ", month, year));

            sheetReport.Cells["H102"].PutValue(string.Format("Tháng {0}/{1} ", month, year - 1));
            sheetReport.Cells["I102"].PutValue(string.Format("Tháng {0}/{1} ", month - 1, year));
            if (month == 1)
            {
                sheetReport.Cells["I102"].PutValue(string.Format("Tháng {0}/{1} ", 12, year - 1));
            }
            sheetReport.Cells["J102"].PutValue(string.Format("Tháng {0}/{1} ", month, year));

            sheetReport.Cells["K102"].PutValue(string.Format("Tháng {0}/{1} ", month, year - 1));
            sheetReport.Cells["L102"].PutValue(string.Format("Tháng {0}/{1} ", month - 1, year));
            if (month == 1)
            {
                sheetReport.Cells["L102"].PutValue(string.Format("Tháng {0}/{1} ", 12, year - 1));
            }
            sheetReport.Cells["M102"].PutValue(string.Format("Tháng {0}/{1} ", month, year));

            sheetReport.Cells["N102"].PutValue(string.Format("Tháng {0}/{1} ", month, year - 1));
            sheetReport.Cells["O102"].PutValue(string.Format("Tháng {0}/{1} ", month - 1, year));
            if (month == 1)
            {
                sheetReport.Cells["O102"].PutValue(string.Format("Tháng {0}/{1} ", 12, year - 1));
            }
            sheetReport.Cells["P102"].PutValue(string.Format("Tháng {0}/{1} ", month, year));

            sheetReport.Cells["Q102"].PutValue(string.Format("Tháng {0}/{1} ", month, year - 1));
            sheetReport.Cells["R102"].PutValue(string.Format("Tháng {0}/{1} ", month - 1, year));
            if (month == 1)
            {
                sheetReport.Cells["R102"].PutValue(string.Format("Tháng {0}/{1} ", 12, year - 1));
            }
            sheetReport.Cells["S102"].PutValue(string.Format("Tháng {0}/{1} ", month, year));

            sheetReport.Cells["T102"].PutValue(string.Format("Tháng {0}/{1} ", month, year - 1));
            sheetReport.Cells["U102"].PutValue(string.Format("Tháng {0}/{1} ", month - 1, year));
            if (month == 1)
            {
                sheetReport.Cells["U102"].PutValue(string.Format("Tháng {0}/{1} ", 12, year - 1));
            }
            sheetReport.Cells["V102"].PutValue(string.Format("Tháng {0}/{1} ", month, year));

            List<ReportDetailtForTotalMoneyType> listDataCompareMonth = new HSReportBL().ReportDetailtMTCompareMonthForOne(year, month, reportTypeID, marketID);

            // Khởi tạo datatable
            DataTable table = new DataTable();
            // Tạo các cột cho datatable
            table.Columns.Add("PartnerName", typeof(String));
            table.Columns.Add("VND1", typeof(double));
            table.Columns.Add("VND2", typeof(double));
            table.Columns.Add("VND3", typeof(double));

            table.Columns.Add("USD1", typeof(double));
            table.Columns.Add("USD2", typeof(double));
            table.Columns.Add("USD3", typeof(double));

            table.Columns.Add("EUR1", typeof(double));
            table.Columns.Add("EUR2", typeof(double));
            table.Columns.Add("EUR3", typeof(double));

            table.Columns.Add("CAD1", typeof(double));
            table.Columns.Add("CAD2", typeof(double));
            table.Columns.Add("CAD3", typeof(double));

            table.Columns.Add("AUD1", typeof(double));
            table.Columns.Add("AUD2", typeof(double));
            table.Columns.Add("AUD3", typeof(double));

            table.Columns.Add("GBP1", typeof(double));
            table.Columns.Add("GBP2", typeof(double));
            table.Columns.Add("GBP3", typeof(double));

            table.Columns.Add("TDS1", typeof(double));
            table.Columns.Add("TDS2", typeof(double));
            table.Columns.Add("TDS3", typeof(double));

            table.Columns.Add("MarketName", typeof(String));

            List<string> listTemp = new List<string>();

            foreach (ReportDetailtForTotalMoneyType item in listDataCompareMonth)
            {
                item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;
            }

            string nameMarket = string.Empty;

            if (!marketID.Contains("005"))
            {
                nameMarket = listDataCompareMonth.Count > 0 ? listDataCompareMonth[0].MarketName : string.Empty;
                CreateTitle("A4", "P4", sheetReport, string.Format("Thị trường: {0}", nameMarket), 14);

                foreach (ReportDetailtForTotalMoneyType item in listDataCompareMonth)
                {
                    // Cùng kì
                    ReportDetailtForTotalMoneyType dataItemLastYear = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == month.ToString() && x.Year == (year - 1).ToString());
                    ReportDetailtForTotalMoneyType dataItemYear = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == month.ToString() && x.Year == year.ToString());
                    ReportDetailtForTotalMoneyType dataItemLastMonth = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == (month - 1).ToString() && x.Year == year.ToString());
                    // Trường hợp tháng 1
                    if (month == 1)
                    {
                        dataItemLastMonth = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == "12" && x.Year == (year - 1).ToString());
                    }

                    if (!listTemp.Contains(item.PartnerCode))
                    {
                        // Trường hợp năm trước có đối tác không có
                        if (dataItemLastYear == null)
                        {
                            dataItemLastYear = new ReportDetailtForTotalMoneyType();
                            dataItemLastYear.MarketCode = item.MarketCode;
                            dataItemLastYear.MarketName = item.MarketName;
                            dataItemLastYear.PartnerName = item.PartnerName;
                            dataItemLastYear.Year = (year - 1).ToString();
                            dataItemLastYear.Month = month.ToString();
                        }

                        // Trường hợp năm có đối tác không có
                        if (dataItemYear == null)
                        {
                            dataItemYear = new ReportDetailtForTotalMoneyType();
                            dataItemYear.MarketCode = item.MarketCode;
                            dataItemYear.MarketName = item.MarketName;
                            dataItemYear.PartnerName = item.PartnerName;
                            dataItemYear.Year = year.ToString();
                            dataItemYear.Month = month.ToString();
                        }

                        // Trường hợp năm có tháng trước có đối tác không có
                        if (dataItemLastMonth == null)
                        {
                            dataItemLastMonth = new ReportDetailtForTotalMoneyType();
                            dataItemLastMonth.MarketCode = item.MarketCode;
                            dataItemLastMonth.MarketName = item.MarketName;
                            dataItemLastMonth.PartnerName = item.PartnerName;
                            if (month == 1)
                            {
                                dataItemLastMonth.Year = (year - 1).ToString();
                                dataItemLastMonth.Month = "12";
                            }
                            else
                            {
                                dataItemLastMonth.Year = (month - 1).ToString();
                                dataItemLastMonth.Month = year.ToString();
                            }
                        }

                        // Check tồn tại của item
                        string value = string.Format("PartnerName='{0}'", item.PartnerName);
                        DataRow[] foundRows = table.Select(value);

                        if (dataItemLastYear != null && dataItemYear != null && dataItemLastMonth != null && foundRows.Count() == 0)
                        {
                            // add item vào table
                            table.Rows.Add(item.PartnerName
                                , dataItemYear.VND, dataItemLastMonth.VND, dataItemLastYear.VND
                                , dataItemYear.USD, dataItemLastMonth.USD, dataItemLastYear.USD
                                , dataItemYear.EUR, dataItemLastMonth.EUR, dataItemLastYear.EUR
                                , dataItemYear.CAD, dataItemLastMonth.CAD, dataItemLastYear.CAD
                                , dataItemYear.AUD, dataItemLastMonth.AUD, dataItemLastYear.AUD
                                , dataItemYear.GBP, dataItemLastMonth.GBP, dataItemLastYear.GBP
                                , dataItemYear.TongDS, dataItemLastMonth.TongDS, dataItemLastYear.TongDS
                                , item.MarketName
                                );
                        }
                    }
                }
            }
            else
            {
                foreach (ReportDetailtForTotalMoneyType item in listDataCompareMonth)
                {
                    if (!listTemp.Contains(item.MarketName))
                    {
                        listTemp.Add(item.MarketName);
                    }
                }

                nameMarket = listTemp.Count > 1 ? "Thị trường Châu Á" : listDataCompareMonth[0].MarketName;
                CreateTitle("A4", "P4", sheetReport, string.Format("Thị trường: {0}", nameMarket), 14);
                
                foreach (string item in listTemp)
                {
                    List<ReportDetailtForTotalMoneyType> dataItemLastYear = listDataCompareMonth.Where(x => x.MarketName == item && x.Month == month.ToString() && x.Year == (year - 1).ToString()).ToList();
                    List<ReportDetailtForTotalMoneyType> dataItemYear = listDataCompareMonth.Where(x => x.MarketName == item && x.Month == month.ToString() && x.Year == year.ToString()).ToList();
                    List<ReportDetailtForTotalMoneyType> dataItemLastMonth = listDataCompareMonth.Where(x => x.MarketName == item && x.Month == (month - 1).ToString() && x.Year == year.ToString()).ToList();
                    // Trường hợp tháng 1
                    if (month == 1)
                    {
                        dataItemLastMonth = listDataCompareMonth.Where(x => x.MarketName == item && x.Month == "12" && x.Year == (year - 1).ToString()).ToList();
                    }

                    // Cùng kì
                    if (dataItemLastYear.Count == 0)
                    {
                        dataItemLastYear = new List<ReportDetailtForTotalMoneyType>()
                        {
                            new ReportDetailtForTotalMoneyType()
                            {
                                MarketName = item,
                                Month = month.ToString(),
                                Year = (year - 1).ToString()
                            }
                        };
                    }

                    // Tháng hiện tại
                    if (dataItemYear.Count == 0)
                    {
                        dataItemYear = new List<ReportDetailtForTotalMoneyType>()
                        {
                            new ReportDetailtForTotalMoneyType()
                            {
                                MarketName = item,
                                Month = month.ToString(),
                                Year = year.ToString()
                            }
                        };
                    }

                    // Tháng trước
                    if (dataItemLastMonth.Count == 0)
                    {
                        if (month == 1)
                        {
                            dataItemLastMonth = new List<ReportDetailtForTotalMoneyType>()
                            {
                                new ReportDetailtForTotalMoneyType()
                                {
                                    MarketName = item,
                                    Month = "12",
                                    Year = (year - 1).ToString()
                                }
                            };
                        }
                        else
                        {

                            dataItemLastMonth = new List<ReportDetailtForTotalMoneyType>()
                            {
                                new ReportDetailtForTotalMoneyType()
                                {
                                    MarketName = item,
                                    Month = (month -1).ToString(),
                                    Year = year.ToString()
                                }
                            };
                        }
                    }

                    // Cùng kì năm trước
                    double sumVNDLastYear = dataItemLastYear.Sum(x => x.VND);
                    double sumUSDLastYear = dataItemLastYear.Sum(x => x.USD);
                    double sumEURLastYear = dataItemLastYear.Sum(x => x.EUR);
                    double sumCADLastYear = dataItemLastYear.Sum(x => x.CAD);
                    double sumAUDLastYear = dataItemLastYear.Sum(x => x.AUD);
                    double sumGBPLastYear = dataItemLastYear.Sum(x => x.GBP);

                    double sumTongDSLastYear = dataItemLastYear.Sum(x => x.TongDS);

                    // Tháng hiện tại
                    double sumVNDYear = dataItemYear.Sum(x => x.VND);
                    double sumUSDYear = dataItemYear.Sum(x => x.USD);
                    double sumEURYear = dataItemYear.Sum(x => x.EUR);
                    double sumCADYear = dataItemYear.Sum(x => x.CAD);
                    double sumAUDYear = dataItemYear.Sum(x => x.AUD);
                    double sumGBPYear = dataItemYear.Sum(x => x.GBP);

                    double sumTongDSYear = dataItemYear.Sum(x => x.TongDS);

                    // Tháng trước
                    double sumVNDLastMonth = dataItemLastMonth.Sum(x => x.VND);
                    double sumUSDLastMonth = dataItemLastMonth.Sum(x => x.USD);
                    double sumEURLastMonth = dataItemLastMonth.Sum(x => x.EUR);
                    double sumCADLastMonth = dataItemLastMonth.Sum(x => x.CAD);
                    double sumAUDLastMonth = dataItemLastMonth.Sum(x => x.AUD);
                    double sumGBPLastMonth = dataItemLastMonth.Sum(x => x.GBP);

                    double sumTongDSLastMonth = dataItemLastMonth.Sum(x => x.TongDS);

                    table.Rows.Add(
                        item
                        , sumVNDYear, sumVNDLastMonth, sumVNDLastYear
                        , sumUSDYear, sumUSDLastMonth, sumUSDLastYear
                        , sumEURYear, sumEURLastMonth, sumEURLastYear
                        , sumCADYear, sumCADLastMonth, sumCADLastYear
                        , sumAUDYear, sumAUDLastMonth, sumAUDLastYear
                        , sumGBPYear, sumGBPLastMonth, sumGBPLastYear
                        , sumTongDSYear, sumTongDSLastMonth, sumTongDSLastYear
                        , "Thị trường Châu Á"
                    );
                }
            }

            DataRow row = table.NewRow();
            row["MarketName"] = "Tổng";
            row["PartnerName"] = "Tổng";
            row["VND1"] = table.Compute("Sum(VND1)", "");
            row["VND2"] = table.Compute("Sum(VND2)", "");
            row["VND3"] = table.Compute("Sum(VND3)", "");

            row["USD1"] = table.Compute("Sum(USD1)", "");
            row["USD2"] = table.Compute("Sum(USD2)", "");
            row["USD3"] = table.Compute("Sum(USD3)", "");

            row["EUR1"] = table.Compute("Sum(EUR1)", "");
            row["EUR2"] = table.Compute("Sum(EUR2)", "");
            row["EUR3"] = table.Compute("Sum(EUR3)", "");

            row["CAD1"] = table.Compute("Sum(CAD1)", "");
            row["CAD2"] = table.Compute("Sum(CAD2)", "");
            row["CAD3"] = table.Compute("Sum(CAD3)", "");

            row["AUD1"] = table.Compute("Sum(AUD1)", "");
            row["AUD2"] = table.Compute("Sum(AUD2)", "");
            row["AUD3"] = table.Compute("Sum(AUD3)", "");

            row["GBP1"] = table.Compute("Sum(GBP1)", "");
            row["GBP2"] = table.Compute("Sum(GBP2)", "");
            row["GBP3"] = table.Compute("Sum(GBP3)", "");

            row["TDS1"] = table.Compute("Sum(TDS1)", "");
            row["TDS2"] = table.Compute("Sum(TDS2)", "");
            row["TDS3"] = table.Compute("Sum(TDS3)", "");

            table.Rows.Add(row);

            // Tổng số row theo table1
            int totalRowTable1 = table.Rows.Count + 102;

            // Set border
            Style style = new CellsFactory().CreateStyle();
            style.SetBorder(BorderType.BottomBorder, CellBorderType.Thin, Color.Black);
            style.SetBorder(BorderType.LeftBorder, CellBorderType.Thin, Color.Black);
            style.SetBorder(BorderType.RightBorder, CellBorderType.Thin, Color.Black);

            if (table.Rows.Count > 0)
            {
                int stepRow = 0;
                // total row = row start + số row hiện có
                int totalRow = table.Rows.Count + 102;
                // Số dòng của row
                for (int a = 102; a < totalRow; a++)
                {
                    int stepColumn = 0;
                    // Số cột trong báo cáo cần hiển thị
                    // Tổng số cột hiển thị = Số cột hiển thị bắt đầu + tổng số cột cần hiển thị
                    int totalCol = 0 + 22;
                    for (int b = 0; b < totalCol; b++)
                    {
                        // Giá trị của value trong table
                        string valueOfTable = table.Rows[stepRow][stepColumn].ToString();

                        // Insert vào dòng cột xác định trong Excel
                        sheetReport.Cells[a, b].PutValue(valueOfTable, true);

                        // set style cho number
                        style.Custom = "#,##0";

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


            // So sánh với tháng trước, cùng kì năm trước
            listDataCompareMonth = new HSReportBL().ReportDetailtMTCompareMonthForOne(year, month, reportTypeID, marketID);

            foreach (ReportDetailtForTotalMoneyType item in listDataCompareMonth)
            {
                item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;
            }

            // Khởi tạo datatable
            table = new DataTable();
            // Tạo các cột cho datatable
            // So sánh với tháng trước
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

            listTemp = new List<string>();

            if (!marketID.Contains("005"))
            {
                foreach (ReportDetailtForTotalMoneyType item in listDataCompareMonth)
                {
                    // Cùng kì
                    ReportDetailtForTotalMoneyType dataItemLastYear = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == month.ToString() && x.Year == (year - 1).ToString());
                    ReportDetailtForTotalMoneyType dataItemYear = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == month.ToString() && x.Year == year.ToString());
                    ReportDetailtForTotalMoneyType dataItemLastMonth = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == (month - 1).ToString() && x.Year == year.ToString());
                    // Trường hợp tháng 1
                    if (month == 1)
                    {
                        dataItemLastMonth = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == "12" && x.Year == (year - 1).ToString());
                    }

                    if (!listTemp.Contains(item.PartnerCode))
                    {
                        // Trường hợp năm trước không có đối tác
                        if (dataItemLastYear == null)
                        {
                            dataItemLastYear = new ReportDetailtForTotalMoneyType();
                            dataItemLastYear.MarketCode = item.MarketCode;
                            dataItemLastYear.MarketName = item.MarketName;
                            dataItemLastYear.PartnerName = item.PartnerName;
                            dataItemLastYear.Year = (year - 1).ToString();
                            dataItemLastYear.Month = month.ToString();
                        }

                        // Trường hợp năm hiện tại không có đối tác
                        if (dataItemYear == null)
                        {
                            dataItemYear = new ReportDetailtForTotalMoneyType();
                            dataItemYear.MarketCode = item.MarketCode;
                            dataItemYear.MarketName = item.MarketName;
                            dataItemYear.PartnerName = item.PartnerName;
                            dataItemYear.Year = year.ToString();
                            dataItemYear.Month = month.ToString();
                        }

                        // Trường hợp tháng trước không có
                        if (dataItemLastMonth == null)
                        {
                            dataItemLastMonth = new ReportDetailtForTotalMoneyType();
                            dataItemLastMonth.MarketCode = item.MarketCode;
                            dataItemLastMonth.MarketName = item.MarketName;
                            dataItemLastMonth.PartnerName = item.PartnerName;
                            if (month == 1)
                            {
                                dataItemLastMonth.Year = (year - 1).ToString();
                                dataItemLastMonth.Month = "12";
                            }
                            else
                            {
                                dataItemLastMonth.Year = (month - 1).ToString();
                                dataItemLastMonth.Month = year.ToString();
                            }
                        }

                        // so với tháng trước
                        double sumVND = dataItemYear.VND - dataItemLastMonth.VND;
                        double sumUSD = dataItemYear.USD - dataItemLastMonth.USD;
                        double sumEUR = dataItemYear.EUR - dataItemLastMonth.EUR;
                        double sumCAD = dataItemYear.CAD - dataItemLastMonth.CAD;
                        double sumAUD = dataItemYear.AUD - dataItemLastMonth.AUD;
                        double sumGBP = dataItemYear.GBP - dataItemLastMonth.GBP;

                        double sumTongDS = dataItemYear.TongDS - dataItemLastMonth.TongDS;

                        // so với cùng kì năm trước
                        double sumVNDLastYear = dataItemYear.VND - dataItemLastYear.VND;
                        double sumUSDLastYear = dataItemYear.USD - dataItemLastYear.USD;
                        double sumEURLastYear = dataItemYear.EUR - dataItemLastYear.EUR;
                        double sumCADLastYear = dataItemYear.CAD - dataItemLastYear.CAD;
                        double sumAUDLastYear = dataItemYear.AUD - dataItemLastYear.AUD;
                        double sumGBPLastYear = dataItemYear.GBP - dataItemLastYear.GBP;

                        double sumTongDSLastYear = dataItemYear.TongDS - dataItemLastYear.TongDS;

                        // Check tồn tại của item
                        string value = string.Format("PartnerName='{0}'", item.PartnerName);
                        DataRow[] foundRows = table.Select(value);

                        if (dataItemLastYear != null && dataItemYear != null && dataItemLastMonth != null && foundRows.Count() == 0)
                        {
                            // add item vào table
                            table.Rows.Add(item.PartnerName
                                // So với tháng hiện tại
                                , Math.Round(sumVND, 2, MidpointRounding.ToEven)
                                , Math.Round(sumVNDLastYear, 2, MidpointRounding.ToEven)

                                , Math.Round(sumUSD, 2, MidpointRounding.ToEven)
                                , Math.Round(sumUSDLastYear, 2, MidpointRounding.ToEven)

                                , Math.Round(sumEUR, 2, MidpointRounding.ToEven)
                                , Math.Round(sumEURLastYear, 2, MidpointRounding.ToEven)

                                , Math.Round(sumCAD, 2, MidpointRounding.ToEven)
                                , Math.Round(sumCADLastYear, 2, MidpointRounding.ToEven)

                                , Math.Round(sumAUD, 2, MidpointRounding.ToEven)
                                , Math.Round(sumAUDLastYear, 2, MidpointRounding.ToEven)

                                , Math.Round(sumGBP, 2, MidpointRounding.ToEven)
                                , Math.Round(sumGBPLastYear, 2, MidpointRounding.ToEven)

                                , Math.Round(sumTongDS, 2, MidpointRounding.ToEven)
                                , Math.Round(sumTongDSLastYear, 2, MidpointRounding.ToEven)
                                , item.MarketName
                                );
                        }

                        // Add partnerCode để kiểm tra
                        listTemp.Add(item.PartnerCode);
                    }
                }
            }
            else
            {
                foreach (ReportDetailtForTotalMoneyType item in listDataCompareMonth)
                {
                    if (!listTemp.Contains(item.MarketName))
                    {
                        listTemp.Add(item.MarketName);
                    }
                }

                foreach (string item in listTemp)
                {
                    List<ReportDetailtForTotalMoneyType> dataItemLastYear = listDataCompareMonth.Where(x => x.MarketName == item && x.Month == month.ToString() && x.Year == (year - 1).ToString()).ToList();
                    List<ReportDetailtForTotalMoneyType> dataItemYear = listDataCompareMonth.Where(x => x.MarketName == item && x.Month == month.ToString() && x.Year == year.ToString()).ToList();
                    List<ReportDetailtForTotalMoneyType> dataItemLastMonth = listDataCompareMonth.Where(x => x.MarketName == item && x.Month == (month - 1).ToString() && x.Year == year.ToString()).ToList();
                    // Trường hợp tháng 1
                    if (month == 1)
                    {
                        dataItemLastMonth = listDataCompareMonth.Where(x => x.MarketName == item && x.Month == "12" && x.Year == (year - 1).ToString()).ToList();
                    }
                    // Cùng kì
                    if (dataItemLastYear.Count == 0)
                    {
                        dataItemLastYear = new List<ReportDetailtForTotalMoneyType>()
                        {
                            new ReportDetailtForTotalMoneyType()
                            {
                                MarketName = item,
                                Month = month.ToString(),
                                Year = (year - 1).ToString()
                            }
                        };
                    }

                    // Tháng hiện tại
                    if (dataItemYear.Count == 0)
                    {
                        dataItemYear = new List<ReportDetailtForTotalMoneyType>()
                        {
                            new ReportDetailtForTotalMoneyType()
                            {
                                MarketName = item,
                                Month = month.ToString(),
                                Year = year.ToString()
                            }
                        };
                    }

                    // Tháng trước
                    if (dataItemLastMonth.Count == 0)
                    {
                        // Trường hợp tháng 1
                        if (month == 1)
                        {
                            dataItemLastMonth = new List<ReportDetailtForTotalMoneyType>()
                            {
                                new ReportDetailtForTotalMoneyType()
                                {
                                    MarketName = item,
                                    Month = "12",
                                    Year = (year - 1).ToString()
                                }
                            };
                        }
                        else
                        {
                            dataItemLastMonth = new List<ReportDetailtForTotalMoneyType>()
                            {
                                new ReportDetailtForTotalMoneyType()
                                {
                                    MarketName = item,
                                    Month = (month -1).ToString(),
                                    Year = year.ToString()
                                }
                            };
                        }

                    }

                    // Cùng kì năm trước
                    double sumVNDLastYear = dataItemLastYear.Sum(x => x.VND);
                    double sumUSDLastYear = dataItemLastYear.Sum(x => x.USD);
                    double sumEURLastYear = dataItemLastYear.Sum(x => x.EUR);
                    double sumCADLastYear = dataItemLastYear.Sum(x => x.CAD);
                    double sumAUDLastYear = dataItemLastYear.Sum(x => x.AUD);
                    double sumGBPLastYear = dataItemLastYear.Sum(x => x.GBP);

                    double sumTongDSLastYear = dataItemLastYear.Sum(x => x.TongDS);

                    // Tháng hiện tại
                    double sumVNDYear = dataItemYear.Sum(x => x.VND);
                    double sumUSDYear = dataItemYear.Sum(x => x.USD);
                    double sumEURYear = dataItemYear.Sum(x => x.EUR);
                    double sumCADYear = dataItemYear.Sum(x => x.CAD);
                    double sumAUDYear = dataItemYear.Sum(x => x.AUD);
                    double sumGBPYear = dataItemYear.Sum(x => x.GBP);

                    double sumTongDSYear = dataItemYear.Sum(x => x.TongDS);

                    // Tháng trước
                    double sumVNDLastMonth = dataItemLastMonth.Sum(x => x.VND);
                    double sumUSDLastMonth = dataItemLastMonth.Sum(x => x.USD);
                    double sumEURLastMonth = dataItemLastMonth.Sum(x => x.EUR);
                    double sumCADLastMonth = dataItemLastMonth.Sum(x => x.CAD);
                    double sumAUDLastMonth = dataItemLastMonth.Sum(x => x.AUD);
                    double sumGBPLastMonth = dataItemLastMonth.Sum(x => x.GBP);

                    double sumTongDSLastMonth = dataItemLastMonth.Sum(x => x.TongDS);

                    table.Rows.Add(
                        item
                        , Math.Round(sumVNDYear - sumVNDLastMonth, 2, MidpointRounding.ToEven), Math.Round(sumVNDYear - sumVNDLastYear, 2, MidpointRounding.ToEven)
                        , Math.Round(sumUSDYear - sumUSDLastMonth, 2, MidpointRounding.ToEven), Math.Round(sumUSDYear - sumUSDLastYear, 2, MidpointRounding.ToEven)
                        , Math.Round(sumEURYear - sumEURLastMonth, 2, MidpointRounding.ToEven), Math.Round(sumEURYear - sumEURLastYear, 2, MidpointRounding.ToEven)
                        , Math.Round(sumCADYear - sumCADLastMonth, 2, MidpointRounding.ToEven), Math.Round(sumCADYear - sumCADLastYear, 2, MidpointRounding.ToEven)
                        , Math.Round(sumAUDYear - sumAUDLastMonth, 2, MidpointRounding.ToEven), Math.Round(sumAUDYear - sumAUDLastYear, 2, MidpointRounding.ToEven)
                        , Math.Round(sumGBPYear - sumGBPLastMonth, 2, MidpointRounding.ToEven), Math.Round(sumGBPYear - sumGBPLastYear, 2, MidpointRounding.ToEven)
                        , Math.Round(sumTongDSYear - sumTongDSLastMonth, 2, MidpointRounding.ToEven), Math.Round(sumTongDSYear - sumTongDSLastYear, 2, MidpointRounding.ToEven)
                        , "Thị trường Châu Á"
                    );
                }

            }

            row = table.NewRow();
            row["MarketName"] = "Tổng";
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

            table.Rows.Add(row);


            // Tổng số row của table2
            // Với 6 là số cách của table1 và table2
            int totalRowTable2 = totalRowTable1 + table.Rows.Count + 6;

            // Tạo title hearder cho table tăng giảm
            // Title cho thị trường
            string title = "Đối tác";

            CreateTitle(string.Format("A{0}", totalRowTable1 + 6 - 1), string.Format("A{0}", totalRowTable1 + 6), sheetReport, title, 12, true);

            title = "VND";
            CreateTitle(string.Format("B{0}", totalRowTable1 + 6 - 1), string.Format("C{0}", totalRowTable1 + 6 - 1), sheetReport, title, 12, true);

            title = "USD";
            CreateTitle(string.Format("D{0}", totalRowTable1 + 6 - 1), string.Format("E{0}", totalRowTable1 + 6 - 1), sheetReport, title, 12, true);

            title = "EUR";
            CreateTitle(string.Format("F{0}", totalRowTable1 + 6 - 1), string.Format("G{0}", totalRowTable1 + 6 - 1), sheetReport, title, 12, true);

            title = "CAD";
            CreateTitle(string.Format("H{0}", totalRowTable1 + 6 - 1), string.Format("I{0}", totalRowTable1 + 6 - 1), sheetReport, title, 12, true);

            title = "AUD";
            CreateTitle(string.Format("J{0}", totalRowTable1 + 6 - 1), string.Format("K{0}", totalRowTable1 + 6 - 1), sheetReport, title, 12, true);

            title = "GBP";
            CreateTitle(string.Format("L{0}", totalRowTable1 + 6 - 1), string.Format("M{0}", totalRowTable1 + 6 - 1), sheetReport, title, 12, true);

            title = "Tổng";
            CreateTitle(string.Format("N{0}", totalRowTable1 + 6 - 1), string.Format("O{0}", totalRowTable1 + 6 - 1), sheetReport, title, 12, true);


            title = "So với tháng trước";
            CreateTitle(string.Format("B{0}", totalRowTable1 + 6), string.Format("B{0}", totalRowTable1 + 6), sheetReport, title, 12, true);

            title = string.Format("So với cùng kì");
            CreateTitle(string.Format("C{0}", totalRowTable1 + 6), string.Format("C{0}", totalRowTable1 + 6), sheetReport, title, 12, true);

            title = "So với tháng trước";
            CreateTitle(string.Format("D{0}", totalRowTable1 + 6), string.Format("D{0}", totalRowTable1 + 6), sheetReport, title, 12, true);

            title = string.Format("So với cùng kì");
            CreateTitle(string.Format("E{0}", totalRowTable1 + 6), string.Format("E{0}", totalRowTable1 + 6), sheetReport, title, 12, true);

            title = "So với tháng trước";
            CreateTitle(string.Format("F{0}", totalRowTable1 + 6), string.Format("F{0}", totalRowTable1 + 6), sheetReport, title, 12, true);

            title = string.Format("So với cùng kì");
            CreateTitle(string.Format("G{0}", totalRowTable1 + 6), string.Format("G{0}", totalRowTable1 + 6), sheetReport, title, 12, true);

            title = "So với tháng trước";
            CreateTitle(string.Format("H{0}", totalRowTable1 + 6), string.Format("H{0}", totalRowTable1 + 6), sheetReport, title, 12, true);

            title = string.Format("So với cùng kì");
            CreateTitle(string.Format("I{0}", totalRowTable1 + 6), string.Format("I{0}", totalRowTable1 + 6), sheetReport, title, 12, true);

            title = "So với tháng trước";
            CreateTitle(string.Format("J{0}", totalRowTable1 + 6), string.Format("J{0}", totalRowTable1 + 6), sheetReport, title, 12, true);

            title = string.Format("So với cùng kì");
            CreateTitle(string.Format("K{0}", totalRowTable1 + 6), string.Format("K{0}", totalRowTable1 + 6), sheetReport, title, 12, true);

            title = "So với tháng trước";
            CreateTitle(string.Format("L{0}", totalRowTable1 + 6), string.Format("L{0}", totalRowTable1 + 6), sheetReport, title, 12, true);

            title = string.Format("So với cùng kì");
            CreateTitle(string.Format("M{0}", totalRowTable1 + 6), string.Format("M{0}", totalRowTable1 + 6), sheetReport, title, 12, true);

            title = "So với tháng trước";
            CreateTitle(string.Format("N{0}", totalRowTable1 + 6), string.Format("N{0}", totalRowTable1 + 6), sheetReport, title, 12, true);

            title = string.Format("So với cùng kì");
            CreateTitle(string.Format("O{0}", totalRowTable1 + 6), string.Format("O{0}", totalRowTable1 + 6), sheetReport, title, 12, true);

            // Table dữ liệu bảng số liệu Hồ sơ Chi Quầy/Chi Nhà/Chuyển Khoản
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
                    int totalCol = 0 + 15;
                    for (int b = 0; b < totalCol; b++)
                    {
                        // Giá trị của value trong table
                        string valueOfTable = table.Rows[stepRow][stepColumn].ToString();

                        // Insert vào dòng cột xác định trong Excel
                        sheetReport.Cells[a, b].PutValue(valueOfTable, true);

                        // Tô màu cho các dòng có giá trị tăng giảm
                        if (b >= 1)
                        {
                            decimal tryParseValue = 0;
                            decimal.TryParse(valueOfTable, out tryParseValue);
                            style.Font.Color = Color.Green;

                            if (tryParseValue < 0)
                            {
                                style.Font.Color = Color.Red;
                            }
                        }

                        // set style cho number
                        style.Custom = "#,##0";

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
            int chartIndex = sheetReport.Charts.Add(ChartType.Bar, 7, 0, 27, 6);
            leadSourceColumnVND = sheetReport.Charts[chartIndex];

            //Chart title
            leadSourceColumnVND.Title.Text = string.Format("Hồ sơ từng loại tiền từng thị trường - VND \n Tháng {0}/{1}", month, year);
            leadSourceColumnVND.Title.Font.Color = Color.Silver;

            // Adding SeriesCollection (chart data source) to the chart ranging from "A1" cell to "B3"
            string totalRowData = string.Format("B103:D{0}", totalRowTable1 - 1);
            leadSourceColumnVND.NSeries.Add(totalRowData, true);

            // Set the category data covering the range A2:A5.
            // Tổng số dòng cần hiển thị là số dòng hiện tại  + số dòng trong listData -2 (trong đó 1 là dòng tổng cuối cùng)
            string categoryData = string.Format("A103:A{0}", totalRowTable1 - 1);
            leadSourceColumnVND.NSeries.CategoryData = categoryData;

            // Set the names of the chart series taken from cells.
            leadSourceColumnVND.NSeries[0].Name = "=B102";
            leadSourceColumnVND.NSeries[1].Name = "=C102";
            leadSourceColumnVND.NSeries[2].Name = "=D102";

            // Set the 1st series fill color.
            leadSourceColumnVND.NSeries[0].Area.ForegroundColor = Color.Orange;
            leadSourceColumnVND.NSeries[0].Area.Formatting = FormattingType.Custom;

            // Set the 2nd series fill color.
            leadSourceColumnVND.NSeries[1].Area.ForegroundColor = Color.Green;
            leadSourceColumnVND.NSeries[1].Area.Formatting = FormattingType.Custom;

            // Set the 2nd series fill color.
            leadSourceColumnVND.NSeries[2].Area.ForegroundColor = Color.Blue;
            leadSourceColumnVND.NSeries[2].Area.Formatting = FormattingType.Custom;


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

            // Vẽ biểu đồ cột cho Excel
            //Chart reference
            Aspose.Cells.Charts.Chart leadSourceColumnUSD;
            //Add Pie Chart
            // USD
            chartIndex = sheetReport.Charts.Add(ChartType.Bar, 7, 7, 27, 13);
            leadSourceColumnUSD = sheetReport.Charts[chartIndex];

            //Chart title
            leadSourceColumnUSD.Title.Text = string.Format("Hồ sơ từng loại tiền từng thị trường - USD \n Tháng {0}/{1}", month, year);
            leadSourceColumnUSD.Title.Font.Color = Color.Silver;

            // Adding SeriesCollection (chart data source) to the chart ranging from "A1" cell to "B3"
            totalRowData = string.Format("E103:G{0}", totalRowTable1 - 1);
            leadSourceColumnUSD.NSeries.Add(totalRowData, true);

            // Set the category data covering the range A2:A5.
            // Tổng số dòng cần hiển thị là số dòng hiện tại  + số dòng trong listData -2 (trong đó 1 là dòng tổng cuối cùng)
            categoryData = string.Format("A103:A{0}", totalRowTable1 - 1);
            leadSourceColumnUSD.NSeries.CategoryData = categoryData;

            // Set the names of the chart series taken from cells.
            leadSourceColumnUSD.NSeries[0].Name = "=B102";
            leadSourceColumnUSD.NSeries[1].Name = "=C102";
            leadSourceColumnUSD.NSeries[2].Name = "=D102";

            // Set the 1st series fill color.
            leadSourceColumnUSD.NSeries[0].Area.ForegroundColor = Color.Orange;
            leadSourceColumnUSD.NSeries[0].Area.Formatting = FormattingType.Custom;

            // Set the 2nd series fill color.
            leadSourceColumnUSD.NSeries[1].Area.ForegroundColor = Color.Green;
            leadSourceColumnUSD.NSeries[1].Area.Formatting = FormattingType.Custom;

            // Set the 2nd series fill color.
            leadSourceColumnUSD.NSeries[2].Area.ForegroundColor = Color.Blue;
            leadSourceColumnUSD.NSeries[2].Area.Formatting = FormattingType.Custom;


            // Set plot area formatting as none and hide its border.
            leadSourceColumnUSD.PlotArea.Area.FillFormat.FillType = FillType.None;
            leadSourceColumnUSD.PlotArea.Border.IsVisible = false;

            // Set value axis major tick mark as none and hide axis line. 
            // Also set the color of value axis major grid lines.
            leadSourceColumnUSD.ValueAxis.MajorTickMark = TickMarkType.None;
            leadSourceColumnUSD.ValueAxis.AxisLine.IsVisible = false;

            leadSourceColumnUSD.ValueAxis.MajorGridLines.Color = Color.FromArgb(217, 217, 217);

            // Vẽ biểu đồ cột cho Excel
            //Chart reference
            Aspose.Cells.Charts.Chart leadSourceColumnEUR;
            //Add Pie Chart
            // EUR
            chartIndex = sheetReport.Charts.Add(ChartType.Bar, 7, 14, 27, 20);
            leadSourceColumnEUR = sheetReport.Charts[chartIndex];

            //Chart title
            leadSourceColumnEUR.Title.Text = string.Format("Hồ sơ từng loại tiền từng thị trường - EUR \n Tháng {0}/{1}", month, year);
            leadSourceColumnEUR.Title.Font.Color = Color.Silver;

            // Adding SeriesCollection (chart data source) to the chart ranging from "A1" cell to "B3"
            totalRowData = string.Format("H103:J{0}", totalRowTable1 - 1);
            leadSourceColumnEUR.NSeries.Add(totalRowData, true);

            // Set the category data covering the range A2:A5.
            // Tổng số dòng cần hiển thị là số dòng hiện tại  + số dòng trong listData -2 (trong đó 1 là dòng tổng cuối cùng)
            categoryData = string.Format("A103:A{0}", totalRowTable1 - 1);
            leadSourceColumnEUR.NSeries.CategoryData = categoryData;

            // Set the names of the chart series taken from cells.
            leadSourceColumnEUR.NSeries[0].Name = "=B102";
            leadSourceColumnEUR.NSeries[1].Name = "=C102";
            leadSourceColumnEUR.NSeries[2].Name = "=D102";

            // Set the 1st series fill color.
            leadSourceColumnEUR.NSeries[0].Area.ForegroundColor = Color.Orange;
            leadSourceColumnEUR.NSeries[0].Area.Formatting = FormattingType.Custom;

            // Set the 2nd series fill color.
            leadSourceColumnEUR.NSeries[1].Area.ForegroundColor = Color.Green;
            leadSourceColumnEUR.NSeries[1].Area.Formatting = FormattingType.Custom;

            // Set the 2nd series fill color.
            leadSourceColumnEUR.NSeries[2].Area.ForegroundColor = Color.Blue;
            leadSourceColumnEUR.NSeries[2].Area.Formatting = FormattingType.Custom;


            // Set plot area formatting as none and hide its border.
            leadSourceColumnEUR.PlotArea.Area.FillFormat.FillType = FillType.None;
            leadSourceColumnEUR.PlotArea.Border.IsVisible = false;

            // Set value axis major tick mark as none and hide axis line. 
            // Also set the color of value axis major grid lines.
            leadSourceColumnEUR.ValueAxis.MajorTickMark = TickMarkType.None;
            leadSourceColumnEUR.ValueAxis.AxisLine.IsVisible = false;

            leadSourceColumnEUR.ValueAxis.MajorGridLines.Color = Color.FromArgb(217, 217, 217);

            // Vẽ biểu đồ cột cho Excel
            //Chart reference
            Aspose.Cells.Charts.Chart leadSourceColumnCAD;
            //Add Pie Chart
            // EUR
            chartIndex = sheetReport.Charts.Add(ChartType.Bar, 29, 0, 49, 6);
            leadSourceColumnCAD = sheetReport.Charts[chartIndex];

            //Chart title
            leadSourceColumnCAD.Title.Text = string.Format("Hồ sơ từng loại tiền từng thị trường - CAD \n Tháng {0}/{1}", month, year);
            leadSourceColumnCAD.Title.Font.Color = Color.Silver;

            // Adding SeriesCollection (chart data source) to the chart ranging from "A1" cell to "B3"
            totalRowData = string.Format("K103:M{0}", totalRowTable1 - 1);
            leadSourceColumnCAD.NSeries.Add(totalRowData, true);

            // Set the category data covering the range A2:A5.
            // Tổng số dòng cần hiển thị là số dòng hiện tại  + số dòng trong listData -2 (trong đó 1 là dòng tổng cuối cùng)
            categoryData = string.Format("A103:A{0}", totalRowTable1 - 1);
            leadSourceColumnCAD.NSeries.CategoryData = categoryData;

            // Set the names of the chart series taken from cells.
            leadSourceColumnCAD.NSeries[0].Name = "=B102";
            leadSourceColumnCAD.NSeries[1].Name = "=C102";
            leadSourceColumnCAD.NSeries[2].Name = "=D102";

            // Set the 1st series fill color.
            leadSourceColumnCAD.NSeries[0].Area.ForegroundColor = Color.Orange;
            leadSourceColumnCAD.NSeries[0].Area.Formatting = FormattingType.Custom;

            // Set the 2nd series fill color.
            leadSourceColumnCAD.NSeries[1].Area.ForegroundColor = Color.Green;
            leadSourceColumnCAD.NSeries[1].Area.Formatting = FormattingType.Custom;

            // Set the 2nd series fill color.
            leadSourceColumnCAD.NSeries[2].Area.ForegroundColor = Color.Blue;
            leadSourceColumnCAD.NSeries[2].Area.Formatting = FormattingType.Custom;


            // Set plot area formatting as none and hide its border.
            leadSourceColumnCAD.PlotArea.Area.FillFormat.FillType = FillType.None;
            leadSourceColumnCAD.PlotArea.Border.IsVisible = false;

            // Set value axis major tick mark as none and hide axis line. 
            // Also set the color of value axis major grid lines.
            leadSourceColumnCAD.ValueAxis.MajorTickMark = TickMarkType.None;
            leadSourceColumnCAD.ValueAxis.AxisLine.IsVisible = false;

            leadSourceColumnCAD.ValueAxis.MajorGridLines.Color = Color.FromArgb(217, 217, 217);

            // Vẽ biểu đồ cột cho Excel
            //Chart reference
            Aspose.Cells.Charts.Chart leadSourceColumnAUD;
            //Add Pie Chart
            // AUD
            chartIndex = sheetReport.Charts.Add(ChartType.Bar, 29, 7, 49, 13);
            leadSourceColumnAUD = sheetReport.Charts[chartIndex];

            //Chart title
            leadSourceColumnAUD.Title.Text = string.Format("Hồ sơ từng loại tiền từng thị trường - AUD \n Tháng {0}/{1}", month, year);
            leadSourceColumnAUD.Title.Font.Color = Color.Silver;

            // Adding SeriesCollection (chart data source) to the chart ranging from "A1" cell to "B3"
            totalRowData = string.Format("N103:P{0}", totalRowTable1 - 1);
            leadSourceColumnAUD.NSeries.Add(totalRowData, true);

            // Set the category data covering the range A2:A5.
            // Tổng số dòng cần hiển thị là số dòng hiện tại  + số dòng trong listData -2 (trong đó 1 là dòng tổng cuối cùng)
            categoryData = string.Format("A103:A{0}", totalRowTable1 - 1);
            leadSourceColumnAUD.NSeries.CategoryData = categoryData;

            // Set the names of the chart series taken from cells.
            leadSourceColumnAUD.NSeries[0].Name = "=B102";
            leadSourceColumnAUD.NSeries[1].Name = "=C102";
            leadSourceColumnAUD.NSeries[2].Name = "=D102";

            // Set the 1st series fill color.
            leadSourceColumnAUD.NSeries[0].Area.ForegroundColor = Color.Orange;
            leadSourceColumnAUD.NSeries[0].Area.Formatting = FormattingType.Custom;

            // Set the 2nd series fill color.
            leadSourceColumnAUD.NSeries[1].Area.ForegroundColor = Color.Green;
            leadSourceColumnAUD.NSeries[1].Area.Formatting = FormattingType.Custom;

            // Set the 2nd series fill color.
            leadSourceColumnAUD.NSeries[2].Area.ForegroundColor = Color.Blue;
            leadSourceColumnAUD.NSeries[2].Area.Formatting = FormattingType.Custom;


            // Set plot area formatting as none and hide its border.
            leadSourceColumnAUD.PlotArea.Area.FillFormat.FillType = FillType.None;
            leadSourceColumnAUD.PlotArea.Border.IsVisible = false;

            // Set value axis major tick mark as none and hide axis line. 
            // Also set the color of value axis major grid lines.
            leadSourceColumnAUD.ValueAxis.MajorTickMark = TickMarkType.None;
            leadSourceColumnAUD.ValueAxis.AxisLine.IsVisible = false;

            leadSourceColumnAUD.ValueAxis.MajorGridLines.Color = Color.FromArgb(217, 217, 217);

            // Vẽ biểu đồ cột cho Excel
            //Chart reference
            Aspose.Cells.Charts.Chart leadSourceColumnGBP;
            //Add Pie Chart
            // GBP
            chartIndex = sheetReport.Charts.Add(ChartType.Bar, 29, 14, 49, 20);
            leadSourceColumnGBP = sheetReport.Charts[chartIndex];

            //Chart title
            leadSourceColumnGBP.Title.Text = string.Format("Hồ sơ từng loại tiền từng thị trường - GBP \n Tháng {0}/{1}", month, year);
            leadSourceColumnGBP.Title.Font.Color = Color.Silver;

            // Adding SeriesCollection (chart data source) to the chart ranging from "A1" cell to "B3"
            totalRowData = string.Format("Q103:S{0}", totalRowTable1 - 1);
            leadSourceColumnGBP.NSeries.Add(totalRowData, true);

            // Set the category data covering the range A2:A5.
            // Tổng số dòng cần hiển thị là số dòng hiện tại  + số dòng trong listData -2 (trong đó 1 là dòng tổng cuối cùng)
            categoryData = string.Format("A103:A{0}", totalRowTable1 - 1);
            leadSourceColumnGBP.NSeries.CategoryData = categoryData;

            // Set the names of the chart series taken from cells.
            leadSourceColumnGBP.NSeries[0].Name = "=B102";
            leadSourceColumnGBP.NSeries[1].Name = "=C102";
            leadSourceColumnGBP.NSeries[2].Name = "=D102";

            // Set the 1st series fill color.
            leadSourceColumnGBP.NSeries[0].Area.ForegroundColor = Color.Orange;
            leadSourceColumnGBP.NSeries[0].Area.Formatting = FormattingType.Custom;

            // Set the 2nd series fill color.
            leadSourceColumnGBP.NSeries[1].Area.ForegroundColor = Color.Green;
            leadSourceColumnGBP.NSeries[1].Area.Formatting = FormattingType.Custom;

            // Set the 2nd series fill color.
            leadSourceColumnGBP.NSeries[2].Area.ForegroundColor = Color.Blue;
            leadSourceColumnGBP.NSeries[2].Area.Formatting = FormattingType.Custom;


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
            chartIndex = sheetReport.Charts.Add(ChartType.Bar3DStacked, 51, 0, 98, 13);
            leadSourceColumn = sheetReport.Charts[chartIndex];


            //Chart title
            leadSourceColumn.Title.Text = string.Format("Tỉ trọng từng loại tiền từng thị trường \n Tháng {0}/{1}", month, year);
            leadSourceColumn.Title.Font.Color = Color.Silver;

            listDataCompareMonth = new HSReportBL().ReportDetailtMTCompareMonthForOneConvertPercent(year, month, reportTypeID, marketID);
            List<ReportDetailtForTotalMoneyType> listDataCompareMonthConvert = new List<ReportDetailtForTotalMoneyType>();
            List<ReportDetailtForTotalMoneyType> listDataCommpareMonthClone = new List<ReportDetailtForTotalMoneyType>(listDataCompareMonth);

            // Trường hợp là thị trường chi tiết của loại tiền theo tháng của Châu Á
            if (marketID.Contains("005"))
            {
                // Danh sách các thị trường con của Châu Á
                List<string> listMarket = new List<string>();
                foreach (ReportDetailtForTotalMoneyType item in listDataCompareMonth)
                {
                    if (!listMarket.Contains(item.MarketName))
                    {
                        listMarket.Add(item.MarketName);
                    }
                }

                foreach (string item in listMarket)
                {
                    List<ReportDetailtForTotalMoneyType> dataItemLastYear = listDataCompareMonth.Where(x => x.MarketName == item && x.Month == month.ToString() && x.Year == (year - 1).ToString()).ToList();
                    List<ReportDetailtForTotalMoneyType> dataItemYear = listDataCompareMonth.Where(x => x.MarketName == item && x.Month == month.ToString() && x.Year == year.ToString()).ToList();
                    List<ReportDetailtForTotalMoneyType> dataItemLastMonth = listDataCompareMonth.Where(x => x.MarketName == item && x.Month == (month - 1).ToString() && x.Year == year.ToString()).ToList();
                    // Trường hợp tháng 1
                    if (month == 1)
                    {
                        dataItemLastMonth = listDataCompareMonth.Where(x => x.MarketName == item && x.Month == "12" && x.Year == (year - 1).ToString()).ToList();
                    }
                    // Cùng kì
                    if (dataItemLastYear.Count == 0)
                    {
                        dataItemLastYear = new List<ReportDetailtForTotalMoneyType>()
                        {
                            new ReportDetailtForTotalMoneyType()
                            {
                                MarketName = item,
                                Month = month.ToString(),
                                Year = (year - 1).ToString()
                            }
                        };
                    }

                    // Tháng hiện tại
                    if (dataItemYear.Count == 0)
                    {
                        dataItemYear = new List<ReportDetailtForTotalMoneyType>()
                        {
                            new ReportDetailtForTotalMoneyType()
                            {
                                MarketName = item,
                                Month = month.ToString(),
                                Year = year.ToString()
                            }
                        };
                    }

                    // Tháng trước
                    if (dataItemLastMonth.Count == 0)
                    {// Trường hợp tháng 1
                        if (month == 1)
                        {
                            dataItemLastMonth = new List<ReportDetailtForTotalMoneyType>()
                            {
                                new ReportDetailtForTotalMoneyType()
                                {
                                    MarketName = item,
                                    Month = "12",
                                    Year = (year - 1).ToString()
                                }
                            };
                        }
                        else
                        {
                            dataItemLastMonth = new List<ReportDetailtForTotalMoneyType>()
                            {
                                new ReportDetailtForTotalMoneyType()
                                {
                                    MarketName = item,
                                    Month = (month -1).ToString(),
                                    Year = year.ToString()
                                }
                            };
                        }

                    }

                    // Cùng kì năm trước
                    double sumVNDLastYear = dataItemLastYear.Sum(x => x.VND);
                    double sumUSDLastYear = dataItemLastYear.Sum(x => x.USD);
                    double sumEURLastYear = dataItemLastYear.Sum(x => x.EUR);
                    double sumCADLastYear = dataItemLastYear.Sum(x => x.CAD);
                    double sumAUDLastYear = dataItemLastYear.Sum(x => x.AUD);
                    double sumGBPLastYear = dataItemLastYear.Sum(x => x.GBP);

                    double sumTongDSLastYear = dataItemLastYear.Sum(x => x.TongDS);

                    // Tháng hiện tại
                    double sumVNDYear = dataItemYear.Sum(x => x.VND);
                    double sumUSDYear = dataItemYear.Sum(x => x.USD);
                    double sumEURYear = dataItemYear.Sum(x => x.EUR);
                    double sumCADYear = dataItemYear.Sum(x => x.CAD);
                    double sumAUDYear = dataItemYear.Sum(x => x.AUD);
                    double sumGBPYear = dataItemYear.Sum(x => x.GBP);

                    double sumTongDSYear = dataItemYear.Sum(x => x.TongDS);

                    // Tháng trước
                    double sumVNDLastMonth = dataItemLastMonth.Sum(x => x.VND);
                    double sumUSDLastMonth = dataItemLastMonth.Sum(x => x.USD);
                    double sumEURLastMonth = dataItemLastMonth.Sum(x => x.EUR);
                    double sumCADLastMonth = dataItemLastMonth.Sum(x => x.CAD);
                    double sumAUDLastMonth = dataItemLastMonth.Sum(x => x.AUD);
                    double sumGBPLastMonth = dataItemLastMonth.Sum(x => x.GBP);

                    double sumTongDSLastMonth = dataItemLastMonth.Sum(x => x.TongDS);

                    // Vì  get  dữ liệu theo partnerName nên thị trường gán cho partner
                    // Tháng hiện tại
                    listDataCompareMonthConvert.Add(
                        new ReportDetailtForTotalMoneyType()
                        {
                            PartnerName = item,
                            Month = month.ToString(),
                            Year = year.ToString(),
                            VND = sumVNDYear,
                            USD = sumUSDYear,
                            EUR = sumEURYear,
                            CAD = sumCADYear,
                            AUD = sumAUDYear,
                            GBP = sumGBPYear,
                        }
                    );

                    // Tháng Trước
                    // Trường hợp tháng 1
                    if (month == 1)
                    {
                        listDataCompareMonthConvert.Add(
                            new ReportDetailtForTotalMoneyType()
                            {
                                PartnerName = item,
                                Month = "12",
                                Year = (year - 1).ToString(),
                                VND = sumVNDLastMonth,
                                USD = sumUSDLastMonth,
                                EUR = sumEURLastMonth,
                                CAD = sumCADLastMonth,
                                AUD = sumAUDLastMonth,
                                GBP = sumGBPLastMonth,
                            }
                        );
                    }
                    else
                    {
                        listDataCompareMonthConvert.Add(
                            new ReportDetailtForTotalMoneyType()
                            {
                                PartnerName = item,
                                Month = (month - 1).ToString(),
                                Year = year.ToString(),
                                VND = sumVNDLastMonth,
                                USD = sumUSDLastMonth,
                                EUR = sumEURLastMonth,
                                CAD = sumCADLastMonth,
                                AUD = sumAUDLastMonth,
                                GBP = sumGBPLastMonth,
                            }
                        );
                    }


                    // Cùng kì năm trước
                    listDataCompareMonthConvert.Add(
                        new ReportDetailtForTotalMoneyType()
                        {
                            PartnerName = item,
                            Month = month.ToString(),
                            Year = (year - 1).ToString(),
                            VND = sumVNDLastYear,
                            USD = sumUSDLastYear,
                            EUR = sumEURLastYear,
                            CAD = sumCADLastYear,
                            AUD = sumAUDLastYear,
                            GBP = sumGBPLastYear,
                        }
                    );
                }

                if (listDataCompareMonthConvert.Count > 0)
                {
                    listDataCompareMonth = new List<ReportDetailtForTotalMoneyType>(listDataCompareMonthConvert);
                }
            }
            else
            {
                foreach (ReportDetailtForTotalMoneyType item in listDataCommpareMonthClone)
                {
                    ReportDetailtForTotalMoneyType dataItemLastYear = listDataCompareMonth.Find(x => x.PartnerName == item.PartnerName && x.Month == month.ToString() && x.Year == (year - 1).ToString());
                    ReportDetailtForTotalMoneyType dataItemYear = listDataCompareMonth.Find(x => x.PartnerName == item.PartnerName && x.Month == month.ToString() && x.Year == year.ToString());
                    ReportDetailtForTotalMoneyType dataItemLastMonth = listDataCompareMonth.Find(x => x.PartnerName == item.PartnerName && x.Month == (month - 1).ToString() && x.Year == year.ToString());

                    // Tháng hiện tại
                    if (dataItemYear == null)
                    {
                        listDataCompareMonth.Add(
                            new ReportDetailtForTotalMoneyType()
                            {
                                PartnerName = item.PartnerName,
                                Month = month.ToString(),
                                Year = year.ToString(),
                                VND = item.VND,
                                USD = item.USD,
                                EUR = item.EUR,
                                CAD = item.CAD,
                                AUD = item.AUD,
                                GBP = item.GBP,
                            }
                        );
                    }

                    // Tháng trước
                    if (dataItemLastMonth == null)
                    {
                        if (month == 1)
                        {
                            listDataCompareMonth.Add(
                                new ReportDetailtForTotalMoneyType()
                                {
                                    PartnerName = item.PartnerName,
                                    Month = "12",
                                    Year = (year - 1).ToString(),
                                    VND = item.VND,
                                    USD = item.USD,
                                    EUR = item.EUR,
                                    CAD = item.CAD,
                                    AUD = item.AUD,
                                    GBP = item.GBP,
                                }
                            );
                        }
                        else
                        {
                            listDataCompareMonth.Add(
                                new ReportDetailtForTotalMoneyType()
                                {
                                    PartnerName = item.PartnerName,
                                    Month = (month - 1).ToString(),
                                    Year = year.ToString(),
                                    VND = item.VND,
                                    USD = item.USD,
                                    EUR = item.EUR,
                                    CAD = item.CAD,
                                    AUD = item.AUD,
                                    GBP = item.GBP,
                                }
                            );
                        }
                    }

                    // Cùng kì
                    if (dataItemLastYear == null)
                    {
                        listDataCompareMonth.Add(
                            new ReportDetailtForTotalMoneyType()
                            {
                                PartnerName = item.PartnerName,
                                Month = month.ToString(),
                                Year = (year - 1).ToString(),
                                VND = item.VND,
                                USD = item.USD,
                                EUR = item.EUR,
                                CAD = item.CAD,
                                AUD = item.AUD,
                                GBP = item.GBP,
                            }
                        );
                    }
                }

            }

            // Danh sách các đối tác
            List<string> listPartnerCurrent = new List<string>();

            foreach (ReportDetailtForTotalMoneyType item in listDataCompareMonth)
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
                List<ReportDetailtForTotalMoneyType> listDataYear = listDataCompareMonth.Where(x => x.PartnerName == item && x.Month == month.ToString() && x.Year == year.ToString()).ToList();
                List<ReportDetailtForTotalMoneyType> listDataLastMonth = listDataCompareMonth.Where(x => x.PartnerName == item && x.Month == (month - 1).ToString() && x.Year == year.ToString()).ToList();
                List<ReportDetailtForTotalMoneyType> listDataMonthLastYear = listDataCompareMonth.Where(x => x.PartnerName == item && x.Month == month.ToString() && x.Year == (year - 1).ToString()).ToList();
                // Trường hợp tháng 1
                if (month == 1)
                {
                    listDataLastMonth = listDataCompareMonth.Where(x => x.PartnerName == item && x.Month == "12" && x.Year == (year - 1).ToString()).ToList();
                }
                List<ReportDetailtForTotalMoneyType> listDataLastYear = listDataCompareMonth.Where(x => x.PartnerName == item && x.Month == month.ToString() && x.Year == (year - 1).ToString()).ToList();

                listTotalRowData[i++] = string.Concat("{"
                    , string.Format("{0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9}, {10}, {11}, {12}, {13}, {14}, {15}, {16}, {17}"
                    , listDataYear.Sum(x => x.VND), listDataLastMonth.Sum(x => x.VND), listDataLastYear.Sum(x => x.VND)
                    , listDataYear.Sum(x => x.USD), listDataLastMonth.Sum(x => x.USD), listDataLastYear.Sum(x => x.USD)
                    , listDataYear.Sum(x => x.EUR), listDataLastMonth.Sum(x => x.EUR), listDataLastYear.Sum(x => x.EUR)
                    , listDataYear.Sum(x => x.CAD), listDataLastMonth.Sum(x => x.CAD), listDataLastYear.Sum(x => x.CAD)
                    , listDataYear.Sum(x => x.AUD), listDataLastMonth.Sum(x => x.AUD), listDataLastYear.Sum(x => x.AUD)
                    , listDataYear.Sum(x => x.GBP), listDataLastMonth.Sum(x => x.GBP), listDataLastYear.Sum(x => x.GBP)
                    )
                    , "}");
            }

            int count = 0;
            string yearCurent = string.Format("Năm {0}", year);
            string lastYear = string.Format("Năm {0}", year - 1);
            foreach (string item in listTotalRowData)
            {
                totalRowData = item;
                leadSourceColumn.NSeries.Add(totalRowData, true);

                categoryData = string.Concat("{"
                    , string.Format("VND {0}/{3}, VND {1}/{2}, VND {0}/{4}, USD {0}/{3}, USD {1}/{2}, USD {0}/{4}, EUR {0}/{3}, EUR {1}/{2}, EUR {0}/{4}, CAD {0}/{3}, CAD {1}/{2}, CAD {0}/{4}, AUD {0}/{3}, AUD {1}/{2}, AUD {0}/{4}, GBP {0}/{3}, GBP {1}/{2}, GBP {0}/{4}"
                    , month
                    , month == 1 ? 12 : month - 1
                    , month == 1 ? year - 1 : year
                    , year
                    , year - 1)
                    , "}");

                leadSourceColumn.NSeries.CategoryData = categoryData;

                leadSourceColumn.NSeries[count].Name = listPartnerCurrent[count];

                count++;
            }

            // Set plot area formatting as none and hide its border.
            leadSourceColumn.PlotArea.Area.FillFormat.FillType = FillType.None;
            leadSourceColumn.PlotArea.Border.IsVisible = false;

            // Format chart percent number
            //leadSourceColumn.ValueAxis.TickLabels.NumberFormat = "0.00%";


            // Set value axis major tick mark as none and hide axis line. 
            // Also set the color of value axis major grid lines.
            leadSourceColumn.ValueAxis.MajorTickMark = TickMarkType.None;
            leadSourceColumn.ValueAxis.MaxValue = 100;
            leadSourceColumn.ValueAxis.AxisLine.IsVisible = false;
            leadSourceColumn.ValueAxis.MajorGridLines.Color = Color.FromArgb(217, 217, 217);

            // Chạy process
            designer.Process();
            return ExportReport("ReportCompareForMonthForOne", designer);
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
                // Add time
                string time = string.Format("{0:yyyy-MM-dd_HH-mm-ss}", DateTime.Now);

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
                    string.Format("{0}_{1}.{2}", ReportID, time, "xlsx"));
            }
            catch (Exception ex)
            {
                MemoryStream memStream = new MemoryStream();
                return File(memStream, PDF);
            }
        }
    }
}
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
    public class ReportHSDetailtPartnerLTExcelController : Controller
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

        // GET: Admin/ReportDetailtExcelForPartnerLT
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
        public ActionResult CreateExcelForDayMonthYear(DateTime fromDate, DateTime toDate, int typeID, string reportTypeID)
        {
            WorkbookDesigner designer = new WorkbookDesigner();
            string templatePath = "~/Content/Report/ReportHSDetailt/ReportHSDetailtByPartnerLT.xlsx";
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
            List<ReportDetailtForTotalMoneyType> listDataTotal = new List<ReportDetailtForTotalMoneyType>();

            // Khởi tạo datatable
            DataTable dataTable = new DataTable();

            switch (typeID)
            {
                // Theo ngày
                case 1:

                    listReportData = new HSReportBL().SearchReportDetailtPartnerLTForDay(fromDate, toDate, reportTypeID);
                    
                    foreach (ReportDetailtForTotalMoneyType item in listReportData)
                    {
                        listDataTotal.Add(
                            new ReportDetailtForTotalMoneyType()
                            {
                                PartnerCode = item.PartnerCode,
                                PartnerName = item.PartnerName,
                                VND = item.VND,
                                USD = item.USD,
                                EUR = item.EUR,
                                CAD = item.CAD,
                                AUD = item.AUD,
                                GBP = item.GBP,
                                TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP
                            }
                        );
                    }

                    // Tạo các cột cho datatable
                    dataTable.Columns.Add("STT", typeof(String));
                    dataTable.Columns.Add("PartnerName", typeof(String));

                    dataTable.Columns.Add("VND1", typeof(double));
                    dataTable.Columns.Add("USD1", typeof(double));
                    dataTable.Columns.Add("EUR1", typeof(double));
                    dataTable.Columns.Add("CAD1", typeof(double));
                    dataTable.Columns.Add("AUD1", typeof(double));
                    dataTable.Columns.Add("GBP1", typeof(double));
                    dataTable.Columns.Add("TDS1", typeof(double));

                    List<string> listPartner = new List<string>();
                    int count = 1;

                    foreach (ReportDetailtForTotalMoneyType item in listDataTotal)
                    {
                        dataTable.Rows.Add(
                            count++
                            , item.PartnerName
                            , item.VND, item.USD, item.EUR, item.CAD, item.AUD, item.GBP, item.TongDS
                        );
                    }

                    DataRow row = dataTable.NewRow();
                    row["STT"] = "";
                    row["PartnerName"] = "Tổng";
                    row["VND1"] = dataTable.Compute("Sum(VND1)", "");
                    row["USD1"] = dataTable.Compute("Sum(USD1)", "");
                    row["EUR1"] = dataTable.Compute("Sum(EUR1)", "");
                    row["CAD1"] = dataTable.Compute("Sum(CAD1)", "");
                    row["AUD1"] = dataTable.Compute("Sum(AUD1)", "");
                    row["GBP1"] = dataTable.Compute("Sum(GBP1)", "");
                    row["TDS1"] = dataTable.Compute("Sum(TDS1)", "");
                    dataTable.Rows.Add(row);
                    
                    // Set from day and to day
                    sheetReport.Cells["E4"].PutValue(fromDate.ToString("dd/MM/yyyy"));
                    sheetReport.Cells["H4"].PutValue(toDate.ToString("dd/MM/yyyy"));
                    break;
                // Theo tháng
                case 2:

                    listReportData = new HSReportBL().SearchReportDetailtPartnerLTForMonth(fromDate, toDate, reportTypeID);

                    foreach (ReportDetailtForTotalMoneyType item in listReportData)
                    {
                        listDataTotal.Add(
                            new ReportDetailtForTotalMoneyType()
                            {
                                PartnerCode = item.PartnerCode,
                                PartnerName = item.PartnerName,
                                VND = item.VND,
                                USD = item.USD,
                                EUR = item.EUR,
                                CAD = item.CAD,
                                AUD = item.AUD,
                                GBP = item.GBP,
                                TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP
                            }
                        );
                    }

                    // Tạo các cột cho datatable
                    dataTable.Columns.Add("STT", typeof(String));
                    dataTable.Columns.Add("PartnerName", typeof(String));

                    dataTable.Columns.Add("VND1", typeof(double));
                    dataTable.Columns.Add("USD1", typeof(double));
                    dataTable.Columns.Add("EUR1", typeof(double));
                    dataTable.Columns.Add("CAD1", typeof(double));
                    dataTable.Columns.Add("AUD1", typeof(double));
                    dataTable.Columns.Add("GBP1", typeof(double));
                    dataTable.Columns.Add("TDS1", typeof(double));

                    listPartner = new List<string>();
                    count = 1;

                    foreach (ReportDetailtForTotalMoneyType item in listDataTotal)
                    {
                        dataTable.Rows.Add(
                            count++
                            , item.PartnerName
                            , item.VND, item.USD, item.EUR, item.CAD, item.AUD, item.GBP, item.TongDS
                        );
                    }

                    row = dataTable.NewRow();
                    row["STT"] = "";
                    row["PartnerName"] = "Tổng";
                    row["VND1"] = dataTable.Compute("Sum(VND1)", "");
                    row["USD1"] = dataTable.Compute("Sum(USD1)", "");
                    row["EUR1"] = dataTable.Compute("Sum(EUR1)", "");
                    row["CAD1"] = dataTable.Compute("Sum(CAD1)", "");
                    row["AUD1"] = dataTable.Compute("Sum(AUD1)", "");
                    row["GBP1"] = dataTable.Compute("Sum(GBP1)", "");
                    row["TDS1"] = dataTable.Compute("Sum(TDS1)", "");
                    dataTable.Rows.Add(row);


                    // Set from day and to day
                    sheetReport.Cells["E4"].PutValue(string.Format("{0}/{1}", fromDate.Month, fromDate.Year));
                    sheetReport.Cells["H4"].PutValue(string.Format("{0}/{1}", toDate.Month, toDate.Year));
                    break;
                // Theo năm
                default:

                    listReportData = new HSReportBL().SearchReportDetailtPartnerLTForYear(fromDate, toDate, reportTypeID);

                    foreach (ReportDetailtForTotalMoneyType item in listReportData)
                    {
                        listDataTotal.Add(
                            new ReportDetailtForTotalMoneyType()
                            {
                                PartnerCode = item.PartnerCode,
                                PartnerName = item.PartnerName,
                                VND = item.VND,
                                USD = item.USD,
                                EUR = item.EUR,
                                CAD = item.CAD,
                                AUD = item.AUD,
                                GBP = item.GBP,
                                TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP
                            }
                        );
                    }

                    // Tạo các cột cho datatable
                    dataTable.Columns.Add("STT", typeof(String));
                    dataTable.Columns.Add("PartnerName", typeof(String));

                    dataTable.Columns.Add("VND1", typeof(double));
                    dataTable.Columns.Add("USD1", typeof(double));
                    dataTable.Columns.Add("EUR1", typeof(double));
                    dataTable.Columns.Add("CAD1", typeof(double));
                    dataTable.Columns.Add("AUD1", typeof(double));
                    dataTable.Columns.Add("GBP1", typeof(double));
                    dataTable.Columns.Add("TDS1", typeof(double));

                    listPartner = new List<string>();
                    count = 1;

                    foreach (ReportDetailtForTotalMoneyType item in listDataTotal)
                    {
                        dataTable.Rows.Add(
                            count++
                            , item.PartnerName
                            , item.VND, item.USD, item.EUR, item.CAD, item.AUD, item.GBP, item.TongDS
                        );
                    }

                    row = dataTable.NewRow();
                    row["STT"] = "";
                    row["PartnerName"] = "Tổng";
                    row["VND1"] = dataTable.Compute("Sum(VND1)", "");
                    row["USD1"] = dataTable.Compute("Sum(USD1)", "");
                    row["EUR1"] = dataTable.Compute("Sum(EUR1)", "");
                    row["CAD1"] = dataTable.Compute("Sum(CAD1)", "");
                    row["AUD1"] = dataTable.Compute("Sum(AUD1)", "");
                    row["GBP1"] = dataTable.Compute("Sum(GBP1)", "");
                    row["TDS1"] = dataTable.Compute("Sum(TDS1)", "");
                    dataTable.Rows.Add(row);


                    // Set from day and to day
                    sheetReport.Cells["E4"].PutValue(fromDate.Year.ToString());
                    sheetReport.Cells["H4"].PutValue(toDate.Year.ToString());
                    break;
            }

            //// Tạo cột hearder cho Quy đô
            //string title = "Đơn vị";
            //CreateTitle("G6", "G6", sheetReport, title, 12);


            //// Tạo cột hearder cho Quy đô
            //title = "Nguyên tệ";
            //CreateTitle("H6", "H6", sheetReport, title, 12);

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
                    int totalCol = 1 + 9;
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
        public ActionResult CreateExcelForDayMonthYearForOne(DateTime fromDate, DateTime toDate, int typeID, string reportTypeID, string partnerID)
        {
            WorkbookDesigner designer = new WorkbookDesigner();
            string templatePath = "~/Content/Report/ReportHSDetailt/ReportHSDetailtByPartnerLTForOne.xlsx";
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
            List<ReportDetailtForTotalMoneyType> listData = new List<ReportDetailtForTotalMoneyType>();
            List<ReportDetailtForTotalMoneyType> listDataConvert = new List<ReportDetailtForTotalMoneyType>();
            List<ReportDetailtForTotalMoneyType> listDataTotal = new List<ReportDetailtForTotalMoneyType>();
            // Get danh sách thị trường
            List<string> listMarket = new List<string>();

            // Khởi tạo datatable
            DataTable dataTable = new DataTable();

            switch (typeID)
            {
                // Theo ngày
                case 1:

                    listData = new HSReportBL().SearchPartnerLTForOne(fromDate, toDate, reportTypeID, partnerID);

                    string namePartner = listData.Count > 0 ? listData[0].PartnerName : string.Empty;
                    CreateTitle("A3", "K3", sheetReport, string.Format("Đối tác: {0}", namePartner), 14);

                    // Tạo các cột cho datatable
                    dataTable.Columns.Add("PartnerName", typeof(String));

                    dataTable.Columns.Add("VND1", typeof(double));
                    dataTable.Columns.Add("USD1", typeof(double));
                    dataTable.Columns.Add("EUR1", typeof(double));
                    dataTable.Columns.Add("CAD1", typeof(double));
                    dataTable.Columns.Add("AUD1", typeof(double));
                    dataTable.Columns.Add("GBP1", typeof(double));

                    dataTable.Columns.Add("TDS1", typeof(double));

                    foreach (ReportDetailtForTotalMoneyType item in listData)
                    {
                        item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;
                        dataTable.Rows.Add(
                            string.Format("Ngày {0}", item.CreatedDate.ToString("dd/MM/yyyy"))
                            , item.VND, item.USD, item.EUR, item.CAD, item.AUD, item.GBP, item.TongDS
                        );
                    }

                    DataRow row = dataTable.NewRow();
                    row["PartnerName"] = "Tổng";
                    row["VND1"] = dataTable.Compute("Sum(VND1)", "");
                    row["USD1"] = dataTable.Compute("Sum(USD1)", "");
                    row["EUR1"] = dataTable.Compute("Sum(EUR1)", "");
                    row["CAD1"] = dataTable.Compute("Sum(CAD1)", "");
                    row["AUD1"] = dataTable.Compute("Sum(AUD1)", "");
                    row["GBP1"] = dataTable.Compute("Sum(GBP1)", "");

                    row["TDS1"] = dataTable.Compute("Sum(TDS1)", "");
                    dataTable.Rows.Add(row);

                    // Set from day and to day
                    sheetReport.Cells["E4"].PutValue(fromDate.ToString("dd/MM/yyyy"));
                    sheetReport.Cells["H4"].PutValue(toDate.ToString("dd/MM/yyyy"));
                    break;
                // Theo tháng
                case 2:

                    listData = new HSReportBL().SearchPartnerLTForOneForMonth(fromDate, toDate, reportTypeID, partnerID);

                    namePartner = listData.Count > 0 ? listData[0].PartnerName : string.Empty;
                    CreateTitle("A3", "K3", sheetReport, string.Format("Đối tác: {0}", namePartner), 14);

                    // Khởi tạo datatable
                    dataTable = new DataTable();
                    // Tạo các cột cho datatable
                    dataTable.Columns.Add("PartnerName", typeof(String));

                    dataTable.Columns.Add("VND1", typeof(double));
                    dataTable.Columns.Add("USD1", typeof(double));
                    dataTable.Columns.Add("EUR1", typeof(double));
                    dataTable.Columns.Add("CAD1", typeof(double));
                    dataTable.Columns.Add("AUD1", typeof(double));
                    dataTable.Columns.Add("GBP1", typeof(double));

                    dataTable.Columns.Add("TDS1", typeof(double));

                    foreach (ReportDetailtForTotalMoneyType item in listData)
                    {
                        item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;
                        dataTable.Rows.Add(
                            string.Format("Tháng {0}/{1}", item.Month, item.Year)
                            , item.VND, item.USD, item.EUR, item.CAD, item.AUD, item.GBP, item.TongDS
                        );
                    }

                    row = dataTable.NewRow();
                    row["PartnerName"] = "Tổng";
                    row["VND1"] = dataTable.Compute("Sum(VND1)", "");
                    row["USD1"] = dataTable.Compute("Sum(USD1)", "");
                    row["EUR1"] = dataTable.Compute("Sum(EUR1)", "");
                    row["CAD1"] = dataTable.Compute("Sum(CAD1)", "");
                    row["AUD1"] = dataTable.Compute("Sum(AUD1)", "");
                    row["GBP1"] = dataTable.Compute("Sum(GBP1)", "");

                    row["TDS1"] = dataTable.Compute("Sum(TDS1)", "");
                    dataTable.Rows.Add(row);


                    // Set from day and to day
                    sheetReport.Cells["E4"].PutValue(string.Format("{0}/{1}", fromDate.Month, fromDate.Year));
                    sheetReport.Cells["H4"].PutValue(string.Format("{0}/{1}", toDate.Month, toDate.Year));
                    break;
                // Theo năm
                default:

                    listData = new HSReportBL().SearchPartnerLTForOneForYear(fromDate, toDate, reportTypeID, partnerID);

                    namePartner = listData.Count > 0 ? listData[0].PartnerName : string.Empty;
                    CreateTitle("A3", "K3", sheetReport, string.Format("Đối tác: {0}", namePartner), 14);

                    // Khởi tạo datatable
                    dataTable = new DataTable();
                    // Tạo các cột cho datatable
                    dataTable.Columns.Add("PartnerName", typeof(String));

                    dataTable.Columns.Add("VND1", typeof(double));
                    dataTable.Columns.Add("USD1", typeof(double));
                    dataTable.Columns.Add("EUR1", typeof(double));
                    dataTable.Columns.Add("CAD1", typeof(double));
                    dataTable.Columns.Add("AUD1", typeof(double));
                    dataTable.Columns.Add("GBP1", typeof(double));

                    dataTable.Columns.Add("TDS1", typeof(double));

                    foreach (ReportDetailtForTotalMoneyType item in listData)
                    {
                        item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;
                        dataTable.Rows.Add(
                            string.Format("Năm {0}", item.Year)
                            , item.VND, item.USD, item.EUR, item.CAD, item.AUD, item.GBP, item.TongDS
                        );
                    }

                    row = dataTable.NewRow();
                    row["PartnerName"] = "Tổng";
                    row["VND1"] = dataTable.Compute("Sum(VND1)", "");
                    row["USD1"] = dataTable.Compute("Sum(USD1)", "");
                    row["EUR1"] = dataTable.Compute("Sum(EUR1)", "");
                    row["CAD1"] = dataTable.Compute("Sum(CAD1)", "");
                    row["AUD1"] = dataTable.Compute("Sum(AUD1)", "");
                    row["GBP1"] = dataTable.Compute("Sum(GBP1)", "");

                    row["TDS1"] = dataTable.Compute("Sum(TDS1)", "");
                    dataTable.Rows.Add(row);

                    // Set from day and to day
                    sheetReport.Cells["E4"].PutValue(fromDate.Year.ToString());
                    sheetReport.Cells["H4"].PutValue(toDate.Year.ToString());
                    break;
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

                int stepRow = 0;
                // total row = row start + số row hiện có
                // Thêm 1 vào để thêm 1 dòng thị trường
                int totalRow = dataTable.Rows.Count + 36;
                // Get total row from table 1
                totalRowTable1 = totalRow;
                // check trùng
                List<string> listDuplicate = new List<string>();
                // Số dòng của row
                for (int a = 36; a < totalRow; a++)
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


            // Vẽ biểu đồ cột cho Excel
            //Chart reference
            Aspose.Cells.Charts.Chart leadSourceColumnStacked;
            //Add Pie Chart
            // VND
            int chartIndex = sheetReport.Charts.Add(ChartType.Column3DStacked, 6, 0, 33, 15);
            leadSourceColumnStacked = sheetReport.Charts[chartIndex];

            //Chart title
            leadSourceColumnStacked.Title.Text = "Hồ sơ dịch vụ theo loại tiền";
            leadSourceColumnStacked.Title.Font.Color = Color.Silver;

            // Adding SeriesCollection (chart data source) to the chart ranging from "A1" cell to "B3"
            string totalRowData = string.Format("C37:H{0}", 36 + dataTable.Rows.Count - 1);
            leadSourceColumnStacked.NSeries.Add(totalRowData, true);

            // Set the category data covering the range A2:A5.
            // Tổng số dòng cần hiển thị là số dòng hiện tại  + số dòng trong listData -2 (trong đó 1 là dòng tổng cuối cùng)
            string categoryData = string.Format("B37:B{0}", 36 + dataTable.Rows.Count - 1);
            leadSourceColumnStacked.NSeries.CategoryData = categoryData;

            // Set the names of the chart series taken from cells.
            leadSourceColumnStacked.NSeries[0].Name = "=C36";
            leadSourceColumnStacked.NSeries[1].Name = "=D36";
            leadSourceColumnStacked.NSeries[2].Name = "=E36";
            leadSourceColumnStacked.NSeries[3].Name = "=F36";
            leadSourceColumnStacked.NSeries[4].Name = "=G36";
            leadSourceColumnStacked.NSeries[5].Name = "=H36";

            //// Set the 1st series fill color.
            //leadSourceColumnStacked.NSeries[0].Area.ForegroundColor = Color.Orange;
            //leadSourceColumnStacked.NSeries[0].Area.Formatting = FormattingType.Custom;

            //// Set the 2nd series fill color.
            //leadSourceColumnStacked.NSeries[1].Area.ForegroundColor = Color.Green;
            //leadSourceColumnStacked.NSeries[1].Area.Formatting = FormattingType.Custom;


            // Set plot area formatting as none and hide its border.
            leadSourceColumnStacked.PlotArea.Area.FillFormat.FillType = FillType.None;
            leadSourceColumnStacked.PlotArea.Border.IsVisible = false;

            // Set value axis major tick mark as none and hide axis line. 
            // Also set the color of value axis major grid lines.
            leadSourceColumnStacked.ValueAxis.MajorTickMark = TickMarkType.None;
            leadSourceColumnStacked.ValueAxis.AxisLine.IsVisible = false;
            //leadSourceColumnChiNha.ValueAxis.IsAutomaticMajorUnit = false;
            //leadSourceColumnChiNha.ValueAxis.MajorUnit = 10000000;
            leadSourceColumnStacked.ValueAxis.MajorGridLines.Color = Color.FromArgb(217, 217, 217);



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
        public ActionResult CreateExcelForGradationCompare(string gradationID, int year, string reportTypeID)
        {
            WorkbookDesigner designer = new WorkbookDesigner();
            string templatePath = "~/Content/Report/ReportHSDetailt/ReportHSDetailtByPartnerLTForGradation.xlsx";
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
            CreateTitle("A2", "W2", sheetReport, typeReport, 14);

            // Tạo title detailt
            string titleDetailt = string.Format("Giai đoạn: {0}", text);
            CreateTitle("A3", "W3", sheetReport, titleDetailt, 12);


            // Tạo giá trị cho cột dữ liệu của Chi quầy/ Chi nhà/ Chuyển khoản
            sheetReport.Cells["D7"].PutValue(string.Format("Năm {0} ", year));
            sheetReport.Cells["E7"].PutValue(string.Format("Năm {0} ", year - 1));
            sheetReport.Cells["F7"].PutValue("Tăng/Giảm");

            sheetReport.Cells["G7"].PutValue(string.Format("Năm {0} ", year));
            sheetReport.Cells["H7"].PutValue(string.Format("Năm {0} ", year - 1));
            sheetReport.Cells["I7"].PutValue("Tăng/Giảm");

            sheetReport.Cells["J7"].PutValue(string.Format("Năm {0} ", year));
            sheetReport.Cells["K7"].PutValue(string.Format("Năm {0} ", year - 1));
            sheetReport.Cells["L7"].PutValue("Tăng/Giảm");

            sheetReport.Cells["M7"].PutValue(string.Format("Năm {0} ", year));
            sheetReport.Cells["N7"].PutValue(string.Format("Năm {0} ", year - 1));
            sheetReport.Cells["O7"].PutValue("Tăng/Giảm");

            sheetReport.Cells["P7"].PutValue(string.Format("Năm {0} ", year));
            sheetReport.Cells["Q7"].PutValue(string.Format("Năm {0} ", year - 1));
            sheetReport.Cells["R7"].PutValue("Tăng/Giảm");

            sheetReport.Cells["S7"].PutValue(string.Format("Năm {0} ", year));
            sheetReport.Cells["T7"].PutValue(string.Format("Năm {0} ", year - 1));
            sheetReport.Cells["U7"].PutValue("Tăng/Giảm");

            sheetReport.Cells["V7"].PutValue(string.Format("Năm {0} ", year));
            sheetReport.Cells["W7"].PutValue(string.Format("Năm {0} ", year - 1));
            sheetReport.Cells["X7"].PutValue("Tăng/Giảm");


            List<ReportDetailtForTotalMoneyType> listDataGradation = new HSReportBL().ReportDetailtPartnerLTGradationCompareForAll(year,int.Parse(gradationID), reportTypeID);
            foreach (ReportDetailtForTotalMoneyType item in listDataGradation)
            {
                item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;
            }

            // Khởi tạo datatable
            DataTable table = new DataTable();
            // Tạo các cột cho datatable
            table.Columns.Add("STT", typeof(String));
            table.Columns.Add("PartnerName", typeof(String));

            table.Columns.Add("VND1", typeof(double));
            table.Columns.Add("VND2", typeof(double));
            table.Columns.Add("TDS1", typeof(double));

            table.Columns.Add("USD1", typeof(double));
            table.Columns.Add("USD2", typeof(double));
            table.Columns.Add("TDS2", typeof(double));

            table.Columns.Add("EUR1", typeof(double));
            table.Columns.Add("EUR2", typeof(double));
            table.Columns.Add("TDS3", typeof(double));

            table.Columns.Add("CAD1", typeof(double));
            table.Columns.Add("CAD2", typeof(double));
            table.Columns.Add("TDS4", typeof(double));

            table.Columns.Add("AUD1", typeof(double));
            table.Columns.Add("AUD2", typeof(double));
            table.Columns.Add("TDS5", typeof(double));

            table.Columns.Add("GBP1", typeof(double));
            table.Columns.Add("GBP2", typeof(double));
            table.Columns.Add("TDS6", typeof(double));

            table.Columns.Add("TDS7", typeof(double));
            table.Columns.Add("TDS8", typeof(double));
            table.Columns.Add("TDS9", typeof(double));

            int count = 1;
            List<string> listPartner = new List<string>();
            foreach (ReportDetailtForTotalMoneyType item in listDataGradation)
            {
                // Check trường hợp đã tồn tại đối tác
                if (listPartner.Contains(item.PartnerCode))
                {
                    continue;
                }

                listPartner.Add(item.PartnerCode);

                // Cùng kì
                ReportDetailtForTotalMoneyType dataItemLastYear = listDataGradation.Find(x => x.PartnerCode == item.PartnerCode && x.Year == (year - 1).ToString());
                ReportDetailtForTotalMoneyType dataItemYear = listDataGradation.Find(x => x.PartnerCode == item.PartnerCode && x.Year == year.ToString());

                // Last year
                if (dataItemLastYear == null)
                {
                    dataItemLastYear = new ReportDetailtForTotalMoneyType()
                    {
                        PartnerCode = item.PartnerCode,
                        PartnerName = item.PartnerName,
                        Year = (year - 1).ToString()
                    };
                }

                // Year hiện tại
                if (dataItemYear == null)
                {
                    dataItemYear = new ReportDetailtForTotalMoneyType()
                    {
                        PartnerCode = item.PartnerCode,
                        PartnerName = item.PartnerName,
                        Year = year.ToString()
                    };
                }

                double sumVND = dataItemYear.VND - dataItemLastYear.VND;
                double sumUSD = dataItemYear.USD - dataItemLastYear.USD;
                double sumEUR = dataItemYear.EUR - dataItemLastYear.EUR;
                double sumCAD = dataItemYear.CAD - dataItemLastYear.CAD;
                double sumAUD = dataItemYear.AUD - dataItemLastYear.AUD;
                double sumGBP = dataItemYear.GBP - dataItemLastYear.GBP;

                double sumTongDS = dataItemYear.TongDS - dataItemLastYear.TongDS;

                table.Rows.Add(
                    count++, item.PartnerName
                    , dataItemYear.VND, dataItemLastYear.VND, Math.Round(sumVND, 2, MidpointRounding.ToEven)
                    , dataItemYear.USD, dataItemLastYear.USD, Math.Round(sumUSD, 2, MidpointRounding.ToEven)
                    , dataItemYear.EUR, dataItemLastYear.EUR, Math.Round(sumEUR, 2, MidpointRounding.ToEven)
                    , dataItemYear.CAD, dataItemLastYear.CAD, Math.Round(sumCAD, 2, MidpointRounding.ToEven)
                    , dataItemYear.AUD, dataItemLastYear.AUD, Math.Round(sumAUD, 2, MidpointRounding.ToEven)
                    , dataItemYear.GBP, dataItemLastYear.GBP, Math.Round(sumGBP, 2, MidpointRounding.ToEven)

                    , dataItemYear.TongDS, dataItemLastYear.TongDS, Math.Round(sumTongDS, 2, MidpointRounding.ToEven)
                );
            }

            DataRow row = table.NewRow();
            row["STT"] = "";
            row["PartnerName"] = "Tổng";

            row["VND1"] = table.Compute("Sum(VND1)", "");
            row["VND2"] = table.Compute("Sum(VND2)", "");
            row["TDS1"] = table.Compute("Sum(TDS1)", "");

            row["USD1"] = table.Compute("Sum(USD1)", "");
            row["USD2"] = table.Compute("Sum(USD2)", "");
            row["TDS2"] = table.Compute("Sum(TDS2)", "");

            row["EUR1"] = table.Compute("Sum(EUR1)", "");
            row["EUR2"] = table.Compute("Sum(EUR2)", "");
            row["TDS3"] = table.Compute("Sum(TDS3)", "");

            row["CAD1"] = table.Compute("Sum(CAD1)", "");
            row["CAD2"] = table.Compute("Sum(CAD2)", "");
            row["TDS4"] = table.Compute("Sum(TDS4)", "");

            row["AUD1"] = table.Compute("Sum(AUD1)", "");
            row["AUD2"] = table.Compute("Sum(AUD2)", "");
            row["TDS5"] = table.Compute("Sum(TDS5)", "");

            row["GBP1"] = table.Compute("Sum(GBP1)", "");
            row["GBP2"] = table.Compute("Sum(GBP2)", "");
            row["TDS6"] = table.Compute("Sum(TDS6)", "");

            row["TDS7"] = table.Compute("Sum(TDS7)", "");
            row["TDS8"] = table.Compute("Sum(TDS8)", "");
            row["TDS9"] = table.Compute("Sum(TDS9)", "");

            table.Rows.Add(row);


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
                int totalRow = table.Rows.Count + 7;
                totalRowTable1 = totalRow;
                // Số dòng của row
                for (int a = 7; a < totalRow; a++)
                {
                    int stepColumn = 0;
                    // Số cột trong báo cáo cần hiển thị
                    // Tổng số cột hiển thị = Số cột hiển thị bắt đầu + tổng số cột cần hiển thị
                    int totalCol = 1 + 23;
                    for (int b = 1; b < totalCol; b++)
                    {
                        // Giá trị của value trong table
                        string valueOfTable = table.Rows[stepRow][stepColumn].ToString();

                        // Tô màu cho các dòng có giá trị tăng giảm
                        if (b == 5 || b == 8 || b == 11 || b == 14 || b == 17 || b == 20 || b == 23)
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

                        // Set lại color cho dòng hiện tại 
                        style.Font.Color = Color.Black;
                    }
                    // Tăng dòng của table lên
                    stepRow++;
                }
            }
            else
            {
                sheetReport.Cells["D10"].PutValue("Không có dữ liệu");
            }
            

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
        public ActionResult CreateExcelForGradationCompareForOne(string gradationID, int year, string reportTypeID, string partnerID)
        {
            WorkbookDesigner designer = new WorkbookDesigner();
            string templatePath = "~/Content/Report/ReportHSDetailt/ReportHSDetailtByPartnerLTForGradationForOne.xlsx";
            // Get đường dẫn
            templatePath = System.Web.HttpContext.Current.Server.MapPath(templatePath);

            designer.Workbook = new DongA.Core.DongAExcel().OpenExcelFile(templatePath);
            designer.Workbook.CalculateFormula();

            WorksheetCollection workSheets = designer.Workbook.Worksheets;
            Worksheet sheetReport = designer.Workbook.Worksheets[0];

            // Tạo title
            string typeReport = "So sánh - Theo giai đoạn - Từng đối tác";

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
            sheetReport.Cells["C30"].PutValue(string.Format("Lũy kế {0} năm {1} ", textValue, year));
            sheetReport.Cells["D30"].PutValue(string.Format("Lũy kế {0} năm {1} ", textValue, year - 1));
            sheetReport.Cells["E30"].PutValue("Tăng/Giảm");
            

            List<ReportDetailtForTotalMoneyType> listDataGradation = new HSReportBL().ReportDetailtPartnerLTGradationCompareForOne(year, int.Parse(gradationID), reportTypeID, partnerID);
            List<ReportDetailtForTotalMoneyType> listdataGradationConvert = new List<ReportDetailtForTotalMoneyType>();

            text = string.Empty;

            switch (gradationID)
            {
                case "1":
                    text = "3 tháng";
                    break;
                case "2":
                    text = "6 tháng";
                    break;
                case "3":
                    text = "9 tháng";
                    break;
                default:
                    text = "12 tháng";
                    break;
            }


            foreach (ReportDetailtForTotalMoneyType item in listDataGradation)
            {
                item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;

                listdataGradationConvert.Add(
                    new ReportDetailtForTotalMoneyType()
                    {
                        PartnerCode = item.PartnerCode,
                        PartnerName = item.PartnerName,
                        VND = item.VND,
                        USD = item.USD,
                        EUR = item.EUR,
                        CAD = item.CAD,
                        AUD = item.AUD,
                        GBP = item.GBP,
                        Year = item.Year
                    }

                );
            }

            // Khởi tạo datatable
            DataTable table = new DataTable();
            // Tạo các cột cho datatable
            table.Columns.Add("PartnerName", typeof(String));

            table.Columns.Add("TDS1", typeof(double));
            table.Columns.Add("TDS2", typeof(double));
            table.Columns.Add("TDS3", typeof(double));

            List<int> listPartner = new List<int>();
            string value = string.Empty;

            string[] listTypeMoney = { "VND", "USD", "EUR", "CAD", "AUD", "GBP" };

            if (listdataGradationConvert.Count > 0)
            {
                string namePartner = string.Empty;

                namePartner = listdataGradationConvert[0].PartnerName;
                CreateTitle("A4", "K4", sheetReport, string.Format("Đối tác: {0}", namePartner), 14);
                
                ReportDetailtForTotalMoneyType dataItemYear = listdataGradationConvert.Find(x => x.Year == year.ToString());
                ReportDetailtForTotalMoneyType dataItemLastYear = listdataGradationConvert.Find(x => x.Year == (year - 1).ToString());

                foreach (string item in listTypeMoney)
                {
                    var propertyInfoYear = dataItemYear.GetType().GetProperty(item);
                    var valueDataYear = propertyInfoYear.GetValue(dataItemYear, null);

                    var propertyInfoLastYear = dataItemLastYear.GetType().GetProperty(item);
                    var valueDataLastYear = propertyInfoYear.GetValue(dataItemLastYear, null);

                    double sum = Math.Round(Convert.ToDouble(valueDataYear) - Convert.ToDouble(valueDataLastYear), 2, MidpointRounding.ToEven);

                    table.Rows.Add(
                        item
                        , valueDataYear, valueDataLastYear, sum
                    );

                }
            }

            DataRow row = table.NewRow();

            row["PartnerName"] = "Tổng";

            row["TDS1"] = table.Compute("Sum(TDS1)", "");
            row["TDS2"] = table.Compute("Sum(TDS2)", "");
            row["TDS3"] = table.Compute("Sum(TDS3)", "");

            table.Rows.Add(row);


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
                int totalRow = table.Rows.Count + 30;
                totalRowTable1 = totalRow;
                // Số dòng của row
                for (int a = 30; a < totalRow; a++)
                {
                    int stepColumn = 0;
                    // Số cột trong báo cáo cần hiển thị
                    // Tổng số cột hiển thị = Số cột hiển thị bắt đầu + tổng số cột cần hiển thị
                    int totalCol = 1 + 4;
                    for (int b = 1; b < totalCol; b++)
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
                        if (b == 1)
                        {
                            style.Custom = "";
                        }
                        else
                        {
                            style.Custom = "#,##0";
                        }

                        // Tô màu cho các dòng có giá trị tăng giảm
                        if (b == 4 || b == 7 || b == 10 || b == 13 || b == 16 || b == 19)
                        {
                            decimal tryParseValue = 0;
                            decimal.TryParse(valueOfTable, out tryParseValue);
                            style.Font.Color = Color.Green;

                            if (tryParseValue < 0)
                            {
                                style.Font.Color = Color.Red;
                            }
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

                        // Set lại color cho dòng hiện tại 
                        style.Font.Color = Color.Black;
                    }
                    // Tăng dòng của table lên
                    stepRow++;
                }
            }
            else
            {
                sheetReport.Cells["D10"].PutValue("Không có dữ liệu");
            }


            // Tạo chart cột tỉ trọng cho các thị trường
            //Add Pie Chart
            //Chart reference
            Aspose.Cells.Charts.Chart leadSourceColumn;
            int chartIndex = sheetReport.Charts.Add(ChartType.Column3DClustered, 6, 0, 27, 13);
            leadSourceColumn = sheetReport.Charts[chartIndex];


            //Chart title
            leadSourceColumn.Title.Text = "Hồ sơ từng loại tiền từng đối tác";
            leadSourceColumn.Title.Font.Color = Color.Silver;
           
            // List dữ liệu dataRow
            string[] listTotalRowData = new string[listdataGradationConvert.Count];
            int i = 0;
            foreach (ReportDetailtForTotalMoneyType item in listdataGradationConvert)
            {
                listTotalRowData[i++] = string.Concat("{"
                    , string.Format("{0}, {1}, {2}, {3}, {4}, {5}"
                    , item.VND
                    , item.USD
                    , item.EUR
                    , item.CAD
                    , item.AUD
                    , item.GBP)
                    , "}");
            }

            foreach (string item in listTotalRowData)
            {
                string totalRowData = item;
                leadSourceColumn.NSeries.Add(totalRowData, true);

                string categoryData = "{VND, USD,  EUR, CAD, AUD, GBP}";
                leadSourceColumn.NSeries.CategoryData = categoryData;

            }

            i = 0;
            foreach (ReportDetailtForTotalMoneyType item in listdataGradationConvert)
            {
                leadSourceColumn.NSeries[i].Name = string.Format("Lũy kế {0} năm {1}", textValue, item.Year);
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

            // tỉ trọng
            List<ReportDetailtForTotalMoneyType> listDataGradationConvert = new HSReportBL().ReportDetailtPartnerLTGradationCompareForOne(year, int.Parse(gradationID), reportTypeID, partnerID);
            List<ReportDetailtForTotalMoneyType> listdataGradationPercent = new List<ReportDetailtForTotalMoneyType>();
            

            foreach (ReportDetailtForTotalMoneyType item in listDataGradationConvert)
            {
                item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;

                listdataGradationPercent.Add(
                    new ReportDetailtForTotalMoneyType()
                    {
                        PartnerCode = item.PartnerCode,
                        PartnerName = item.PartnerName,
                        VND = item.TongDS == 0 ? 0 : Math.Round((item.VND / item.TongDS) * 100, 2, MidpointRounding.ToEven),
                        USD = item.TongDS == 0 ? 0 : Math.Round((item.USD / item.TongDS) * 100, 2, MidpointRounding.ToEven),
                        EUR = item.TongDS == 0 ? 0 : Math.Round((item.EUR / item.TongDS) * 100, 2, MidpointRounding.ToEven),
                        CAD = item.TongDS == 0 ? 0 : Math.Round((item.CAD / item.TongDS) * 100, 2, MidpointRounding.ToEven),
                        AUD = item.TongDS == 0 ? 0 : Math.Round((item.AUD / item.TongDS) * 100, 2, MidpointRounding.ToEven),
                        GBP = item.TongDS == 0 ? 0 : Math.Round((item.GBP / item.TongDS) * 100, 2, MidpointRounding.ToEven),
                        Year = item.Year
                    }

                );
            }

            // Khởi tạo datatable
            table = new DataTable();
            // Tạo các cột cho datatable
            table.Columns.Add("PartnerName", typeof(String));

            table.Columns.Add("TDS1", typeof(double));
            table.Columns.Add("TDS2", typeof(double));
            table.Columns.Add("TDS3", typeof(double));

            listPartner = new List<int>();
            value = string.Empty;
            
            if (listdataGradationPercent.Count > 0)
            {
                ReportDetailtForTotalMoneyType dataItemYear = listdataGradationPercent.Find(x => x.Year == year.ToString());
                ReportDetailtForTotalMoneyType dataItemLastYear = listdataGradationPercent.Find(x => x.Year == (year - 1).ToString());

                foreach (string item in listTypeMoney)
                {
                    var propertyInfoYear = dataItemYear.GetType().GetProperty(item);
                    var valueDataYear = propertyInfoYear.GetValue(dataItemYear, null);

                    var propertyInfoLastYear = dataItemLastYear.GetType().GetProperty(item);
                    var valueDataLastYear = propertyInfoYear.GetValue(dataItemLastYear, null);

                    double sum = Math.Round(Convert.ToDouble(valueDataYear) - Convert.ToDouble(valueDataLastYear), 2, MidpointRounding.ToEven);

                    table.Rows.Add(
                        item
                        , valueDataYear, valueDataLastYear, sum
                    );

                }
            }

            row = table.NewRow();

            row["PartnerName"] = "Tổng";

            row["TDS1"] = 100;
            row["TDS2"] = 100;
            row["TDS3"] = 0;

            table.Rows.Add(row);


            // Tạo hearder
            string title = string.Format("Lũy kế {0} năm {1}", textValue, year);
            CreateTitle("C65", "C65", sheetReport, title, 12, true);

            title = string.Format("Lũy kế {0} năm {1}", textValue, year - 1);
            CreateTitle("D65", "D65", sheetReport, title, 12, true);

            title = "Tăng/Giảm so với cùng kì";
            CreateTitle("E65", "E65", sheetReport, title, 12, true);

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
                int totalRow = table.Rows.Count + 65;
                totalRowTable2 = totalRow;
                // Số dòng của row
                for (int a = 65; a < totalRow; a++)
                {
                    int stepColumn = 0;
                    // Số cột trong báo cáo cần hiển thị
                    // Tổng số cột hiển thị = Số cột hiển thị bắt đầu + tổng số cột cần hiển thị
                    int totalCol = 1 + 4;
                    for (int b = 1; b < totalCol; b++)
                    {
                        // Giá trị của value trong table
                        string valueOfTable = table.Rows[stepRow][stepColumn].ToString();

                        // Tô màu cho các dòng có giá trị tăng giảm
                        if (b > 3)
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

            ReportDetailtForTotalMoneyType dataPieYear = listdataGradationPercent.Find(x => x.Year == year.ToString());

            ReportDetailtForTotalMoneyType dataPieLastYear = listdataGradationPercent.Find(x => x.Year == (year - 1).ToString());

            if (dataPieLastYear != null)
            {
                //Add Pie Chart
                Aspose.Cells.Charts.Chart leadSourcePieLasYear;
                chartIndex = sheetReport.Charts.Add(ChartType.Pie3D, 40, 1, 63, 6);
                leadSourcePieLasYear = sheetReport.Charts[chartIndex];

                // Set some properties of chart plot area.
                // To set the fill color and make the border invisible.
                leadSourcePieLasYear.PlotArea.Border.IsVisible = false;
                leadSourcePieLasYear.Elevation = 45;
                // Set properties of chart title
                leadSourcePieLasYear.Title.Text = string.Format("Ti trọng từng loại tiền từng đối tác \n Giai đoạn: {0} năm {1}", textValue, year - 1);
                leadSourcePieLasYear.Title.Font.Color = Color.Silver;
                leadSourcePieLasYear.Title.Font.IsBold = true;
                leadSourcePieLasYear.Title.Font.Size = 12;

                // Set properties of nseries
                string totalRowData = string.Concat("{", string.Format("{0}, {1}, {2}, {3}, {4}, {5}", dataPieLastYear.VND, dataPieLastYear.USD, dataPieLastYear.EUR, dataPieLastYear.CAD, dataPieLastYear.AUD, dataPieLastYear.GBP), "}");
                leadSourcePieLasYear.NSeries.Add(totalRowData, true);

                string categoryData = "{VND, USD, EUR, CAD, AUD, GBP}";
                leadSourcePieLasYear.NSeries.CategoryData = categoryData;

                leadSourcePieLasYear.NSeries.IsColorVaried = true;

                // Set the DataLabels in the chart
                Aspose.Cells.Charts.DataLabels datalabels;
                for (i = 0; i < leadSourcePieLasYear.NSeries.Count; i++)
                {
                    datalabels = leadSourcePieLasYear.NSeries[i].DataLabels;
                    datalabels.Position = Aspose.Cells.Charts.LabelPositionType.InsideBase;
                    datalabels.ShowCategoryName = true;
                    //datalabels.ShowValue = true;
                    datalabels.ShowLegendKey = false;
                    datalabels.NumberFormat = "0.00%";
                    datalabels.ShowPercentage = true;
                }
            }

            if (dataPieYear != null)
            {
                //Add Pie Chart
                Aspose.Cells.Charts.Chart leadSourcePieLasYear;
                chartIndex = sheetReport.Charts.Add(ChartType.Pie3D, 40, 7, 63, 12);
                leadSourcePieLasYear = sheetReport.Charts[chartIndex];

                // Set some properties of chart plot area.
                // To set the fill color and make the border invisible.
                leadSourcePieLasYear.PlotArea.Border.IsVisible = false;
                leadSourcePieLasYear.Elevation = 45;
                // Set properties of chart title
                leadSourcePieLasYear.Title.Text = string.Format("Ti trọng từng loại tiền từng đối tác \n Giai đoạn: {0} năm {1}", text, year);
                leadSourcePieLasYear.Title.Font.Color = Color.Silver;
                leadSourcePieLasYear.Title.Font.IsBold = true;
                leadSourcePieLasYear.Title.Font.Size = 12;

                // Set properties of nseries
                string totalRowData = string.Concat("{", string.Format("{0}, {1}, {2}, {3}, {4}, {5}", dataPieYear.VND, dataPieYear.USD, dataPieYear.EUR, dataPieYear.CAD, dataPieYear.AUD, dataPieYear.GBP), "}");
                leadSourcePieLasYear.NSeries.Add(totalRowData, true);

                string categoryData = "{VND, USD, EUR, CAD, AUD, GBP}";
                leadSourcePieLasYear.NSeries.CategoryData = categoryData;

                leadSourcePieLasYear.NSeries.IsColorVaried = true;

                // Set the DataLabels in the chart
                Aspose.Cells.Charts.DataLabels datalabels;
                for (i = 0; i < leadSourcePieLasYear.NSeries.Count; i++)
                {
                    datalabels = leadSourcePieLasYear.NSeries[i].DataLabels;
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
        /// Tạo mẫu cho Excel cho so sánh theo tháng cho tất cả
        /// </summary>
        /// <param name="gradationID"></param>
        /// <param name="year"></param>
        /// <param name="typeID"></param>
        /// <returns></returns>
        public ActionResult CreateExcelForCompareForMonth(int year, int month, string reportTypeID)
        {
            WorkbookDesigner designer = new WorkbookDesigner();
            string templatePath = "~/Content/Report/ReportHSDetailt/ReportHSDetailtByPartnerLTForCompare.xlsx";
            // Get đường dẫn
            templatePath = System.Web.HttpContext.Current.Server.MapPath(templatePath);

            designer.Workbook = new DongA.Core.DongAExcel().OpenExcelFile(templatePath);
            designer.Workbook.CalculateFormula();

            WorksheetCollection workSheets = designer.Workbook.Worksheets;
            Worksheet sheetReport = designer.Workbook.Worksheets[0];

            // Tạo title
            string typeReport = "So sánh - Theo tháng - Tất cả";
            // Tạo title
            CreateTitle("A2", "W2", sheetReport, typeReport, 14);

            // Tạo title detailt
            string titleDetailt = string.Format("Tháng {0}/{1}", month, year);
            CreateTitle("A3", "W3", sheetReport, titleDetailt, 12);

            // Tạo giá trị cho cột dữ liệu của Chi quầy/ Chi nhà/ Chuyển khoản
            // VND
            sheetReport.Cells["D7"].PutValue(string.Format("Tháng {0}/{1}", month, year));
            if (month == 1)
            {
                sheetReport.Cells["E7"].PutValue(string.Format("Tháng {0}/{1}", 12, year - 1));
            }
            else
            {
                sheetReport.Cells["E7"].PutValue(string.Format("Tháng {0}/{1}", month - 1, year));
            }
            sheetReport.Cells["F7"].PutValue(string.Format("Tháng {0}/{1}", month, year - 1));

            // USD
            sheetReport.Cells["G7"].PutValue(string.Format("Tháng {0}/{1}", month, year));
            if (month == 1)
            {
                sheetReport.Cells["H7"].PutValue(string.Format("Tháng {0}/{1}", 12, year - 1));
            }
            else
            {
                sheetReport.Cells["H7"].PutValue(string.Format("Tháng {0}/{1}", month - 1, year));
            }
            sheetReport.Cells["I7"].PutValue(string.Format("Tháng {0}/{1}", month, year - 1));

            // EUR
            sheetReport.Cells["J7"].PutValue(string.Format("Tháng {0}/{1}", month, year));
            if (month == 1)
            {
                sheetReport.Cells["K7"].PutValue(string.Format("Tháng {0}/{1}", 12, year - 1));
            }
            else
            {
                sheetReport.Cells["K7"].PutValue(string.Format("Tháng {0}/{1}", month - 1, year));
            }
            sheetReport.Cells["L7"].PutValue(string.Format("Tháng {0}/{1}", month, year - 1));

            // CAD
            sheetReport.Cells["M7"].PutValue(string.Format("Tháng {0}/{1}", month, year));
            if (month == 1)
            {
                sheetReport.Cells["N7"].PutValue(string.Format("Tháng {0}/{1}", 12, year - 1));
            }
            else
            {
                sheetReport.Cells["N7"].PutValue(string.Format("Tháng {0}/{1}", month - 1, year));
            }
            sheetReport.Cells["O7"].PutValue(string.Format("Tháng {0}/{1}", month, year - 1));

            // AUD
            sheetReport.Cells["P7"].PutValue(string.Format("Tháng {0}/{1}", month, year));
            if (month == 1)
            {
                sheetReport.Cells["Q7"].PutValue(string.Format("Tháng {0}/{1}", 12, year - 1));
            }
            else
            {
                sheetReport.Cells["Q7"].PutValue(string.Format("Tháng {0}/{1}", month - 1, year));
            }
            sheetReport.Cells["R7"].PutValue(string.Format("Tháng {0}/{1}", month, year - 1));

            // GBP
            sheetReport.Cells["S7"].PutValue(string.Format("Tháng {0}/{1}", month, year));
            if (month == 1)
            {
                sheetReport.Cells["T7"].PutValue(string.Format("Tháng {0}/{1}", 12, year - 1));
            }
            else
            {
                sheetReport.Cells["T7"].PutValue(string.Format("Tháng {0}/{1}", month - 1, year));
            }
            sheetReport.Cells["U7"].PutValue(string.Format("Tháng {0}/{1}", month, year - 1));

            // Tổng
            sheetReport.Cells["V7"].PutValue(string.Format("Tháng {0}/{1}", month, year));
            if (month == 1)
            {
                sheetReport.Cells["W7"].PutValue(string.Format("Tháng {0}/{1}", 12, year - 1));
            }
            else
            {
                sheetReport.Cells["W7"].PutValue(string.Format("Tháng {0}/{1}", month - 1, year));
            }
            sheetReport.Cells["X7"].PutValue(string.Format("Tháng {0}/{1}", month, year - 1));


            List<ReportDetailtForTotalMoneyType> listDataCompareMonth = new HSReportBL().ReportDetailtPartnerLTCompareMonthForAll(year, month, reportTypeID);
            foreach (ReportDetailtForTotalMoneyType item in listDataCompareMonth)
            {
                item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;
            }

            DataTable table = new DataTable();
            // Khởi tạo datatable
            table = new DataTable();
            // Tạo các cột cho datatable
            // tháng hiện tại
            table.Columns.Add("STT", typeof(String));

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

            List<string> listPartner = new List<string>();
            int count = 1;

            foreach (ReportDetailtForTotalMoneyType item in listDataCompareMonth)
            {
                // Check tồn tại của đối tác
                if (listPartner.Contains(item.PartnerCode))
                {
                    continue;
                }
                listPartner.Add(item.PartnerCode);

                // Cùng kì
                ReportDetailtForTotalMoneyType dataItemLastYear = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == month.ToString() && x.Year == (year - 1).ToString());
                ReportDetailtForTotalMoneyType dataItemYear = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == month.ToString() && x.Year == year.ToString());
                ReportDetailtForTotalMoneyType dataItemLastMonth = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == (month - 1).ToString() && x.Year == year.ToString());
                // Trường hợp tháng 1
                if (month == 1)
                {
                    dataItemLastMonth = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == "12" && x.Year == (year - 1).ToString());
                }
                // Cung kì
                if (dataItemLastYear == null)
                {
                    dataItemLastYear = new ReportDetailtForTotalMoneyType()
                    {
                        PartnerCode = item.PartnerCode,
                        PartnerName = item.PartnerName,
                        Month = month.ToString(),
                        Year = (year - 1).ToString()
                    };
                }

                // Tháng hiện tại
                if (dataItemYear == null)
                {
                    dataItemYear = new ReportDetailtForTotalMoneyType()
                    {
                        PartnerCode = item.PartnerCode,
                        PartnerName = item.PartnerName,
                        Month = month.ToString(),
                        Year = year.ToString()
                    };
                }

                // Tháng trước
                if (dataItemLastMonth == null)
                {
                    // Trường hợp tháng 1
                    if (month == 1)
                    {
                        dataItemLastMonth = new ReportDetailtForTotalMoneyType()
                        {
                            PartnerCode = item.PartnerCode,
                            PartnerName = item.PartnerName,
                            Month = "12",
                            Year = (year - 1).ToString()
                        };
                    }
                    else
                    {
                        dataItemLastMonth = new ReportDetailtForTotalMoneyType()
                        {
                            PartnerCode = item.PartnerCode,
                            PartnerName = item.PartnerName,
                            Month = (month - 1).ToString(),
                            Year = year.ToString()
                        };
                    }

                }

                if (dataItemLastYear != null && dataItemYear != null && dataItemLastMonth != null)
                {
                    // add item vào table
                    table.Rows.Add(count++, item.PartnerName
                        , dataItemYear.VND, dataItemLastMonth.VND, dataItemLastYear.VND
                        , dataItemYear.USD, dataItemLastMonth.USD, dataItemLastYear.USD
                        , dataItemYear.EUR, dataItemLastMonth.EUR, dataItemLastYear.EUR
                        , dataItemYear.CAD, dataItemLastMonth.CAD, dataItemLastYear.CAD
                        , dataItemYear.AUD, dataItemLastMonth.AUD, dataItemLastYear.AUD
                        , dataItemYear.GBP, dataItemLastMonth.GBP, dataItemLastYear.GBP

                        , dataItemYear.TongDS, dataItemLastMonth.TongDS, dataItemLastYear.TongDS
                    );
                }
            }

            // Add dòng tổng
            DataRow row = table.NewRow();
            row["STT"] = "";
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

            // Set border
            Style style = new CellsFactory().CreateStyle();
            style.SetBorder(BorderType.BottomBorder, CellBorderType.Thin, Color.Black);
            style.SetBorder(BorderType.LeftBorder, CellBorderType.Thin, Color.Black);
            style.SetBorder(BorderType.RightBorder, CellBorderType.Thin, Color.Black);

            // Tổng số row theo table1
            int totalRowTable1 = table.Rows.Count + 7;

            if (table.Rows.Count > 0)
            {
                int stepRow = 0;
                // total row = row start + số row hiện có
                int totalRow = table.Rows.Count + 7;
                // Số dòng của row
                for (int a = 7; a < totalRow; a++)
                {
                    int stepColumn = 0;
                    // Số cột trong báo cáo cần hiển thị
                    // Tổng số cột hiển thị = Số cột hiển thị bắt đầu + tổng số cột cần hiển thị
                    int totalCol = 1 + 23;
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
                        // Set lại color cho dòng hiện tại 
                        style.Font.Color = Color.Black;
                    }
                    // Tăng dòng của table lên
                    stepRow++;

                }
            }
            else
            {
                sheetReport.Cells["D10"].PutValue("Không có dữ liệu");
            }


            string title = "Bảng so sánh dữ liệu tháng hiện tại với tháng trước và cùng kì năm trước";
            CreateTitle(string.Format("B{0}", totalRowTable1 + 6 - 2), string.Format("E{0}", totalRowTable1 + 6 - 2), sheetReport, title, 12);


            title = "STT";
            CreateTitle(string.Format("B{0}", totalRowTable1 + 6 - 1), string.Format("B{0}", totalRowTable1 + 6), sheetReport, title, 12, true);

            title = "Đối tác";
            CreateTitle(string.Format("C{0}", totalRowTable1 + 6 - 1), string.Format("C{0}", totalRowTable1 + 6), sheetReport, title, 12, true);

            title = "VND";
            CreateTitle(string.Format("D{0}", totalRowTable1 + 6 - 1), string.Format("E{0}", totalRowTable1 + 6 - 1), sheetReport, title, 12, true);

            title = "Tăng/Giảm \n so với tháng trước";
            CreateTitle(string.Format("D{0}", totalRowTable1 + 6), string.Format("D{0}", totalRowTable1 + 6), sheetReport, title, 12, true);

            title = "Tăng/Giảm \n so với cùng kì";
            CreateTitle(string.Format("E{0}", totalRowTable1 + 6), string.Format("E{0}", totalRowTable1 + 6), sheetReport, title, 12, true);

            // USD
            title = "USD";
            CreateTitle(string.Format("F{0}", totalRowTable1 + 6 - 1), string.Format("G{0}", totalRowTable1 + 6 - 1), sheetReport, title, 12, true);

            title = "Tăng/Giảm \n so với tháng trước";
            CreateTitle(string.Format("F{0}", totalRowTable1 + 6), string.Format("F{0}", totalRowTable1 + 6), sheetReport, title, 12, true);

            title = "Tăng/Giảm \n so với cùng kì";
            CreateTitle(string.Format("G{0}", totalRowTable1 + 6), string.Format("G{0}", totalRowTable1 + 6), sheetReport, title, 12, true);

            // EUR
            title = "EUR";
            CreateTitle(string.Format("H{0}", totalRowTable1 + 6 - 1), string.Format("I{0}", totalRowTable1 + 6 - 1), sheetReport, title, 12, true);

            title = "Tăng/Giảm \n so với tháng trước";
            CreateTitle(string.Format("H{0}", totalRowTable1 + 6), string.Format("H{0}", totalRowTable1 + 6), sheetReport, title, 12, true);

            title = "Tăng/Giảm \n so với cùng kì";
            CreateTitle(string.Format("I{0}", totalRowTable1 + 6), string.Format("I{0}", totalRowTable1 + 6), sheetReport, title, 12, true);

            // CAD
            title = "CAD";
            CreateTitle(string.Format("J{0}", totalRowTable1 + 6 - 1), string.Format("K{0}", totalRowTable1 + 6 - 1), sheetReport, title, 12, true);

            title = "Tăng/Giảm \n so với tháng trước";
            CreateTitle(string.Format("J{0}", totalRowTable1 + 6), string.Format("J{0}", totalRowTable1 + 6), sheetReport, title, 12, true);

            title = "Tăng/Giảm \n so với cùng kì";
            CreateTitle(string.Format("K{0}", totalRowTable1 + 6), string.Format("K{0}", totalRowTable1 + 6), sheetReport, title, 12, true);

            // AUD
            title = "AUD";
            CreateTitle(string.Format("L{0}", totalRowTable1 + 6 - 1), string.Format("M{0}", totalRowTable1 + 6 - 1), sheetReport, title, 12, true);

            title = "Tăng/Giảm \n so với tháng trước";
            CreateTitle(string.Format("L{0}", totalRowTable1 + 6), string.Format("L{0}", totalRowTable1 + 6), sheetReport, title, 12, true);

            title = "Tăng/Giảm \n so với cùng kì";
            CreateTitle(string.Format("M{0}", totalRowTable1 + 6), string.Format("M{0}", totalRowTable1 + 6), sheetReport, title, 12, true);

            // GBP
            title = "GBP";
            CreateTitle(string.Format("N{0}", totalRowTable1 + 6 - 1), string.Format("O{0}", totalRowTable1 + 6 - 1), sheetReport, title, 12, true);

            title = "Tăng/Giảm \n so với tháng trước";
            CreateTitle(string.Format("N{0}", totalRowTable1 + 6), string.Format("N{0}", totalRowTable1 + 6), sheetReport, title, 12, true);

            title = "Tăng/Giảm \n so với cùng kì";
            CreateTitle(string.Format("O{0}", totalRowTable1 + 6), string.Format("O{0}", totalRowTable1 + 6), sheetReport, title, 12, true);

            // Tổng
            title = "Tổng";
            CreateTitle(string.Format("P{0}", totalRowTable1 + 6 - 1), string.Format("Q{0}", totalRowTable1 + 6 - 1), sheetReport, title, 12, true);

            title = "Tăng/Giảm \n so với tháng trước";
            CreateTitle(string.Format("P{0}", totalRowTable1 + 6), string.Format("P{0}", totalRowTable1 + 6), sheetReport, title, 12, true);

            title = "Tăng/Giảm \n so với cùng kì";
            CreateTitle(string.Format("Q{0}", totalRowTable1 + 6), string.Format("Q{0}", totalRowTable1 + 6), sheetReport, title, 12, true);


            listDataCompareMonth = new HSReportBL().ReportDetailtPartnerLTCompareMonthForAll(year, month, reportTypeID);
            foreach (ReportDetailtForTotalMoneyType item in listDataCompareMonth)
            {
                item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;
            }
            
            table = new DataTable();
            // Tạo các cột cho datatable
            // tháng hiện tại
            table.Columns.Add("STT", typeof(String));

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

            listPartner = new List<string>();
            count = 1;

            foreach (ReportDetailtForTotalMoneyType item in listDataCompareMonth)
            {
                // Check tồn tại của đối tác
                if (listPartner.Contains(item.PartnerCode))
                {
                    continue;
                }
                listPartner.Add(item.PartnerCode);

                // Cùng kì
                ReportDetailtForTotalMoneyType dataItemLastYear = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == month.ToString() && x.Year == (year - 1).ToString());
                ReportDetailtForTotalMoneyType dataItemYear = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == month.ToString() && x.Year == year.ToString());
                ReportDetailtForTotalMoneyType dataItemLastMonth = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == (month - 1).ToString() && x.Year == year.ToString());

                // Cung kì
                if (dataItemLastYear == null)
                {
                    dataItemLastYear = new ReportDetailtForTotalMoneyType()
                    {
                        PartnerCode = item.PartnerCode,
                        PartnerName = item.PartnerName,
                        Month = month.ToString(),
                        Year = (year - 1).ToString()
                    };
                }

                // Tháng hiện tại
                if (dataItemYear == null)
                {
                    dataItemYear = new ReportDetailtForTotalMoneyType()
                    {
                        PartnerCode = item.PartnerCode,
                        PartnerName = item.PartnerName,
                        Month = month.ToString(),
                        Year = year.ToString()
                    };
                }

                // Tháng trước
                if (dataItemLastMonth == null)
                {
                    dataItemLastMonth = new ReportDetailtForTotalMoneyType()
                    {
                        PartnerCode = item.PartnerCode,
                        PartnerName = item.PartnerName,
                        Month = (month - 1).ToString(),
                        Year = year.ToString()
                    };
                }

                if (dataItemLastYear != null && dataItemYear != null && dataItemLastMonth != null)
                {
                    double sumVNDYear = Math.Round(dataItemYear.VND - dataItemLastMonth.VND, 2, MidpointRounding.ToEven);
                    double sumUSDYear = Math.Round(dataItemYear.USD - dataItemLastMonth.USD, 2, MidpointRounding.ToEven);
                    double sumEURYear = Math.Round(dataItemYear.EUR - dataItemLastMonth.EUR, 2, MidpointRounding.ToEven);
                    double sumCADYear = Math.Round(dataItemYear.CAD - dataItemLastMonth.CAD, 2, MidpointRounding.ToEven);
                    double sumAUDYear = Math.Round(dataItemYear.AUD - dataItemLastMonth.AUD, 2, MidpointRounding.ToEven);
                    double sumGBPYear = Math.Round(dataItemYear.GBP - dataItemLastMonth.GBP, 2, MidpointRounding.ToEven);

                    double sumTongDSYear = Math.Round(dataItemYear.TongDS - dataItemLastMonth.TongDS, 2, MidpointRounding.ToEven);

                    double sumVNDLastYear = Math.Round(dataItemYear.VND - dataItemLastYear.VND, 2, MidpointRounding.ToEven);
                    double sumUSDLastYear = Math.Round(dataItemYear.USD - dataItemLastYear.USD, 2, MidpointRounding.ToEven);
                    double sumEURLastYear = Math.Round(dataItemYear.EUR - dataItemLastYear.EUR, 2, MidpointRounding.ToEven);
                    double sumCADLastYear = Math.Round(dataItemYear.CAD - dataItemLastYear.CAD, 2, MidpointRounding.ToEven);
                    double sumAUDLastYear = Math.Round(dataItemYear.AUD - dataItemLastYear.AUD, 2, MidpointRounding.ToEven);
                    double sumGBPLastYear = Math.Round(dataItemYear.GBP - dataItemLastYear.GBP, 2, MidpointRounding.ToEven);

                    double sumTongDSLastYear = Math.Round(dataItemYear.TongDS - dataItemLastYear.TongDS, 2, MidpointRounding.ToEven);

                    // add item vào table
                    table.Rows.Add(count++, item.PartnerName
                        , sumVNDYear, sumVNDLastYear
                        , sumUSDYear, sumUSDLastYear
                        , sumEURYear, sumEURLastYear
                        , sumCADYear, sumCADLastYear
                        , sumAUDYear, sumAUDLastYear
                        , sumGBPYear, sumGBPLastYear
                        , sumTongDSYear, sumTongDSLastYear
                    );
                }
            }

            // Add dòng tổng
            row = table.NewRow();
            row["STT"] = "";
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

            int totalRowTable2 = totalRowTable1 + table.Rows.Count + 6;

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
                    int totalCol = 1 + 16;
                    for (int b = 1; b < totalCol; b++)
                    {
                        // Giá trị của value trong table
                        string valueOfTable = table.Rows[stepRow][stepColumn].ToString();

                        // Insert vào dòng cột xác định trong Excel
                        sheetReport.Cells[a, b].PutValue(valueOfTable, true);

                        // Tô màu cho các dòng có giá trị tăng giảm
                        if (b >= 3)
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

                        // Set lại color cho dòng hiện tại 
                        style.Font.Color = Color.Black;
                    }
                    // Tăng dòng của table lên
                    stepRow++;

                }
            }
            else
            {
                sheetReport.Cells["D10"].PutValue("Không có dữ liệu");
            }
            

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
        public ActionResult CreateExcelCompareMonthForOne(int year, int month, string reportTypeID, string partnerID)
        {
            WorkbookDesigner designer = new WorkbookDesigner();
            string templatePath = "~/Content/Report/ReportHSDetailt/ReportHSDetailtByPartnerLTForCompareForOne.xlsx";
            // Get đường dẫn
            templatePath = System.Web.HttpContext.Current.Server.MapPath(templatePath);

            designer.Workbook = new DongA.Core.DongAExcel().OpenExcelFile(templatePath);
            designer.Workbook.CalculateFormula();

            WorksheetCollection workSheets = designer.Workbook.Worksheets;
            Worksheet sheetReport = designer.Workbook.Worksheets[0];

            // Tạo title
            string typeReport = "So sánh - Theo tháng - Từng đối tác";
            // Tạo title
            CreateTitle("A2", "P2", sheetReport, typeReport, 14);

            // Tạo title detailt
            string titleDetailt = string.Format("Tháng {0}/{1}", month, year);
            CreateTitle("A3", "P3", sheetReport, titleDetailt, 12);


            // Tạo giá trị cho cột dữ liệu của Chi quầy/ Chi nhà/ Chuyển khoản

            sheetReport.Cells["C30"].PutValue(string.Format("Tháng {0}/{1} ", month, year));
            if (month == 1)
            {
                sheetReport.Cells["D30"].PutValue(string.Format("Tháng {0}/{1} ", 12, year - 1));
            }
            else
            {
                sheetReport.Cells["D30"].PutValue(string.Format("Tháng {0}/{1} ", month - 1, year));
            }

            sheetReport.Cells["E30"].PutValue(string.Format("Tháng {0}/{1} ", month, year - 1));
            

            // Get data
            List<ReportDetailtForTotalMoneyType> listDataCompareMonth = new HSReportBL().ReportDetailtPartnerLTCompareMonthForOne(year, month, reportTypeID, partnerID);

            DataTable table = new DataTable();
            // Khởi tạo datatable
            table = new DataTable();
            // Tạo các cột cho datatable
            table.Columns.Add("PartnerName", typeof(String));

            table.Columns.Add("COL1", typeof(double));
            table.Columns.Add("COL2", typeof(double));
            table.Columns.Add("COL3", typeof(double));

            table.Columns.Add("TDS1", typeof(double));
            table.Columns.Add("TDS2", typeof(double));

            string[] listTypeMoney = { "VND", "USD", "EUR", "CAD", "AUD", "GBP" };

            if (listDataCompareMonth.Count > 0)
            {
                string namePartner = listDataCompareMonth[0].PartnerName;

                CreateTitle("A4", "P4", sheetReport, string.Format("Đối tác: {0}", namePartner), 14);
                
                ReportDetailtForTotalMoneyType dataItemLastYear = listDataCompareMonth.Find(x => x.Month == month.ToString() && x.Year == (year - 1).ToString());
                ReportDetailtForTotalMoneyType dataItemYear = listDataCompareMonth.Find(x => x.Month == month.ToString() && x.Year == year.ToString());
                ReportDetailtForTotalMoneyType dataItemLastMonth = listDataCompareMonth.Find(x => x.Month == (month - 1).ToString() && x.Year == year.ToString());
                // Trường hợp tháng 1
                if (month == 1)
                {
                    dataItemLastMonth = listDataCompareMonth.Find(x => x.Month == "12" && x.Year == (year - 1).ToString());
                }

                foreach (string item in listTypeMoney)
                {
                    // Tháng hiện tại
                    var propertyInfoYear = dataItemYear.GetType().GetProperty(item);
                    var valueDataYear = propertyInfoYear.GetValue(dataItemYear, null);

                    // Tháng trước
                    var propertyInfoLastMonth = dataItemLastMonth.GetType().GetProperty(item);
                    var valueDataLastMonth = propertyInfoLastMonth.GetValue(dataItemLastMonth, null);

                    //Cùng kì
                    var propertyInfoLastYear = dataItemLastYear.GetType().GetProperty(item);
                    var valueDataLastYear = propertyInfoYear.GetValue(dataItemLastYear, null);

                    double sumYear = Math.Round(Convert.ToDouble(valueDataYear) - Convert.ToDouble(valueDataLastMonth), 2, MidpointRounding.ToEven);
                    double sumLastYear = Math.Round(Convert.ToDouble(valueDataYear) - Convert.ToDouble(valueDataLastYear), 2, MidpointRounding.ToEven);

                    table.Rows.Add(
                        item
                        , valueDataYear, valueDataLastMonth, valueDataLastYear, sumYear, sumLastYear
                    );

                }
            }

            DataRow row = table.NewRow();
            row["PartnerName"] = "Tổng";
            row["COL1"] = table.Compute("Sum(COL1)", "");
            row["COL2"] = table.Compute("Sum(COL2)", "");
            row["COL3"] = table.Compute("Sum(COL3)", "");

            row["TDS1"] = table.Compute("Sum(TDS1)", "");
            row["TDS2"] = table.Compute("Sum(TDS2)", "");
            table.Rows.Add(row);

            // Tổng số row theo table1
            int totalRowTable1 = table.Rows.Count + 30;

            // Set border
            Style style = new CellsFactory().CreateStyle();
            style.SetBorder(BorderType.BottomBorder, CellBorderType.Thin, Color.Black);
            style.SetBorder(BorderType.LeftBorder, CellBorderType.Thin, Color.Black);
            style.SetBorder(BorderType.RightBorder, CellBorderType.Thin, Color.Black);

            if (table.Rows.Count > 0)
            {
                int stepRow = 0;
                // total row = row start + số row hiện có
                int totalRow = table.Rows.Count + 30;
                // Số dòng của row
                for (int a = 30; a < totalRow; a++)
                {
                    int stepColumn = 0;
                    // Số cột trong báo cáo cần hiển thị
                    // Tổng số cột hiển thị = Số cột hiển thị bắt đầu + tổng số cột cần hiển thị
                    int totalCol = 1 + 6;
                    for (int b = 1; b < totalCol; b++)
                    {
                        // Giá trị của value trong table
                        string valueOfTable = table.Rows[stepRow][stepColumn].ToString();

                        // Insert vào dòng cột xác định trong Excel
                        sheetReport.Cells[a, b].PutValue(valueOfTable, true);

                        // set style cho number
                        style.Custom = "#,##0";

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



            // Tạo chart cột tỉ trọng cho các thị trường
            //Add Pie Chart
            //Chart reference
            Aspose.Cells.Charts.Chart leadSourceColumn;
            int chartIndex = sheetReport.Charts.Add(ChartType.Column3DClustered, 7, 1, 27, 9);
            leadSourceColumn = sheetReport.Charts[chartIndex];


            //Chart title
            leadSourceColumn.Title.Text = string.Format("Hồ sơ từng loại tiền từng đối tác \n Tháng {0}/{1}", month, year);
            leadSourceColumn.Title.Font.Color = Color.Silver;

            // List dữ liệu dataRow
            int i = 0;
            foreach (ReportDetailtForTotalMoneyType item in listDataCompareMonth)
            {

                string totalRowData = string.Concat("{", string.Format("{0}, {1}, {2}, {3}, {4}, {5}", item.VND, item.USD, item.EUR, item.CAD, item.AUD, item.GBP), "}");
                leadSourceColumn.NSeries.Add(totalRowData, true);


                string categoryData = "{VND, USD, EUR, CAD, AUD, GBP}";
                leadSourceColumn.NSeries.CategoryData = categoryData;

                leadSourceColumn.NSeries[i].Name = string.Format("Tháng {0}/{1}", item.Month, item.Year);

                // Set the 2nd series fill color.
                leadSourceColumn.NSeries[i].Area.ForegroundColor = Color.Orange;
                leadSourceColumn.NSeries[i].Area.Formatting = FormattingType.Custom;

                if (i.Equals(1))
                {
                    // Set the 1st series fill color.
                    leadSourceColumn.NSeries[i].Area.ForegroundColor = Color.Green;
                    leadSourceColumn.NSeries[i].Area.Formatting = FormattingType.Custom;
                }

                if (i.Equals(2))
                {
                    // Set the 1st series fill color.
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
            leadSourceColumn.ValueAxis.AxisLine.IsVisible = false;
            leadSourceColumn.ValueAxis.MajorGridLines.Color = Color.FromArgb(217, 217, 217);
            
            // tỉ trọng
            sheetReport.Cells["C69"].PutValue(string.Format("Tháng {0}/{1} ", month, year));
            if (month == 1)
            {
                sheetReport.Cells["D69"].PutValue(string.Format("Tháng {0}/{1} ", 12, year - 1));
            }
            else
            {
                sheetReport.Cells["D69"].PutValue(string.Format("Tháng {0}/{1} ", month - 1, year));
            }

            sheetReport.Cells["E69"].PutValue(string.Format("Tháng {0}/{1} ", month, year - 1));

            listDataCompareMonth = new HSReportBL().ReportDetailtPartnerLTCompareMonthForOne(year, month, reportTypeID, partnerID);
            List<ReportDetailtForTotalMoneyType> listDataCompareMonthConvert = new List<ReportDetailtForTotalMoneyType>();

            foreach (ReportDetailtForTotalMoneyType item in listDataCompareMonth)
            {
                item.PartnerName = string.Format("Tháng {0}/{1}", item.Month, item.Year);
                item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;

                listDataCompareMonthConvert.Add(
                    new ReportDetailtForTotalMoneyType()
                    {
                        PartnerName = string.Format("Tháng {0}/{1}", item.Month, item.Year),
                        VND = item.TongDS == 0 ? 0 : Math.Round((item.VND / item.TongDS) * 100, 2, MidpointRounding.ToEven),
                        USD = item.TongDS == 0 ? 0 : Math.Round((item.USD / item.TongDS) * 100, 2, MidpointRounding.ToEven),
                        EUR = item.TongDS == 0 ? 0 : Math.Round((item.EUR / item.TongDS) * 100, 2, MidpointRounding.ToEven),
                        CAD = item.TongDS == 0 ? 0 : Math.Round((item.CAD / item.TongDS) * 100, 2, MidpointRounding.ToEven),
                        AUD = item.TongDS == 0 ? 0 : Math.Round((item.AUD / item.TongDS) * 100, 2, MidpointRounding.ToEven),
                        GBP = item.TongDS == 0 ? 0 : Math.Round((item.GBP / item.TongDS) * 100, 2, MidpointRounding.ToEven),
                        TongDS = 100,
                        Month = item.Month,
                        Year = item.Year
                    }
                );
            }

            table = new DataTable();
            // Khởi tạo datatable
            table = new DataTable();
            // Tạo các cột cho datatable
            table.Columns.Add("PartnerName", typeof(String));

            table.Columns.Add("COL1", typeof(double));
            table.Columns.Add("COL2", typeof(double));
            table.Columns.Add("COL3", typeof(double));

            table.Columns.Add("TDS1", typeof(double));
            table.Columns.Add("TDS2", typeof(double));
            
            if (listDataCompareMonthConvert.Count > 0)
            {
                ReportDetailtForTotalMoneyType dataItemLastYear = listDataCompareMonthConvert.Find(x => x.Month == month.ToString() && x.Year == (year - 1).ToString());
                ReportDetailtForTotalMoneyType dataItemYear = listDataCompareMonthConvert.Find(x => x.Month == month.ToString() && x.Year == year.ToString());
                ReportDetailtForTotalMoneyType dataItemLastMonth = listDataCompareMonthConvert.Find(x => x.Month == (month - 1).ToString() && x.Year == year.ToString());
                // Trường hợp tháng 1
                if (month == 1)
                {
                    dataItemLastMonth = listDataCompareMonthConvert.Find(x => x.Month == "12" && x.Year == (year - 1).ToString());
                }

                foreach (string item in listTypeMoney)
                {
                    // Tháng hiện tại
                    var propertyInfoYear = dataItemYear.GetType().GetProperty(item);
                    var valueDataYear = propertyInfoYear.GetValue(dataItemYear, null);

                    // Tháng trước
                    var propertyInfoLastMonth = dataItemLastMonth.GetType().GetProperty(item);
                    var valueDataLastMonth = propertyInfoLastMonth.GetValue(dataItemLastMonth, null);

                    //Cùng kì
                    var propertyInfoLastYear = dataItemLastYear.GetType().GetProperty(item);
                    var valueDataLastYear = propertyInfoYear.GetValue(dataItemLastYear, null);

                    double sumYear = Math.Round(Convert.ToDouble(valueDataYear) - Convert.ToDouble(valueDataLastMonth), 2, MidpointRounding.ToEven);
                    double sumLastYear = Math.Round(Convert.ToDouble(valueDataYear) - Convert.ToDouble(valueDataLastYear), 2, MidpointRounding.ToEven);

                    table.Rows.Add(
                        item
                        , valueDataYear, valueDataLastMonth, valueDataLastYear, sumYear, sumLastYear
                    );

                }
            }

            row = table.NewRow();
            row["PartnerName"] = "Tổng";
            row["COL1"] = 100;
            row["COL2"] = 100;
            row["COL3"] = 100;

            row["TDS1"] = 0;
            row["TDS2"] = 0;
            table.Rows.Add(row);

            if (table.Rows.Count > 0)
            {
                int stepRow = 0;
                // total row = row start + số row hiện có
                int totalRow = table.Rows.Count + 69;
                // Số dòng của row
                for (int a = 69; a < totalRow; a++)
                {
                    int stepColumn = 0;
                    // Số cột trong báo cáo cần hiển thị
                    // Tổng số cột hiển thị = Số cột hiển thị bắt đầu + tổng số cột cần hiển thị
                    int totalCol = 1 + 6;
                    for (int b = 1; b < totalCol; b++)
                    {
                        // Giá trị của value trong table
                        string valueOfTable = table.Rows[stepRow][stepColumn].ToString();

                        // Insert vào dòng cột xác định trong Excel
                        sheetReport.Cells[a, b].PutValue(valueOfTable, true);

                        if (b >= 2)
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
                        style.Custom = "#,##0.00";

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

            // Vẻ biểu đồ
            // Vẻ biểu đồ tròn
            // Month
            ReportDetailtForTotalMoneyType dataPieMonth = listDataCompareMonthConvert.Find(x => x.Month == month.ToString() && x.Year == year.ToString());

            // last month
            ReportDetailtForTotalMoneyType dataPieLastMonth = listDataCompareMonthConvert.Find(x => x.Month == (month - 1).ToString() && x.Year == year.ToString());
            if (month == 1)
            {
                dataPieLastMonth = listDataCompareMonthConvert.Find(x => x.Month == "12" && x.Year == (year - 1).ToString());
            }

            // Last Year
            ReportDetailtForTotalMoneyType dataPieMonthLastYear = listDataCompareMonthConvert.Find(x => x.Year == (year - 1).ToString());

            if (dataPieMonth != null)
            {
                //Add Pie Chart
                Aspose.Cells.Charts.Chart leadSourcePie;
                chartIndex = sheetReport.Charts.Add(ChartType.Pie3D, 42, 1, 66, 6);
                leadSourcePie = sheetReport.Charts[chartIndex];

                // Set some properties of chart plot area.
                // To set the fill color and make the border invisible.
                leadSourcePie.PlotArea.Border.IsVisible = false;
                leadSourcePie.Elevation = 45;
                // Set properties of chart title
                leadSourcePie.Title.Text = string.Format("Tỉ trọng các đối tác theo loại tiền Tháng {0}/{1}", month, year);
                leadSourcePie.Title.Font.Color = Color.Silver;
                leadSourcePie.Title.Font.IsBold = true;
                leadSourcePie.Title.Font.Size = 12;

                string totalRowData = string.Concat("{", string.Format("{0}, {1}, {2}, {3}, {4}, {5}", dataPieMonth.VND, dataPieMonth.USD, dataPieMonth.EUR, dataPieMonth.CAD, dataPieMonth.AUD, dataPieMonth.GBP), "}");
                leadSourcePie.NSeries.Add(totalRowData, true);

                string categoryData = "{VND, USD, EUR, CAD, AUD, GBP}";
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

            // last month
            if (dataPieLastMonth != null)
            {
                //Add Pie Chart
                Aspose.Cells.Charts.Chart leadSourcePie;
                chartIndex = sheetReport.Charts.Add(ChartType.Pie3D, 42, 7, 66, 12);
                leadSourcePie = sheetReport.Charts[chartIndex];

                // Set some properties of chart plot area.
                // To set the fill color and make the border invisible.
                leadSourcePie.PlotArea.Border.IsVisible = false;
                leadSourcePie.Elevation = 45;
                // Set properties of chart title
                leadSourcePie.Title.Text = string.Format("Tỉ trọng các đối tác theo loại tiền Tháng {0}/{1}", month - 1, year);
                if (month == 1)
                {
                    leadSourcePie.Title.Text = string.Format("Tỉ trọng các đối tác theo loại tiền Tháng {0}/{1}", 12, year - 1);
                }
                leadSourcePie.Title.Font.Color = Color.Silver;
                leadSourcePie.Title.Font.IsBold = true;
                leadSourcePie.Title.Font.Size = 12;

                string totalRowData = string.Concat("{", string.Format("{0}, {1}, {2}, {3}, {4}, {5}", dataPieLastMonth.VND, dataPieLastMonth.USD, dataPieLastMonth.EUR, dataPieLastMonth.CAD, dataPieLastMonth.AUD, dataPieLastMonth.GBP), "}");
                leadSourcePie.NSeries.Add(totalRowData, true);

                string categoryData = "{VND, USD, EUR, CAD, AUD, GBP}";
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

            // last Year
            if (dataPieMonthLastYear != null)
            {
                //Add Pie Chart
                Aspose.Cells.Charts.Chart leadSourcePie;

                // cách với biểu đồ 1 20 ô
                chartIndex = sheetReport.Charts.Add(ChartType.Pie3D, 42, 13, 66, 18);
                leadSourcePie = sheetReport.Charts[chartIndex];

                // Set some properties of chart plot area.
                // To set the fill color and make the border invisible.
                leadSourcePie.PlotArea.Border.IsVisible = false;
                leadSourcePie.Elevation = 45;
                // Set properties of chart title
                leadSourcePie.Title.Text = string.Format("Tỉ trọng các đối tác theo loại tiền Tháng {0}/{1}", month, year - 1);
                leadSourcePie.Title.Font.Color = Color.Silver;
                leadSourcePie.Title.Font.IsBold = true;
                leadSourcePie.Title.Font.Size = 12;

                string totalRowData = string.Concat("{", string.Format("{0}, {1}, {2}, {3}, {4}, {5}", dataPieMonthLastYear.VND, dataPieMonthLastYear.USD, dataPieMonthLastYear.EUR, dataPieMonthLastYear.CAD, dataPieMonthLastYear.AUD, dataPieMonthLastYear.GBP), "}");
                leadSourcePie.NSeries.Add(totalRowData, true);

                string categoryData = "{VND, USD, EUR, CAD, AUD, GBP}";
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
                styleTitle.IsTextWrapped = true;
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
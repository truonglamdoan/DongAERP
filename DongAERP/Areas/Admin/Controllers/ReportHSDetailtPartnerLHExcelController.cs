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
    public class ReportHSDetailtPartnerLHExcelController : Controller
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

        // GET: Admin/ReportDetailtExcelForPartner
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
            string templatePath = "~/Content/Report/ReportHSDetailt/ReportHSDetailtByPartnerLH.xlsx";
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
            List<ReportDetailtForPartner> listReportData = new List<ReportDetailtForPartner>();

            switch (typeID)
            {
                // Theo ngày
                case 1:
                    listReportData = new HSReportBL().SearchReportDetailtPartnerForDay(fromDate, toDate, reportTypeID);

                    int count = 1;
                    foreach (ReportDetailtForPartner item in listReportData)
                    {
                        item.STT = (count++).ToString();
                        item.ReportID = item.PartnerName;
                        item.TongDS = item.DSChiNha + item.DSChiQuay + item.DSCK;
                    }
                    
                    // Set from day and to day
                    sheetReport.Cells["E4"].PutValue(fromDate.ToString("dd/MM/yyyy"));
                    sheetReport.Cells["G4"].PutValue(toDate.ToString("dd/MM/yyyy"));
                    break;
                // Theo tháng
                case 2:

                    listReportData = new HSReportBL().SearchReportDetailtPartnerForMonth(fromDate, toDate, reportTypeID);

                    count = 1;
                    foreach (ReportDetailtForPartner item in listReportData)
                    {
                        item.STT = (count++).ToString();
                        item.ReportID = item.PartnerName;
                        item.TongDS = item.DSChiNha + item.DSChiQuay + item.DSCK;
                    }
                    // Set from day and to day
                    sheetReport.Cells["E4"].PutValue(string.Format("{0}/{1}", fromDate.Month, fromDate.Year));
                    sheetReport.Cells["G4"].PutValue(string.Format("{0}/{1}", toDate.Month, toDate.Year));
                    break;
                // Theo năm
                default:

                    listReportData = new HSReportBL().SearchReportDetailtPartnerForYear(fromDate, toDate, reportTypeID);

                    count = 1;
                    foreach (ReportDetailtForPartner item in listReportData)
                    {
                        item.STT = (count++).ToString();
                        item.ReportID = item.PartnerName;
                        item.TongDS = item.DSChiNha + item.DSChiQuay + item.DSCK;
                    }
                    // Set from day and to day
                    sheetReport.Cells["E4"].PutValue(fromDate.Year.ToString());
                    sheetReport.Cells["G4"].PutValue(toDate.Year.ToString());
                    break;
            }

            DataTable dataTable = new DataTable();

            if (listReportData.Count > 0)
            {
                // Add dòng tổng vào list danh sách
                ReportDetailtForPartner reportForMaket = new ReportDetailtForPartner()
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
                            style.HorizontalAlignment = TextAlignmentType.Center;
                        }
                        else
                        {
                            style.Custom = "#,##0";
                            style.HorizontalAlignment = TextAlignmentType.General;
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
            string templatePath = "~/Content/Report/ReportHSDetailt/ReportHSDetailtByPartnerLHForOne.xlsx";
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
            List<ReportDetailtForPartner> listReportData = new List<ReportDetailtForPartner>();
            // Get danh sách thị trường
            List<string> listMarket = new List<string>();

            switch (typeID)
            {
                // Theo ngày
                case 1:
                    listReportData = new HSReportBL().SearchPartnerForOne(fromDate, toDate, reportTypeID, partnerID);
                    
                    foreach (ReportDetailtForPartner item in listReportData)
                    {
                        item.ReportID = item.PartnerName;
                        item.STT = string.Format("Ngày {0}", item.CreatedDate.ToString("dd/MM/yyyy"));
                        item.TongDS = item.DSChiNha + item.DSChiQuay + item.DSCK;
                    }

                    listReportData.Add(
                        new ReportDetailtForPartner()
                        {
                            STT = "Tổng",
                            DSChiQuay = listReportData.Sum(x => x.DSChiQuay),
                            DSChiNha = listReportData.Sum(x => x.DSChiNha),
                            DSCK = listReportData.Sum(x => x.DSCK),
                            TongDS = listReportData.Sum(x => x.TongDS)
                        }
                    );

                    // Set from day and to day
                    sheetReport.Cells["E4"].PutValue(fromDate.ToString("dd/MM/yyyy"));
                    sheetReport.Cells["G4"].PutValue(toDate.ToString("dd/MM/yyyy"));
                    break;
                // Theo tháng
                case 2:

                    listReportData = new HSReportBL().SearchPartnerForOneForMonth(fromDate, toDate, reportTypeID, partnerID);

                    foreach (ReportDetailtForPartner item in listReportData)
                    {
                        item.ReportID = item.PartnerName;
                        item.STT = string.Format("Tháng {0}/{1}", item.Month, item.Year);
                        item.TongDS = item.DSChiNha + item.DSChiQuay + item.DSCK;
                    }

                    listReportData.Add(
                        new ReportDetailtForPartner()
                        {
                            STT = "Tổng",
                            DSChiQuay = listReportData.Sum(x => x.DSChiQuay),
                            DSChiNha = listReportData.Sum(x => x.DSChiNha),
                            DSCK = listReportData.Sum(x => x.DSCK),
                            TongDS = listReportData.Sum(x => x.TongDS)
                        }
                    );
                    // Set from day and to day
                    sheetReport.Cells["E4"].PutValue(string.Format("{0}/{1}", fromDate.Month, fromDate.Year));
                    sheetReport.Cells["G4"].PutValue(string.Format("{0}/{1}", toDate.Month, toDate.Year));
                    break;
                // Theo năm
                default:

                    listReportData = new HSReportBL().SearchPartnerForOneForYear(fromDate, toDate, reportTypeID, partnerID);

                    foreach (ReportDetailtForPartner item in listReportData)
                    {
                        item.ReportID = item.PartnerName;
                        item.STT = string.Format("Năm {0}", item.Year);
                        item.TongDS = item.DSChiNha + item.DSChiQuay + item.DSCK;
                    }

                    listReportData.Add(
                        new ReportDetailtForPartner()
                        {
                            STT = "Tổng",
                            DSChiQuay = listReportData.Sum(x => x.DSChiQuay),
                            DSChiNha = listReportData.Sum(x => x.DSChiNha),
                            DSCK = listReportData.Sum(x => x.DSCK),
                            TongDS = listReportData.Sum(x => x.TongDS)
                        }
                    );

                    // Set from day and to day
                    sheetReport.Cells["E4"].PutValue(fromDate.Year.ToString());
                    sheetReport.Cells["G4"].PutValue(toDate.Year.ToString());
                    break;
            }

            DataTable dataTable = new DataTable();

            if (listReportData.Count > 0)
            {
                // Tạo các col cho table
                dataTable = CreateDataTableFormart(false);

                // Danh sách dataSet của báo cáo ngày
                DataSet dataReport = ConvertListObjectToDataSet(listReportData);

                // Đổ data vào datatble mới
                FillData(dataReport.Tables[0], dataTable, false);
            }

            // Set border
            Style style = new CellsFactory().CreateStyle();
            style.SetBorder(BorderType.TopBorder, CellBorderType.Thin, Color.Black);
            style.SetBorder(BorderType.BottomBorder, CellBorderType.Thin, Color.Black);
            style.SetBorder(BorderType.LeftBorder, CellBorderType.Thin, Color.Black);
            style.SetBorder(BorderType.RightBorder, CellBorderType.Thin, Color.Black);

            if (dataTable.Rows.Count > 0)
            {
                int stepRow = 0;
                // total row = row start + số row hiện có
                // Thêm 1 vào để thêm 1 dòng thị trường
                int totalRow = dataTable.Rows.Count + 35;
                // check trùng
                List<string> listDuplicate = new List<string>();
                // Số dòng của row
                for (int a = 35; a < totalRow; a++)
                {
                    int stepColumn = 0;
                    // Số cột trong báo cáo cần hiển thị
                    // Tổng số cột hiển thị = Số cột hiển thị bắt đầu + tổng số cột cần hiển thị
                    // Số 6 là cột marketName
                    int totalCol = 1 + 5;

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
                            style.HorizontalAlignment = TextAlignmentType.Center;
                        }
                        else
                        {
                            style.Custom = "#,##0";
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

            }
            else
            {
                sheetReport.Cells["D10"].PutValue("Không có dữ liệu");
            }

            // Vẽ biểu đồ cột cho Excel
            //Chart reference
            Aspose.Cells.Charts.Chart leadSourceLine;
            //Add Pie Chart
            // Chi Nhà
            int chartIndex = sheetReport.Charts.Add(ChartType.Line, 7, 0, 30, 10);
            leadSourceLine = sheetReport.Charts[chartIndex];

            //Chart title
            leadSourceLine.Title.Text = string.Format("Biểu đồ đường loại hình của từng đối tác");
            leadSourceLine.Title.Font.Color = Color.Silver;

            // Adding SeriesCollection (chart data source) to the chart ranging from "A1" cell to "B3"
            string totalRowData = string.Format("C36:E{0}", 36 + dataTable.Rows.Count - 2);
            leadSourceLine.NSeries.Add(totalRowData, true);

            // Set the category data covering the range A2:A5.
            // Tổng số dòng cần hiển thị là số dòng hiện tại  + số dòng trong listData -2 (trong đó 1 là dòng tổng cuối cùng)
            string categoryData = string.Format("B36:B{0}", 36 + dataTable.Rows.Count - 2);
            leadSourceLine.NSeries.CategoryData = categoryData;

            leadSourceLine.CategoryAxis.TickLabels.RotationAngle = 15;

            // Set the names of the chart series taken from cells.
            leadSourceLine.NSeries[0].Name = "=C35";
            leadSourceLine.NSeries[1].Name = "=D35";
            leadSourceLine.NSeries[2].Name = "=E35";

            // Set the 1st series fill color.
            leadSourceLine.NSeries[0].Area.ForegroundColor = Color.Orange;
            leadSourceLine.NSeries[0].Area.Formatting = FormattingType.Custom;

            // Set the 2nd series fill color.
            leadSourceLine.NSeries[1].Area.ForegroundColor = Color.Green;
            leadSourceLine.NSeries[1].Area.Formatting = FormattingType.Custom;


            // Set the 2nd series fill color.
            leadSourceLine.NSeries[2].Area.ForegroundColor = Color.Blue;
            leadSourceLine.NSeries[2].Area.Formatting = FormattingType.Custom;


            // Set plot area formatting as none and hide its border.
            leadSourceLine.PlotArea.Area.FillFormat.FillType = FillType.None;
            leadSourceLine.PlotArea.Border.IsVisible = false;

            // Set value axis major tick mark as none and hide axis line. 
            // Also set the color of value axis major grid lines.
            leadSourceLine.ValueAxis.MajorTickMark = TickMarkType.None;
            leadSourceLine.ValueAxis.AxisLine.IsVisible = false;
            //leadSourceColumnChiNha.ValueAxis.IsAutomaticMajorUnit = false;
            //leadSourceColumnChiNha.ValueAxis.MajorUnit = 10000000;
            leadSourceLine.ValueAxis.MajorGridLines.Color = Color.FromArgb(217, 217, 217);

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
            string templatePath = "~/Content/Report/ReportHSDetailt/ReportHSDetailtByPartnerLHForGradation.xlsx";
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


            // Tạo giá trị cho cột dữ liệu của Chi quầy/ Chi nhà/ Chuyển khoản
            sheetReport.Cells["D7"].PutValue(string.Format("Năm {0} ", year));
            sheetReport.Cells["E7"].PutValue(string.Format("Năm {0} ", year - 1));


            sheetReport.Cells["G7"].PutValue(string.Format("Năm {0} ", year));
            sheetReport.Cells["H7"].PutValue(string.Format("Năm {0} ", year - 1));


            sheetReport.Cells["J7"].PutValue(string.Format("Năm {0} ", year));
            sheetReport.Cells["K7"].PutValue(string.Format("Năm {0} ", year - 1));


            sheetReport.Cells["M7"].PutValue(string.Format("Năm {0} ", year));
            sheetReport.Cells["N7"].PutValue(string.Format("Năm {0} ", year - 1));

            List<ReportDetailtForPartner> listDataGradation = new HSReportBL().ReportDetailtPartnerGradationCompareForAll(year, int.Parse(gradationID), reportTypeID);

            // Khởi tạo datatable
            DataTable table = new DataTable();
            foreach (ReportDetailtForPartner item in listDataGradation)
            {
                item.TongDS = item.DSChiQuay + item.DSChiNha + item.DSCK;
            }

            // Khởi tạo datatable
            table = new DataTable();
            // Tạo các cột cho datatable
            table.Columns.Add("STT", typeof(String));
            table.Columns.Add("PartnerName", typeof(String));

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

            if (listDataGradation.Count() > 0)
            {
                List<string> listPartnerName = new List<string>();
                int count = 1;
                foreach (ReportDetailtForPartner item in listDataGradation)
                {
                    if (listPartnerName.Contains(item.PartnerName))
                    {
                        continue;
                    }
                    // Thêm đối tác vào bảng tạm
                    listPartnerName.Add(item.PartnerName);

                    // Cùng kì
                    ReportDetailtForPartner dataItemLastYear = listDataGradation.Find(x => x.PartnerCode == item.PartnerCode && x.Year == (year - 1).ToString());
                    ReportDetailtForPartner dataItemYear = listDataGradation.Find(x => x.PartnerCode == item.PartnerCode && x.Year == year.ToString());

                    // Trường hợp năm hiện tại không có
                    if (dataItemLastYear == null)
                    {
                        dataItemLastYear = new ReportDetailtForPartner()
                        {
                            PartnerCode = item.PartnerCode,
                            PartnerName = item.PartnerName,
                            Year = (year - 1).ToString()
                        };
                    }

                    // Trường hợp năm hiện tại không có
                    if (dataItemYear == null)
                    {
                        dataItemYear = new ReportDetailtForPartner()
                        {
                            PartnerCode = item.PartnerCode,
                            PartnerName = item.PartnerName,
                            Year = year.ToString()
                        };
                    }
                    // add item vào table
                    table.Rows.Add(
                        count++
                        , item.PartnerName
                        , dataItemYear.DSChiQuay, dataItemLastYear.DSChiQuay, Math.Round(dataItemYear.DSChiQuay - dataItemLastYear.DSChiQuay, 2, MidpointRounding.ToEven)
                        , dataItemYear.DSChiNha, dataItemLastYear.DSChiNha, Math.Round(dataItemYear.DSChiNha - dataItemLastYear.DSChiNha, 2, MidpointRounding.ToEven)
                        , dataItemYear.DSCK, dataItemLastYear.DSCK, Math.Round(dataItemYear.DSCK - dataItemLastYear.DSCK, 2, MidpointRounding.ToEven)
                        , dataItemYear.TongDS, dataItemLastYear.TongDS, Math.Round(dataItemYear.TongDS - dataItemLastYear.TongDS, 2, MidpointRounding.ToEven)
                    );
                }

                DataRow row = table.NewRow();
                row["STT"] = "";
                row["PartnerName"] = "Tổng";
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
                int totalRow = table.Rows.Count + 7;
                // Số dòng của row
                for (int a = 7; a < totalRow; a++)
                {
                    int stepColumn = 0;
                    // Số cột trong báo cáo cần hiển thị
                    // Tổng số cột hiển thị = Số cột hiển thị bắt đầu + tổng số cột cần hiển thị
                    int totalCol = 1 + 14;
                    for (int b = 1; b < totalCol; b++)
                    {
                        // Giá trị của value trong table
                        string valueOfTable = table.Rows[stepRow][stepColumn].ToString();

                        // Tô màu cho các dòng có giá trị tăng giảm
                        if (b == 5 || b == 8 || b == 11 || b == 14)
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

                        if (b.Equals(totalCol - 2))
                        {
                            sheetReport.Cells[a, b].PutValue(valueOfTable, true, true);
                            style.Font.IsBold = true;
                            sheetReport.Cells[a, b].SetStyle(style);
                        }

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
                        style.Font.Color = Color.Black;
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
        public ActionResult CreateExcelGradationCompareForOne(string gradationID, int year, string reportTypeID, string partnerID)
        {
            WorkbookDesigner designer = new WorkbookDesigner();

            string templatePath = "~/Content/Report/ReportHSDetailt/ReportHSDetailtByPartnerLHForGradationOne.xlsx";
            // Get đường dẫn
            templatePath = System.Web.HttpContext.Current.Server.MapPath(templatePath);

            designer.Workbook = new DongA.Core.DongAExcel().OpenExcelFile(templatePath);
            designer.Workbook.CalculateFormula();

            WorksheetCollection workSheets = designer.Workbook.Worksheets;
            Worksheet sheetReport = designer.Workbook.Worksheets[0];

            // Tạo title
            string typeReport = "So sánh - Theo giai đoạn - Từng thị trường";

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

            List<ReportDetailtForPartner> listDataGradation = new HSReportBL().ReportDetailtPartnerGradationCompareForOne(year, int.Parse(gradationID), reportTypeID, partnerID);

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

            foreach (ReportDetailtForPartner item in listDataGradation)
            {
                item.PartnerName = string.Format("Lũy kế {0} năm {1}", text, item.Year);
                item.TongDS = item.DSChiQuay + item.DSChiNha + item.DSCK;
            }

            double sumDSChiQuay = 0;
            double sumDSChiNha = 0;
            double sumDSCK = 0;
            double sumTongDS = 0;

            if (listDataGradation.Count().Equals(2))
            {
                sumDSChiQuay = listDataGradation[1].DSChiQuay - listDataGradation[0].DSChiQuay;
                sumDSChiNha = listDataGradation[1].DSChiNha - listDataGradation[0].DSChiNha;
                sumDSCK = listDataGradation[1].DSCK - listDataGradation[0].DSCK;
                sumTongDS = listDataGradation[1].TongDS - listDataGradation[0].TongDS;

                // Tăng giảm
                listDataGradation.Add(
                    new ReportDetailtForPartner()
                    {
                        PartnerName = string.Format("Tăng giảm so với cùng kì {0} (+/-)", year - 1),
                        DSChiQuay = Math.Round(sumDSChiQuay, 2, MidpointRounding.ToEven),
                        DSChiNha = Math.Round(sumDSChiNha, 2, MidpointRounding.ToEven),
                        DSCK = Math.Round(sumDSCK, 2, MidpointRounding.ToEven),
                        TongDS = Math.Round(sumTongDS, 2, MidpointRounding.ToEven),
                    }
                );

                // tỉ trọng
                listDataGradation.Add(
                    new ReportDetailtForPartner()
                    {
                        PartnerName = string.Format("Tăng giảm so với cùng kì {0} (%)", year - 1),
                        DSChiQuay = listDataGradation[0].DSChiQuay == 0 ? 0 : Math.Round((sumDSChiQuay / listDataGradation[0].DSChiQuay) * 100, 2, MidpointRounding.ToEven),
                        DSChiNha = listDataGradation[0].DSChiNha == 0 ? 0 : Math.Round((sumDSChiNha / listDataGradation[0].DSChiNha) * 100, 2, MidpointRounding.ToEven),
                        DSCK = listDataGradation[0].DSCK == 0 ? 0 : Math.Round((sumDSCK / listDataGradation[0].DSCK) * 100, 2, MidpointRounding.ToEven),
                        TongDS = listDataGradation[0].TongDS == 0 ? 0 : Math.Round((sumTongDS / listDataGradation[0].TongDS) * 100, 2, MidpointRounding.ToEven),
                    }
                );
            }

            // Khởi tạo datatable
            DataTable table = new DataTable();
            // Tạo các cột cho datatable
            table.Columns.Add("STT", typeof(String));

            table.Columns.Add("CQ1", typeof(double));
            table.Columns.Add("CN1", typeof(double));
            table.Columns.Add("CK1", typeof(double));
            table.Columns.Add("TDS1", typeof(double));

            foreach (ReportDetailtForPartner item in listDataGradation)
            {
                table.Rows.Add(item.PartnerName, item.DSChiQuay, item.DSChiNha, item.DSCK, item.TongDS);
            }

            // Set border
            Style style = new CellsFactory().CreateStyle();
            style.SetBorder(BorderType.BottomBorder, CellBorderType.Thin, Color.Black);
            style.SetBorder(BorderType.LeftBorder, CellBorderType.Thin, Color.Black);
            style.SetBorder(BorderType.RightBorder, CellBorderType.Thin, Color.Black);
            style.SetBorder(BorderType.TopBorder, CellBorderType.Thin, Color.Black);

            // Tổng số row theo table1
            int totalRowTable1 = table.Rows.Count + 35;

            // Table dữ liệu bảng số liệu Doanh số Chi Quầy/Chi Nhà/Chuyển Khoản
            if (table.Rows.Count > 0)
            {
                int stepRow = 0;
                // total row = row start + số row hiện có
                int totalRow = table.Rows.Count + 35;
                // Số dòng của row
                for (int a = 35; a < totalRow; a++)
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

                        // Trường hợp thuộc 3 dòng cuối
                        if (a.Equals(totalRow - 1))
                        {
                            sheetReport.Cells[a, b].PutValue(valueOfTable, true, true);
                            style.Font.IsBold = true;
                            sheetReport.Cells[a, b].SetStyle(style);
                        }

                        if (a.Equals(totalRow - 2))
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
            Aspose.Cells.Charts.Chart leadSourceColumn;
            //Add Pie Chart
            // Chi Nhà
            int chartIndex = sheetReport.Charts.Add(ChartType.Column3DClustered, 8, 0, 32, 7);
            leadSourceColumn = sheetReport.Charts[chartIndex];

            //Chart title
            leadSourceColumn.Title.Text = string.Format("Doanh số dịch vụ chi quầy từng thị trường \n Giai đoạn: {0}", text);
            leadSourceColumn.Title.Font.Color = Color.Silver;


            // List dữ liệu dataRow
            string[] listTotalRowData = new string[2];
            
            ReportDetailtForPartner dataItemYear = listDataGradation.Find(x => x.Year == year.ToString());
            ReportDetailtForPartner dataItemLastYear = listDataGradation.Find(x => x.Year == (year - 1).ToString());

            if (dataItemYear != null && dataItemLastYear != null)
            {
                listTotalRowData[0] = string.Concat("{", string.Format("{0}, {1}, {2}", dataItemYear.DSChiQuay, dataItemYear.DSChiNha, dataItemYear.DSCK), "}");
                listTotalRowData[1] = string.Concat("{", string.Format("{0}, {1}, {2}", dataItemLastYear.DSChiQuay, dataItemLastYear.DSChiNha, dataItemLastYear.DSCK), "}");
            }
            
            foreach(string item in listTotalRowData)
            {
                leadSourceColumn.NSeries.Add(item, true);
                
                leadSourceColumn.NSeries.CategoryData = "{DS Chi Quầy, DS Chi Nhà, DS Chuyển Khoản}";
            }


            leadSourceColumn.NSeries[0].Name = string.Format("Lũy kế {0} {1}", text, year);
            leadSourceColumn.NSeries[1].Name = string.Format("Lũy kế {0} {1}", text, year - 1);

            // Set the 1st series fill color.
            leadSourceColumn.NSeries[0].Area.ForegroundColor = Color.Orange;
            leadSourceColumn.NSeries[0].Area.Formatting = FormattingType.Custom;

            // Set the 2nd series fill color.
            leadSourceColumn.NSeries[1].Area.ForegroundColor = Color.Green;
            leadSourceColumn.NSeries[1].Area.Formatting = FormattingType.Custom;
            
            // Set plot area formatting as none and hide its border.
            leadSourceColumn.PlotArea.Area.FillFormat.FillType = FillType.None;
            leadSourceColumn.PlotArea.Border.IsVisible = false;

            // Set value axis major tick mark as none and hide axis line. 
            // Also set the color of value axis major grid lines.
            leadSourceColumn.ValueAxis.MajorTickMark = TickMarkType.None;
            leadSourceColumn.ValueAxis.AxisLine.IsVisible = false;
            leadSourceColumn.ValueAxis.MajorGridLines.Color = Color.FromArgb(217, 217, 217);

            // Hiển thị % cho việc tạo bảng tỉ trọng của từng thị trường
            listDataGradation = new HSReportBL().ReportDetailtPartnerGradationCompareForOne(year, int.Parse(gradationID), reportTypeID, partnerID);
            List<ReportDetailtForPartner> listDataGradationConvert = new List<ReportDetailtForPartner>();
            
            foreach (ReportDetailtForPartner item in listDataGradation)
            {
                item.PartnerName = string.Format("Lũy kế {0} năm {1}", text, item.Year);
                item.TongDS = item.DSChiQuay + item.DSChiNha + item.DSCK;

                listDataGradationConvert.Add(
                    new ReportDetailtForPartner()
                    {
                        PartnerCode = item.PartnerCode,
                        PartnerName = item.PartnerName,
                        DSChiQuay = item.TongDS == 0 ? 0 : Math.Round((item.DSChiQuay / item.TongDS) * 100, 2, MidpointRounding.ToEven),
                        DSChiNha = item.TongDS == 0 ? 0 : Math.Round((item.DSChiNha / item.TongDS) * 100, 2, MidpointRounding.ToEven),
                        DSCK = item.TongDS == 0 ? 0 : Math.Round((item.DSCK / item.TongDS) * 100, 2, MidpointRounding.ToEven),
                        TongDS = 100,
                        Year = item.Year
                    }
                );
            }
            
            if (listDataGradation.Count().Equals(2))
            {
                sumDSChiQuay = listDataGradation[1].DSChiQuay - listDataGradation[0].DSChiQuay;
                sumDSChiNha = listDataGradation[1].DSChiNha - listDataGradation[0].DSChiNha;
                sumDSCK = listDataGradation[1].DSCK - listDataGradation[0].DSCK;
                sumTongDS = listDataGradation[1].TongDS - listDataGradation[0].TongDS;
            }

            if (listDataGradationConvert.Count().Equals(2))
            {
                listDataGradationConvert.Add(
                    new ReportDetailtForPartner()
                    {
                        PartnerName = string.Format("Tăng giảm so với cùng kì {0}", year - 1),
                        DSChiQuay = listDataGradationConvert[1].DSChiQuay - listDataGradationConvert[0].DSChiQuay,
                        DSChiNha = listDataGradationConvert[1].DSChiNha - listDataGradationConvert[0].DSChiNha,
                        DSCK = listDataGradationConvert[1].DSCK - listDataGradationConvert[0].DSCK,
                    }
                );
            }

            // Xóa dữ liệu trong datatable
            table.Rows.Clear();

            foreach (ReportDetailtForPartner item in listDataGradationConvert)
            {
                table.Rows.Add(item.PartnerName, item.DSChiQuay, item.DSChiNha, item.DSCK, item.TongDS);
            }



            // Tổng số row của table2
            // Với 6 là số cách của table1 và table2
            int totalRowTable2 = totalRowTable1 + table.Rows.Count + 26;
            // Tạo title hearder cho table tăng giảm
            // Title cho thị trường
            string title = "2. Theo tỷ trọng doanh số đối tác - loại hình dịch vụ";

            CreateTitle(string.Format("B{0}", totalRowTable1 + 6 - 3), string.Format("C{0}", totalRowTable1 + 6 - 3), sheetReport, title, 12);

            title = "";
            CreateTitle(string.Format("B{0}", totalRowTable1 + 26), string.Format("B{0}", totalRowTable1 + 26), sheetReport, title, 12, true);

            title = string.Format("Chi Quầy");
            CreateTitle(string.Format("C{0}", totalRowTable1 + 26), string.Format("C{0}", totalRowTable1 + 26), sheetReport, title, 12, true);

            title = string.Format("Chi Nhà");
            CreateTitle(string.Format("D{0}", totalRowTable1 + 26), string.Format("D{0}", totalRowTable1 + 26), sheetReport, title, 12, true);

            title = string.Format("Chuyển Khoản");
            CreateTitle(string.Format("E{0}", totalRowTable1 + 26), string.Format("E{0}", totalRowTable1 + 26), sheetReport, title, 12, true);

            title = string.Format("Tổng");
            CreateTitle(string.Format("F{0}", totalRowTable1 + 26), string.Format("F{0}", totalRowTable1 + 26), sheetReport, title, 12, true);


            // Table dữ liệu bảng số liệu Doanh số Chi Quầy/Chi Nhà/Chuyển Khoản
            if (table.Rows.Count > 0)
            {
                int stepRow = 0;
                // total row = row start + số row hiện có
                int totalRow = totalRowTable2;
                int rowStart = totalRowTable1 + 26;
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

            ReportDetailtForPartner dataPieYear = null;
            ReportDetailtForPartner dataPieLastYear = null;
            // Data report năm hiện tại nhập vào
            dataPieYear = listDataGradationConvert.Find(x => x.Year == year.ToString());
            // Data report năm ngoái so với năm hiện tại nhập vào
            dataPieLastYear = listDataGradationConvert.Find(x => x.Year == (year - 1).ToString());

            string[] str = { "Chi Quầy", "Chi Nhà", "Chuyển Khoản" };

            if (dataPieYear != null)
            {
                //Add Pie Chart
                Aspose.Cells.Charts.Chart leadSourcePie;
                chartIndex = sheetReport.Charts.Add(ChartType.Pie3D, 44, 0, 60, 5);
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
                string totalRowData = string.Concat("{", string.Format("{0}, {1}, {2}", dataPieYear.DSChiQuay, dataPieYear.DSChiNha, dataPieYear.DSCK), "}");
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
                    //datalabels.ShowValue = true;
                    datalabels.ShowLegendKey = false;
                    datalabels.NumberFormat = "0.00%";
                    datalabels.ShowPercentage = true;
                }
            }

            if (dataPieLastYear != null)
            {
                //Add Pie Chart
                Aspose.Cells.Charts.Chart leadSourcePieLasYear;
                chartIndex = sheetReport.Charts.Add(ChartType.Pie3D, 44, 6, 60, 13);
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
                string totalRowData = string.Concat("{", string.Format("{0}, {1}, {2}", dataPieLastYear.DSChiQuay, dataPieLastYear.DSChiNha, dataPieLastYear.DSCK), "}");
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
            string templatePath = "~/Content/Report/ReportHSDetailt/ReportHSDetailtByPartnerLHForCompare.xlsx";
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
            sheetReport.Cells["D7"].PutValue(string.Format("Tháng {0}/{1} ", month, year));
            sheetReport.Cells["E7"].PutValue(string.Format("Tháng {0}/{1} ", month - 1, year));
            sheetReport.Cells["F7"].PutValue(string.Format("Tháng {0}/{1} ", month, year - 1));

            sheetReport.Cells["F7"].PutValue(string.Format("Tháng {0}/{1} ", month, year));
            sheetReport.Cells["G7"].PutValue(string.Format("Tháng {0}/{1} ", month - 1, year));
            sheetReport.Cells["H7"].PutValue(string.Format("Tháng {0}/{1} ", month, year - 1));

            sheetReport.Cells["I7"].PutValue(string.Format("Tháng {0}/{1} ", month, year));
            sheetReport.Cells["J7"].PutValue(string.Format("Tháng {0}/{1} ", month - 1, year));
            sheetReport.Cells["K7"].PutValue(string.Format("Tháng {0}/{1} ", month, year - 1));

            sheetReport.Cells["L7"].PutValue(string.Format("Tháng {0}/{1} ", month, year));
            sheetReport.Cells["M7"].PutValue(string.Format("Tháng {0}/{1} ", month - 1, year));
            sheetReport.Cells["N7"].PutValue(string.Format("Tháng {0}/{1} ", month, year - 1));

            List<ReportDetailtForPartner> listDataCompareMonth = new HSReportBL().ReportDetailtPartnerCompareMonthForAll(year, month, reportTypeID);

            foreach (ReportDetailtForPartner item in listDataCompareMonth)
            {
                item.TongDS = item.DSChiQuay + item.DSChiNha + item.DSCK;
            }

            DataTable table = new DataTable();
            // Khởi tạo datatable
            table = new DataTable();
            // Tạo các cột cho datatable
            // tháng hiện tại
            table.Columns.Add("STT", typeof(String));

            table.Columns.Add("PartnerName", typeof(String));
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


            List<string> listPartner = new List<string>();
            int count = 1;

            foreach (ReportDetailtForPartner item in listDataCompareMonth)
            {
                // Check tồn tại của đối tác
                if (listPartner.Contains(item.PartnerCode))
                {
                    continue;
                }
                listPartner.Add(item.PartnerCode);

                // Cùng kì
                ReportDetailtForPartner dataItemLastYear = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == month.ToString() && x.Year == (year - 1).ToString());
                ReportDetailtForPartner dataItemYear = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == month.ToString() && x.Year == year.ToString());
                ReportDetailtForPartner dataItemLastMonth = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == (month - 1).ToString() && x.Year == year.ToString());
                // Trường hợp tháng 1
                if (month == 1)
                {
                    dataItemLastMonth = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == "12" && x.Year == (year - 1).ToString());
                }
                // Cung kì
                if (dataItemLastYear == null)
                {
                    dataItemLastYear = new ReportDetailtForPartner()
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
                    dataItemYear = new ReportDetailtForPartner()
                    {
                        PartnerCode = item.PartnerCode,
                        PartnerName = item.PartnerName,
                        Month = month.ToString(),
                        Year = year.ToString()
                    };
                }

                // Tháng trước
                // Trường hợp tháng 1
                if (month == 1)
                {
                    if (dataItemLastMonth == null)
                    {
                        dataItemLastMonth = new ReportDetailtForPartner()
                        {
                            PartnerCode = item.PartnerCode,
                            PartnerName = item.PartnerName,
                            Month = "12",
                            Year = (year - 1).ToString()
                        };
                    }
                }
                else
                {
                    if (dataItemLastMonth == null)
                    {
                        dataItemLastMonth = new ReportDetailtForPartner()
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
                        , dataItemYear.DSChiQuay, dataItemLastMonth.DSChiQuay, dataItemLastYear.DSChiQuay
                        , dataItemYear.DSChiNha, dataItemLastMonth.DSChiNha, dataItemLastYear.DSChiNha
                        , dataItemYear.DSCK, dataItemLastMonth.DSCK, dataItemLastYear.DSCK
                        , dataItemYear.TongDS, dataItemLastMonth.TongDS, dataItemLastYear.TongDS
                    );
                }
            }

            // Add dòng tổng
            DataRow row = table.NewRow();
            row["STT"] = "";
            row["PartnerName"] = "Tổng";
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
            table.Rows.Add(row);


            // Tổng số row theo table1
            int totalRowTable1 = table.Rows.Count + 7;

            // Set border
            Style style = new CellsFactory().CreateStyle();
            style.SetBorder(BorderType.BottomBorder, CellBorderType.Thin, Color.Black);
            style.SetBorder(BorderType.LeftBorder, CellBorderType.Thin, Color.Black);
            style.SetBorder(BorderType.RightBorder, CellBorderType.Thin, Color.Black);

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
                    int totalCol = 1 + 14;
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
                            style.HorizontalAlignment = TextAlignmentType.Center;
                        }
                        else
                        {
                            style.Custom = "#,##0";
                            style.HorizontalAlignment = TextAlignmentType.General;
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

            //// Tạo table cho việc tăng giảm

            listDataCompareMonth = new HSReportBL().ReportDetailtPartnerCompareMonthForAll(year, month, reportTypeID);

            foreach (ReportDetailtForPartner item in listDataCompareMonth)
            {
                item.TongDS = item.DSChiQuay + item.DSChiNha + item.DSCK;
            }

            table = new DataTable();
            // Khởi tạo datatable
            table = new DataTable();
            // Tạo các cột cho datatable
            // tháng hiện tại
            table.Columns.Add("STT", typeof(String));

            table.Columns.Add("PartnerName", typeof(String));
            table.Columns.Add("CQ1", typeof(double));
            table.Columns.Add("CQ2", typeof(double));

            table.Columns.Add("CN1", typeof(double));
            table.Columns.Add("CN2", typeof(double));

            table.Columns.Add("CK1", typeof(double));
            table.Columns.Add("CK2", typeof(double));

            table.Columns.Add("TDS1", typeof(double));
            table.Columns.Add("TDS2", typeof(double));


            listPartner = new List<string>();
            count = 1;

            foreach (ReportDetailtForPartner item in listDataCompareMonth)
            {
                // Check tồn tại của đối tác
                if (listPartner.Contains(item.PartnerCode))
                {
                    continue;
                }
                listPartner.Add(item.PartnerCode);

                // Cùng kì
                ReportDetailtForPartner dataItemLastYear = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == month.ToString() && x.Year == (year - 1).ToString());
                ReportDetailtForPartner dataItemYear = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == month.ToString() && x.Year == year.ToString());
                ReportDetailtForPartner dataItemLastMonth = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == (month - 1).ToString() && x.Year == year.ToString());
                // Trường hợp tháng 1
                if (month == 1)
                {
                    dataItemLastMonth = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == "12" && x.Year == (year - 1).ToString());
                }
                // Cung kì
                if (dataItemLastYear == null)
                {
                    dataItemLastYear = new ReportDetailtForPartner()
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
                    dataItemYear = new ReportDetailtForPartner()
                    {
                        PartnerCode = item.PartnerCode,
                        PartnerName = item.PartnerName,
                        Month = month.ToString(),
                        Year = year.ToString()
                    };
                }

                // Tháng trước
                // Trường hợp tháng 1
                if (month == 1)
                {
                    if (dataItemLastMonth == null)
                    {
                        dataItemLastMonth = new ReportDetailtForPartner()
                        {
                            PartnerCode = item.PartnerCode,
                            PartnerName = item.PartnerName,
                            Month = "12",
                            Year = (year - 1).ToString()
                        };
                    }
                }
                else
                {
                    if (dataItemLastMonth == null)
                    {
                        dataItemLastMonth = new ReportDetailtForPartner()
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
                    // So với tháng trước
                    double sumDSChiQuayYear = dataItemYear.DSChiQuay - dataItemLastMonth.DSChiQuay;
                    double sumDSChiNhaYear = dataItemYear.DSChiNha - dataItemLastMonth.DSChiNha;
                    double sumDSCKYear = dataItemYear.DSCK - dataItemLastMonth.DSCK;
                    double sumTongDSYear = dataItemYear.TongDS - dataItemLastMonth.TongDS;

                    // So với cùng kì năm trước
                    double sumDSChiQuayLastYear = dataItemYear.DSChiQuay - dataItemLastYear.DSChiQuay;
                    double sumDSChiNhaLastYear = dataItemYear.DSChiNha - dataItemLastYear.DSChiNha;
                    double sumDSCKLastYear = dataItemYear.DSCK - dataItemLastYear.DSCK;
                    double sumTongDSLastYear = dataItemYear.TongDS - dataItemLastYear.TongDS;

                    // add item vào table
                    table.Rows.Add(count++, item.PartnerName
                        , sumDSChiQuayYear, sumDSChiQuayLastYear
                        , sumDSChiNhaYear, sumDSChiNhaLastYear
                        , sumDSCKYear, sumDSCKLastYear
                        , sumTongDSYear, sumTongDSLastYear
                    );
                }
            }

            // Add dòng tổng
            row = table.NewRow();
            row["STT"] = "";
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

            // Tổng số row của table2
            // Với 6 là số cách của table1 và table2
            int totalRowTable2 = totalRowTable1 + table.Rows.Count + 6;

            // Tạo title hearder cho table tăng giảm
            // Title cho thị trường

            string title = "Bảng số liệu Tăng giảm so với cùng kì theo từng thị trường";
            CreateTitle(string.Format("B{0}", totalRowTable1 + 3), string.Format("D{0}", totalRowTable1 + 3), sheetReport, title, 12);

            title = "STT";
            CreateTitle(string.Format("B{0}", totalRowTable1 + 6 - 1), string.Format("B{0}", totalRowTable1 + 6), sheetReport, title, 12, true);

            title = "Đối tác";
            CreateTitle(string.Format("C{0}", totalRowTable1 + 6 - 1), string.Format("C{0}", totalRowTable1 + 6), sheetReport, title, 12, true);

            title = "Chi Quầy";
            CreateTitle(string.Format("D{0}", totalRowTable1 + 6 - 1), string.Format("E{0}", totalRowTable1 + 6 - 1), sheetReport, title, 12, true);

            title = "Chi Nhà";
            CreateTitle(string.Format("F{0}", totalRowTable1 + 6 - 1), string.Format("G{0}", totalRowTable1 + 6 - 1), sheetReport, title, 12, true);

            title = "Chuyển Khoản";
            CreateTitle(string.Format("H{0}", totalRowTable1 + 6 - 1), string.Format("I{0}", totalRowTable1 + 6 - 1), sheetReport, title, 12, true);

            title = "Tổng";
            CreateTitle(string.Format("J{0}", totalRowTable1 + 6 - 1), string.Format("K{0}", totalRowTable1 + 6 - 1), sheetReport, title, 12, true);


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
                    int totalCol = 1 + 10;
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
                            style.Custom = "#,##0";
                            style.HorizontalAlignment = TextAlignmentType.General;
                        }

                        // Tô màu cho các dòng có giá trị tăng giảm
                        if (b > 2)
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
            string templatePath = "~/Content/Report/ReportHSDetailt/ReportHSDetailtByPartnerLHCompareForOne.xlsx";
            // Get đường dẫn
            templatePath = System.Web.HttpContext.Current.Server.MapPath(templatePath);

            designer.Workbook = new DongA.Core.DongAExcel().OpenExcelFile(templatePath);
            designer.Workbook.CalculateFormula();

            WorksheetCollection workSheets = designer.Workbook.Worksheets;
            Worksheet sheetReport = designer.Workbook.Worksheets[0];

            // Tạo title
            string typeReport = "So sánh - Theo tháng - Từng thị trường";
            // Tạo title
            CreateTitle("A2", "K2", sheetReport, typeReport, 14);

            // Tạo title detailt
            string titleDetailt = string.Format("Tháng {0}/{1}", month, year);
            CreateTitle("A3", "K3", sheetReport, titleDetailt, 12);

            titleDetailt = "1. Theo doanh số thị trường - loại hình chi trả";
            CreateTitle("B5", "D5", sheetReport, titleDetailt, 12);

            List<ReportDetailtForPartner> listDataCompareMonth = new HSReportBL().ReportDetailtPartnerCompareMonthForOne(year, month, reportTypeID, partnerID);
            List<ReportDetailtForPartner> listDataCompareMonthConvert = new List<ReportDetailtForPartner>();

            foreach (ReportDetailtForPartner item in listDataCompareMonth)
            {
                item.PartnerName = string.Format("Tháng {0}/{1}", item.Month, item.Year);
                item.TongDS = item.DSChiQuay + item.DSChiNha + item.DSCK;
            }

            if (listDataCompareMonth.Count > 0)
            {
                // Month
                ReportDetailtForPartner dataItemMonth = listDataCompareMonth.Find(x => x.Month == month.ToString() && x.Year == year.ToString());
                // Last month
                ReportDetailtForPartner dataItemLastMonth = listDataCompareMonth.Find(x => x.Month == (month - 1).ToString() && x.Year == year.ToString());
                // month == 1
                if (month == 1)
                {
                    dataItemLastMonth = listDataCompareMonth.Find(x => x.Month == "12" && x.Year == (year - 1).ToString());
                }

                // Month last Year

                ReportDetailtForPartner dataItemMonthLastYear = listDataCompareMonth.Find(x => x.Month == month.ToString() && x.Year == (year - 1).ToString());

                // month
                if (dataItemMonth == null)
                {
                    dataItemMonth = new ReportDetailtForPartner()
                    {
                        Month = month.ToString(),
                        Year = year.ToString()
                    };
                }

                // last month
                if (dataItemLastMonth == null)
                {
                    if (month == 1)
                    {
                        dataItemLastMonth = new ReportDetailtForPartner()
                        {
                            Month = "12",
                            Year = (year - 1).ToString()
                        };
                    }
                    else
                    {
                        dataItemLastMonth = new ReportDetailtForPartner()
                        {
                            Month = (month - 1).ToString(),
                            Year = year.ToString()
                        };
                    }

                }

                // month last year
                if (dataItemMonthLastYear == null)
                {
                    dataItemMonthLastYear = new ReportDetailtForPartner()
                    {
                        Month = month.ToString(),
                        Year = (year - 1).ToString()
                    };
                }

                // Tháng hiện tại
                listDataCompareMonthConvert.Add(
                    new ReportDetailtForPartner()
                    {
                        PartnerName = dataItemMonth.PartnerName,
                        DSChiQuay = dataItemMonth.DSChiQuay,
                        DSChiNha = dataItemMonth.DSChiNha,
                        DSCK = dataItemMonth.DSCK,
                        TongDS = dataItemMonth.TongDS,
                        Type = 1
                    }
                );

                // Tháng trước
                listDataCompareMonthConvert.Add(
                    new ReportDetailtForPartner()
                    {
                        PartnerName = dataItemLastMonth.PartnerName,
                        DSChiQuay = dataItemLastMonth.DSChiQuay,
                        DSChiNha = dataItemLastMonth.DSChiNha,
                        DSCK = dataItemLastMonth.DSCK,
                        TongDS = dataItemLastMonth.TongDS,
                        Type = 1
                    }
                );

                // Cùng kì năm trước
                listDataCompareMonthConvert.Add(
                    new ReportDetailtForPartner()
                    {
                        PartnerName = dataItemMonthLastYear.PartnerName,
                        DSChiQuay = dataItemMonthLastYear.DSChiQuay,
                        DSChiNha = dataItemMonthLastYear.DSChiNha,
                        DSCK = dataItemMonthLastYear.DSCK,
                        TongDS = dataItemMonthLastYear.TongDS,
                        Type = 1
                    }
                );

                // Tăng giảm so với tháng trước (+/-)
                listDataCompareMonthConvert.Add(
                    new ReportDetailtForPartner()
                    {
                        PartnerName = "Tăng giảm so với tháng trước (+/-)",
                        DSChiQuay = dataItemMonth.DSChiQuay - dataItemLastMonth.DSChiQuay,
                        DSChiNha = dataItemMonth.DSChiNha - dataItemLastMonth.DSChiNha,
                        DSCK = dataItemMonth.DSCK - dataItemLastMonth.DSCK,
                        TongDS = dataItemMonth.TongDS - dataItemLastMonth.TongDS,
                    }
                );

                // Tăng giảm so với tháng trước (%)
                listDataCompareMonthConvert.Add(
                    new ReportDetailtForPartner()
                    {
                        PartnerName = "Tăng giảm so với tháng trước (%)",
                        DSChiQuay = dataItemLastMonth.DSChiQuay == 0 ? 0 : Math.Round(((dataItemMonth.DSChiQuay - dataItemLastMonth.DSChiQuay) / dataItemLastMonth.DSChiQuay) * 100, 2, MidpointRounding.ToEven),
                        DSChiNha = dataItemLastMonth.DSChiNha == 0 ? 0 : Math.Round(((dataItemMonth.DSChiNha - dataItemLastMonth.DSChiNha) / dataItemLastMonth.DSChiNha) * 100, 2, MidpointRounding.ToEven),
                        DSCK = dataItemLastMonth.DSCK == 0 ? 0 : Math.Round(((dataItemMonth.DSCK - dataItemLastMonth.DSCK) / dataItemLastMonth.DSCK) * 100, 2, MidpointRounding.ToEven),
                        TongDS = dataItemLastMonth.TongDS == 0 ? 0 : Math.Round(((dataItemMonth.TongDS - dataItemLastMonth.TongDS) / dataItemLastMonth.TongDS) * 100, 2, MidpointRounding.ToEven),
                    }
                );

                // Tăng giảm so với cùng kì năm trước (+/-)
                listDataCompareMonthConvert.Add(
                    new ReportDetailtForPartner()
                    {
                        PartnerName = "Tăng giảm so với cùng kì năm trước (+/-)",
                        DSChiQuay = dataItemMonth.DSChiQuay - dataItemMonthLastYear.DSChiQuay,
                        DSChiNha = dataItemMonth.DSChiNha - dataItemMonthLastYear.DSChiNha,
                        DSCK = dataItemMonth.DSCK - dataItemMonthLastYear.DSCK,
                        TongDS = dataItemMonth.TongDS - dataItemMonthLastYear.TongDS,
                    }
                );

                // Tăng giảm so với cùng kì năm trước (%)
                listDataCompareMonthConvert.Add(
                    new ReportDetailtForPartner()
                    {
                        PartnerName = "Tăng giảm so với cùng kì năm trước (%)",
                        DSChiQuay = dataItemMonthLastYear.DSChiQuay == 0 ? 0 : Math.Round(((dataItemMonth.DSChiQuay - dataItemMonthLastYear.DSChiQuay) / dataItemMonthLastYear.DSChiQuay) * 100, 2, MidpointRounding.ToEven),
                        DSChiNha = dataItemMonthLastYear.DSChiNha == 0 ? 0 : Math.Round(((dataItemMonth.DSChiNha - dataItemMonthLastYear.DSChiNha) / dataItemMonthLastYear.DSChiNha) * 100, 2, MidpointRounding.ToEven),
                        DSCK = dataItemMonthLastYear.DSCK == 0 ? 0 : Math.Round(((dataItemMonth.DSCK - dataItemMonthLastYear.DSCK) / dataItemMonthLastYear.DSCK) * 100, 2, MidpointRounding.ToEven),
                        TongDS = dataItemMonthLastYear.TongDS == 0 ? 0 : Math.Round(((dataItemMonth.TongDS - dataItemMonthLastYear.TongDS) / dataItemMonthLastYear.TongDS) * 100, 2, MidpointRounding.ToEven),
                    }
                );
            }

            // Khởi tạo datatable
            DataTable table = new DataTable();
            // Tạo các cột cho datatable
            // tháng hiện tại
            table.Columns.Add("PartnerName", typeof(String));
            table.Columns.Add("CQ1", typeof(double));
            table.Columns.Add("CN1", typeof(double));
            table.Columns.Add("CK1", typeof(double));
            table.Columns.Add("TDS1", typeof(double));


            foreach (ReportDetailtForPartner item in listDataCompareMonthConvert)
            {
                table.Rows.Add(item.PartnerName, item.DSChiQuay, item.DSChiNha, item.DSCK, item.TongDS);
            }


            // Tổng số row theo table1
            int totalRowTable1 = table.Rows.Count + 21;

            // Set border
            Style style = new CellsFactory().CreateStyle();
            style.SetBorder(BorderType.BottomBorder, CellBorderType.Thin, Color.Black);
            style.SetBorder(BorderType.LeftBorder, CellBorderType.Thin, Color.Black);
            style.SetBorder(BorderType.RightBorder, CellBorderType.Thin, Color.Black);
            style.SetBorder(BorderType.TopBorder, CellBorderType.Thin, Color.Black);

            if (table.Rows.Count > 0)
            {
                int stepRow = 0;
                // total row = row start + số row hiện có
                int totalRow = table.Rows.Count + 21;
                // Số dòng của row
                for (int a = 21; a < totalRow; a++)
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
                        if (b > 1 && a > 23)
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
                        if (a > 23)
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
            Aspose.Cells.Charts.Chart leadSourceColumn;
            //Add Pie Chart
            // Chi Quầy
            int chartIndex = sheetReport.Charts.Add(ChartType.Column3DClustered, 7, 1, 19, 6);
            leadSourceColumn = sheetReport.Charts[chartIndex];

            //Chart title
            leadSourceColumn.Title.Text = string.Format("Doanh số từng dịch vụ từng thị trường \n Tháng {0}/{1}", month, year);
            leadSourceColumn.Title.Font.Color = Color.Silver;


            List<ReportDetailtForPartner> listDataCompareMonthClone = listDataCompareMonthConvert.Where(x => x.Type == 1).ToList();

            string[] totalRowData = new string[listDataCompareMonthClone.Count];
            int count = 0;

            foreach (ReportDetailtForPartner item in listDataCompareMonthClone)
            {

                totalRowData[count++] = string.Concat("{", item.DSChiQuay, ",", item.DSChiNha, ",", item.DSCK, "}");
            }

            string categoryData = string.Empty;

            foreach (string item in totalRowData)
            {
                leadSourceColumn.NSeries.Add(item, true);

                categoryData = "{DS Chi Quầy, DS Chi Nhà, DS Chuyển Khoản}";
                leadSourceColumn.NSeries.CategoryData = categoryData;
            }
            
            // Set the names of the chart series taken from cells.
            leadSourceColumn.NSeries[0].Name = "=B22";
            leadSourceColumn.NSeries[1].Name = "=B23";
            leadSourceColumn.NSeries[2].Name = "=B24";

            // Set the 1st series fill color.
            leadSourceColumn.NSeries[0].Area.ForegroundColor = Color.Orange;
            leadSourceColumn.NSeries[0].Area.Formatting = FormattingType.Custom;

            // Set the 2nd series fill color.
            leadSourceColumn.NSeries[1].Area.ForegroundColor = Color.Green;
            leadSourceColumn.NSeries[1].Area.Formatting = FormattingType.Custom;

            // Set the 2nd series fill color.
            leadSourceColumn.NSeries[2].Area.ForegroundColor = Color.Blue;
            leadSourceColumn.NSeries[2].Area.Formatting = FormattingType.Custom;


            // Set plot area formatting as none and hide its border.
            leadSourceColumn.PlotArea.Area.FillFormat.FillType = FillType.None;
            leadSourceColumn.PlotArea.Border.IsVisible = false;

            // Set value axis major tick mark as none and hide axis line. 
            // Also set the color of value axis major grid lines.
            leadSourceColumn.ValueAxis.MajorTickMark = TickMarkType.None;
            leadSourceColumn.ValueAxis.AxisLine.IsVisible = false;
            //leadSourceColumnChiNha.ValueAxis.IsAutomaticMajorUnit = false;
            //leadSourceColumnChiNha.ValueAxis.MajorUnit = 10000000;
            leadSourceColumn.ValueAxis.MajorGridLines.Color = Color.FromArgb(217, 217, 217);

            // Tổng số row của table2
            // Với 6 là số cách của table1 và table2
            // Tạo title hearder cho table tăng giảm
            // Title cho thị trường
            string title = "2. Theo tỷ trọng doanh số đối tác - loại hình dịch vụ";

            CreateTitle(string.Format("B{0}", totalRowTable1 + 6 - 4), string.Format("D{0}", totalRowTable1 + 6 - 4), sheetReport, title, 12);

            title = "";
            CreateTitle(string.Format("B{0}", totalRowTable1 + 26 - 1), string.Format("B{0}", totalRowTable1 + 26 - 1), sheetReport, title, 12, true);

            title = "Chi Quầy";
            CreateTitle(string.Format("C{0}", totalRowTable1 + 26 - 1), string.Format("C{0}", totalRowTable1 + 26 - 1), sheetReport, title, 12, true);

            title = "Chi Nhà";
            CreateTitle(string.Format("D{0}", totalRowTable1 + 26 - 1), string.Format("D{0}", totalRowTable1 + 26 - 1), sheetReport, title, 12, true);

            title = "Chuyển Khoản";
            CreateTitle(string.Format("E{0}", totalRowTable1 + 26 - 1), string.Format("E{0}", totalRowTable1 + 26 - 1), sheetReport, title, 12, true);

            title = "Tổng";
            CreateTitle(string.Format("F{0}", totalRowTable1 + 26 - 1), string.Format("F{0}", totalRowTable1 + 26 - 1), sheetReport, title, 12, true);

            listDataCompareMonth = new HSReportBL().ReportDetailtPartnerCompareMonthForOne(year, month, reportTypeID, partnerID);
            listDataCompareMonthConvert = new List<ReportDetailtForPartner>();

            foreach (ReportDetailtForPartner item in listDataCompareMonth)
            {
                item.PartnerName = string.Format("Tháng {0}/{1}", item.Month, item.Year);
                item.TongDS = item.DSChiQuay + item.DSChiNha + item.DSCK;

                listDataCompareMonthConvert.Add(
                    new ReportDetailtForPartner()
                    {
                        PartnerName = string.Format("Tháng {0}/{1}", item.Month, item.Year),
                        DSChiQuay = item.TongDS == 0 ? 0 : Math.Round((item.DSChiQuay / item.TongDS) * 100, 2, MidpointRounding.ToEven),
                        DSChiNha = item.TongDS == 0 ? 0 : Math.Round((item.DSChiNha / item.TongDS) * 100, 2, MidpointRounding.ToEven),
                        DSCK = item.TongDS == 0 ? 0 : Math.Round((item.DSCK / item.TongDS) * 100, 2, MidpointRounding.ToEven),
                        TongDS = 100,
                        Type = 1,
                        Month = item.Month,
                        Year = item.Year
                    }
                );
            }


            if (listDataCompareMonthConvert.Count().Equals(3))
            {
                // so với tháng trước
                double sumDSChiQuayYear = listDataCompareMonthConvert[2].DSChiQuay - listDataCompareMonthConvert[1].DSChiQuay;
                double sumDSChiNhaYear = listDataCompareMonthConvert[2].DSChiNha - listDataCompareMonthConvert[1].DSChiNha;
                double sumDSCKYear = listDataCompareMonthConvert[2].DSCK - listDataCompareMonthConvert[1].DSCK;
                double sumTongDSYear = listDataCompareMonthConvert[2].TongDS - listDataCompareMonthConvert[1].TongDS;

                // so với cùng kì
                double sumDSChiQuayLastYear = listDataCompareMonthConvert[2].DSChiQuay - listDataCompareMonthConvert[0].DSChiQuay;
                double sumDSChiNhaLastYear = listDataCompareMonthConvert[2].DSChiNha - listDataCompareMonthConvert[0].DSChiNha;
                double sumDSCKLastYear = listDataCompareMonthConvert[2].DSCK - listDataCompareMonthConvert[0].DSCK;
                double sumTongDSLastYear = listDataCompareMonthConvert[2].TongDS - listDataCompareMonthConvert[0].TongDS;

                // Tăng giảm so với tháng trước (%)
                listDataCompareMonthConvert.Add(
                    new ReportDetailtForPartner()
                    {
                        PartnerName = "Tăng giảm so với tháng trước (%)",
                        DSChiQuay = Math.Round(sumDSChiQuayYear, 2, MidpointRounding.ToEven),
                        DSChiNha = Math.Round(sumDSChiNhaYear, 2, MidpointRounding.ToEven),
                        DSCK = Math.Round(sumDSCKYear, 2, MidpointRounding.ToEven),
                        TongDS = 0
                    }
                );

                // Tăng giảm so với cùng kì năm trước (%)
                listDataCompareMonthConvert.Add(
                    new ReportDetailtForPartner()
                    {
                        PartnerName = "Tăng giảm so với cùng kì năm trước (%)",
                        DSChiQuay = Math.Round(sumDSChiQuayLastYear, 2, MidpointRounding.ToEven),
                        DSChiNha = Math.Round(sumDSChiNhaLastYear, 2, MidpointRounding.ToEven),
                        DSCK = Math.Round(sumDSCKLastYear, 2, MidpointRounding.ToEven),
                        TongDS = 0
                    }
                );
            }

            table.Rows.Clear();

            foreach (ReportDetailtForPartner item in listDataCompareMonthConvert)
            {
                table.Rows.Add(item.PartnerName, item.DSChiQuay, item.DSChiNha, item.DSCK, item.TongDS);
            }


            int totalRowTable2 = totalRowTable1 + table.Rows.Count + 25;

            // Table dữ liệu bảng số liệu Doanh số Chi Quầy/Chi Nhà/Chuyển Khoản
            if (table.Rows.Count > 0)
            {
                int stepRow = 0;
                // total row = row start + số row hiện có
                int totalRow = totalRowTable2;
                int rowStart = totalRowTable1 + 26 - 1;
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

                        // Insert vào dòng cột xác định trong Excel
                        sheetReport.Cells[a, b].PutValue(valueOfTable, true);

                        // set style cho number
                        style.Custom = "#,##0.00";

                        // Tô màu cho các dòng có giá trị tăng giảm
                        if (b > 1 && a > totalRow - 3)
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

                        if (a.Equals(totalRow - 2))
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
            // Month
            ReportDetailtForPartner dataPieMonth = listDataCompareMonthConvert.Find(x => x.Month == month.ToString() && x.Year == year.ToString() && x.Type == 1);

            // last month
            ReportDetailtForPartner dataPieLastMonth = listDataCompareMonthConvert.Find(x => x.Month == (month - 1).ToString() && x.Year == year.ToString() && x.Type == 1);
            if (month == 1)
            {
                dataPieLastMonth = listDataCompareMonthConvert.Find(x => x.Month == "12" && x.Year == (year - 1).ToString() && x.Type == 1);
            }

            // Last Year
            ReportDetailtForPartner dataPieMonthLastYear = listDataCompareMonth.Find(x => x.Year == (year - 1).ToString());

            if (dataPieMonth != null)
            {
                //Add Pie Chart
                Aspose.Cells.Charts.Chart leadSourcePie;
                chartIndex = sheetReport.Charts.Add(ChartType.Pie3D, 32, 1, 50, 4);
                leadSourcePie = sheetReport.Charts[chartIndex];

                // Set some properties of chart plot area.
                // To set the fill color and make the border invisible.
                leadSourcePie.PlotArea.Border.IsVisible = false;
                leadSourcePie.Elevation = 45;
                // Set properties of chart title
                leadSourcePie.Title.Text = string.Format("Tỉ trọng các đối tác theo thị trường Tháng {0}/{1}", month, year);
                leadSourcePie.Title.Font.Color = Color.Silver;
                leadSourcePie.Title.Font.IsBold = true;
                leadSourcePie.Title.Font.Size = 12;

                string totalRowDataPie = string.Concat("{", string.Format("{0}, {1}, {2}", dataPieMonth.DSChiQuay, dataPieMonth.DSChiNha, dataPieMonth.DSCK), "}");
                leadSourcePie.NSeries.Add(totalRowDataPie, true);

                categoryData = "{DS Chi Quầy, DS Chi Nhà, DS Chuyển Khoản}";
                leadSourcePie.NSeries.CategoryData = categoryData;

                leadSourcePie.NSeries.IsColorVaried = true;

                // Set the DataLabels in the chart
                Aspose.Cells.Charts.DataLabels datalabels;
                for (int i = 0; i < leadSourcePie.NSeries.Count; i++)
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

            if (dataPieLastMonth != null)
            {
                //Add Pie Chart
                Aspose.Cells.Charts.Chart leadSourcePie;
                chartIndex = sheetReport.Charts.Add(ChartType.Pie3D, 32, 5, 50, 9);
                leadSourcePie = sheetReport.Charts[chartIndex];

                // Set some properties of chart plot area.
                // To set the fill color and make the border invisible.
                leadSourcePie.PlotArea.Border.IsVisible = false;
                leadSourcePie.Elevation = 45;
                // Set properties of chart title
                if (month == 1)
                {
                    leadSourcePie.Title.Text = string.Format("Tỉ trọng các đối tác theo thị trường Tháng {0}/{1}", 12, year - 1);
                }
                else
                {

                    leadSourcePie.Title.Text = string.Format("Tỉ trọng các đối tác theo thị trường Tháng {0}/{1}", month - 1, year);
                }
                leadSourcePie.Title.Font.Color = Color.Silver;
                leadSourcePie.Title.Font.IsBold = true;
                leadSourcePie.Title.Font.Size = 12;

                string totalRowDataPie = string.Concat("{", string.Format("{0}, {1}, {2}", dataPieLastMonth.DSChiQuay, dataPieLastMonth.DSChiNha, dataPieLastMonth.DSCK), "}");
                leadSourcePie.NSeries.Add(totalRowDataPie, true);

                categoryData = "{DS Chi Quầy, DS Chi Nhà, DS Chuyển Khoản}";
                leadSourcePie.NSeries.CategoryData = categoryData;

                leadSourcePie.NSeries.IsColorVaried = true;

                // Set the DataLabels in the chart
                Aspose.Cells.Charts.DataLabels datalabels;
                for (int i = 0; i < leadSourcePie.NSeries.Count; i++)
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

            if (dataPieMonthLastYear != null)
            {
                //Add Pie Chart
                Aspose.Cells.Charts.Chart leadSourcePie;
                chartIndex = sheetReport.Charts.Add(ChartType.Pie3D, 32, 10, 50, 14);
                leadSourcePie = sheetReport.Charts[chartIndex];

                // Set some properties of chart plot area.
                // To set the fill color and make the border invisible.
                leadSourcePie.PlotArea.Border.IsVisible = false;
                leadSourcePie.Elevation = 45;
                // Set properties of chart title
                leadSourcePie.Title.Text = string.Format("Tỉ trọng các đối tác theo thị trường Tháng {0}/{1}", month, year - 1);
                leadSourcePie.Title.Font.Color = Color.Silver;
                leadSourcePie.Title.Font.IsBold = true;
                leadSourcePie.Title.Font.Size = 12;

                string totalRowDataPie = string.Concat("{", string.Format("{0}, {1}, {2}", dataPieMonthLastYear.DSChiQuay, dataPieMonthLastYear.DSChiNha, dataPieMonthLastYear.DSCK), "}");
                leadSourcePie.NSeries.Add(totalRowDataPie, true);

                categoryData = "{DS Chi Quầy, DS Chi Nhà, DS Chuyển Khoản}";
                leadSourcePie.NSeries.CategoryData = categoryData;

                leadSourcePie.NSeries.IsColorVaried = true;

                // Set the DataLabels in the chart
                Aspose.Cells.Charts.DataLabels datalabels;
                for (int i = 0; i < leadSourcePie.NSeries.Count; i++)
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
        private DataTable CreateDataTableFormart(bool type = true)
        {
            DataTable db = new DataTable();

            db.Columns.Add("STT", typeof(string));
            if(type)
            {
                db.Columns.Add("ReportID", typeof(string));
            }
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
        private void FillData(DataTable mother, DataTable fill, bool type = true)
        {
            int stt = 1;
            foreach (DataRow dr in mother.Rows)
            {
                var row = fill.NewRow();

                row["STT"] = string.IsNullOrEmpty(dr["STT"].ToString()) ? "" : (string)dr["STT"];
                if (type)
                {
                    row["ReportID"] = string.IsNullOrEmpty(dr["ReportID"].ToString()) ? "" : (string)dr["ReportID"];
                }
                
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
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
    public class ReportHSTotalHSExcelController : Controller
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

        // GET: Admin/ReportHSTotalHSExcel
        public ActionResult Index()
        {
            return View();
        }


        /// <summary>
        /// Xuất Excel theo tháng
        /// </summary>
        /// <param name="fromDate"></param>
        /// <param name="toDate"></param>
        /// <returns></returns>
        public ActionResult CreateExcelForDayMonthYear(DateTime fromDate, DateTime toDate, int typeID, string reportTypeID)
        {
            WorkbookDesigner designer = new WorkbookDesigner();
            string templatePath = "~/Content/Report/ReportHS/ReportHSTotalHS.xlsx";
            // Get đường dẫn
            templatePath = System.Web.HttpContext.Current.Server.MapPath(templatePath);

            designer.Workbook = new DongA.Core.DongAExcel().OpenExcelFile(templatePath);
            designer.Workbook.CalculateFormula();

            WorksheetCollection workSheets = designer.Workbook.Worksheets;
            Worksheet sheetReport = designer.Workbook.Worksheets[0];

            // Tạo title
            string typeReport = string.Empty;
            switch (typeID)
            {
                // Theo ngày
                case 1:
                    typeReport = "Chi tiết - Theo Ngày";
                    // Set from day and to day
                    sheetReport.Cells["M4"].PutValue(string.Format("{0}", fromDate.ToString("dd/MM/yyyy")));
                    sheetReport.Cells["Q4"].PutValue(string.Format("{0}", toDate.ToString("dd/MM/yyyy")));

                    break;
                // Theo tháng
                case 2:
                    typeReport = "Chi tiết - Theo Tháng";

                    // Set from day and to day
                    sheetReport.Cells["M4"].PutValue(string.Format("{0}/{1}", fromDate.Month, fromDate.Year));
                    sheetReport.Cells["Q4"].PutValue(string.Format("{0}/{1}", toDate.Month, toDate.Year));
                    break;
                // Theo năm
                default:
                    typeReport = "Chi tiết - Theo Năm";
                    // Set from day and to day
                    sheetReport.Cells["M4"].PutValue(fromDate.Year.ToString());
                    sheetReport.Cells["Q4"].PutValue(toDate.Year.ToString());
                    break;
            }

            CreateTitle("A2", "U2", sheetReport, typeReport, 14);

            // Create Chart Line
            //Chart reference
            Aspose.Cells.Charts.Chart leadSourceLine;
            //Add Pie Chart
            int chartIndex = sheetReport.Charts.Add(ChartType.Line, 6, 0, 30, 13);
            leadSourceLine = sheetReport.Charts[chartIndex];

            // Canh hiển thị CategoryAxis nghiên phù hợp
            leadSourceLine.CategoryAxis.TickLabels.RotationAngle = 45;

            //Chart title
            leadSourceLine.Title.Text = "Tổng doanh số chi trả";
            leadSourceLine.Title.Font.Color = Color.Silver;

            // Get data report ngày
            List<ReportForTotalPayment> listReportData = new List<ReportForTotalPayment>();

            switch (typeID)
            {
                // Theo ngày
                case 1:
                    listReportData = new HSReportBL().SearchReportTPForDay(fromDate, toDate, reportTypeID);
                    break;
                // Theo tháng
                case 2:
                    listReportData = new HSReportBL().SearchReportTPForMonth(fromDate, toDate, reportTypeID);
                    break;
                // Theo năm
                default:
                    listReportData = new HSReportBL().SearchReportTPForYear(fromDate, toDate, reportTypeID);
                    break;
            }

            DataTable dataTable = new DataTable();

            // Set border
            Style style = new CellsFactory().CreateStyle();
            style.SetBorder(BorderType.TopBorder, CellBorderType.Thin, Color.Black);
            style.SetBorder(BorderType.BottomBorder, CellBorderType.Thin, Color.Black);
            style.SetBorder(BorderType.LeftBorder, CellBorderType.Thin, Color.Black);
            style.SetBorder(BorderType.RightBorder, CellBorderType.Thin, Color.Black);

            if (listReportData.Count > 0)
            {
                foreach (ReportForTotalPayment item in listReportData)
                {
                    switch (typeID)
                    {
                        // Theo ngày
                        case 1:
                            item.ReportID = string.Concat("Ngày ", item.CreatedDate.ToString("dd/MM/yyyy"));
                            break;
                        // Theo tháng
                        case 2:
                            item.ReportID = string.Concat("Tháng ", item.Month, "/", item.Year);
                            break;
                        // Theo năm
                        default:
                            item.ReportID = string.Concat("Năm ", item.Year);
                            break;
                    }
                    item.Type = 0;
                }


                // Add dòng tổng vào list danh sách
                ReportForTotalPayment reportForTotalPayment = new ReportForTotalPayment()
                {
                    ReportID = "Tổng",
                    Payed = listReportData.Sum(item => item.Payed)
                };
                listReportData.Add(reportForTotalPayment);

                // Tạo các col cho table
                dataTable = CreateDataTableFormart();

                // Danh sách dataSet của báo cáo ngày
                DataSet dataReport = ConvertListObjectToDataSet(listReportData);

                // Đổ data vào datatble mới
                FillData(dataReport.Tables[0], dataTable);

                // Adding SeriesCollection (chart data source) to the chart ranging from "A1" cell to "B3"
                string totalRowData = string.Format("Q8:Q{0}", listReportData.Count + 8 - 2);
                leadSourceLine.NSeries.Add(totalRowData, true);

                // Set the category data covering the range A2:A5.
                // Tổng số dòng cần hiển thị là số dòng hiện tại  + số dòng trong listData -2 (trong đó 1 là dòng tổng cuối cùng)
                string categoryData = string.Format("P8:P{0}", listReportData.Count + 8 - 2);
                leadSourceLine.NSeries.CategoryData = categoryData;

                // Set the names of the chart series taken from cells.
                leadSourceLine.NSeries[0].Name = "=Q7";

                // Set the 1st series fill color.
                //leadSourceLine.NSeries[0].Area.ForegroundColor = Color.Pink;
                leadSourceLine.NSeries[0].Border.Color = Color.Blue;
                leadSourceLine.NSeries[0].Area.Formatting = FormattingType.Custom;


                // Set plot area formatting as none and hide its border.
                leadSourceLine.PlotArea.Area.FillFormat.FillType = FillType.None;
                leadSourceLine.PlotArea.Border.IsVisible = false;

                // Set value axis major tick mark as none and hide axis line. 
                // Also set the color of value axis major grid lines.
                leadSourceLine.ValueAxis.MajorTickMark = TickMarkType.None;
                leadSourceLine.ValueAxis.AxisLine.IsVisible = false;
                leadSourceLine.ValueAxis.MajorGridLines.Color = Color.FromArgb(217, 217, 217);
            }

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
                    // trường hợp với report chọn tháng
                    int totalCol = 15 + 2;

                    for (int b = 15; b < totalCol; b++)
                    {
                        // Giá trị của value trong table
                        string valueOfTable = dataTable.Rows[stepRow][stepColumn].ToString();

                        // Insert vào dòng cột xác định trong Excel
                        sheetReport.Cells[a, b].PutValue(valueOfTable, true);

                        // set style cho number
                        style.Custom = "#,##";

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
        /// Tạo mẫu cho Excel cho so sánh theo giai đoạn
        /// </summary>
        /// <param name="gradationID"></param>
        /// <param name="year"></param>
        /// <param name="typeID"></param>
        /// <returns></returns>
        public ActionResult CreateExcelForGradationCompare(string gradation, int year, string reportTypeID)
        {
            WorkbookDesigner designer = new WorkbookDesigner();
            string templatePath = "~/Content/Report/ReportHS/ReportHSTotalHS.xlsx";
            // Get đường dẫn
            templatePath = System.Web.HttpContext.Current.Server.MapPath(templatePath);

            designer.Workbook = new DongA.Core.DongAExcel().OpenExcelFile(templatePath);
            designer.Workbook.CalculateFormula();

            WorksheetCollection workSheets = designer.Workbook.Worksheets;
            Worksheet sheetReport = designer.Workbook.Worksheets[0];

            // Tạo title
            string typeReport = "So sánh - Theo giai đoạn";

            string text = " tháng đầu năm";
            switch (gradation)
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

            // xóa từ/đến
            sheetReport.Cells["L4"].PutValue("");
            sheetReport.Cells["P4"].PutValue("");


            // Set width cho column
            sheetReport.Cells.SetColumnWidthPixel(15, 250);

            List<ReportForTotalPayment> listReportData = new HSReportBL().SearchReportTotalHSForGradationCompare(year, int.Parse(gradation), reportTypeID);

            // clone Object
            List<ReportForTotalPayment> listReportDataClone = new List<ReportForTotalPayment>(listReportData);

            DataTable dataTable = new DataTable();

            // Theo doanh số chi trả loại hình dịch vụ
            if (listReportData.Count.Equals(2))
            {
                bool check = true;
                foreach (ReportForTotalPayment item in listReportData)
                {
                    item.ReportID = string.Concat("Lũy kế ", text, " ", year);
                    if (!check)
                    {
                        item.ReportID = string.Concat("Lũy kế ", text, " ", year - 1);
                    }
                    item.Type = 0;
                    // Set lại giá trị cho check để lấy giá trị của năm trước
                    check = false;
                }

                // Object báo cáo tăng giảm so với cùng kỳ (+/-)
                // ds so với tháng trước
                double tongDSSum = (listReportData[0].Payed - listReportData[1].Payed);

                // Object báo cáo tăng giảm so với cùng kỳ (%)
                ReportForTotalPayment dataDifferencePercent = new ReportForTotalPayment()
                {
                    ReportID = string.Format("Tăng giảm so với cùng kì {0} (%)", year - 1),
                    Payed = Math.Round(tongDSSum / listReportData[1].Payed * 100, 2, MidpointRounding.ToEven)
                };

                listReportData.Add(dataDifferencePercent);

                ReportForTotalPayment dataDifference = new ReportForTotalPayment()
                {
                    ReportID = string.Format("Tăng giảm so với cùng kì {0} (+/-)", year - 1),
                    Payed = Math.Round(tongDSSum, 2, MidpointRounding.ToEven)
                };

                listReportData.Add(dataDifference);

                // Add item to datatable
                dataTable = CreateDataTableFormart();

                // Danh sách dataSet của báo cáo ngày
                DataSet dataReport = new DongA.Core.DongAExcel().ConvertListObjectToDataSet(listReportData);

                // Đổ data vào datatble mới
                FillData(dataReport.Tables[0], dataTable);


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
                        int totalCol = 15 + 2;
                        for (int b = 15; b < totalCol; b++)
                        {
                            // Giá trị của value trong table
                            string valueOfTable = dataTable.Rows[stepRow][stepColumn].ToString();


                            // Tô màu cho các dòng có giá trị tăng giảm
                            if (b >= 16 && a >= 9)
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

                            if (a.Equals(totalRow - 2))
                            {
                                // set style cho number
                                style.Custom = "#,##0.00";
                            }
                            else
                            {
                                // set style cho number
                                style.Custom = "#,##";
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
                            if (a.Equals(totalRow - 1) || a.Equals(totalRow - 2))
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


                // Create Chart Line
                //Chart reference
                Aspose.Cells.Charts.Chart leadSourceLine;
                //Add Pie Chart
                int chartIndex = sheetReport.Charts.Add(ChartType.Column3DClustered, 6, 0, 30, 13);
                leadSourceLine = sheetReport.Charts[chartIndex];

                //// Canh hiển thị CategoryAxis nghiên phù hợp
                //leadSourceLine.CategoryAxis.TickLabels.RotationAngle = 45;

                //Chart title
                leadSourceLine.Title.Text = "Tổng doanh số chi trả";
                leadSourceLine.Title.Font.Color = Color.Silver;


                // get data
                string totalRowData = "Q8:Q9";
                leadSourceLine.NSeries.Add(totalRowData, true);

                // Get row
                string categoryData = "P8:P9";
                leadSourceLine.NSeries.CategoryData = categoryData;
                // Set the names of the chart series taken from cells.
                leadSourceLine.NSeries[0].Name = "=Q7";

                //// Set the 1st series fill color.
                //leadSourceLine.NSeries[0].Border.Color = Color.Orange;
                //leadSourceLine.NSeries[0].Area.Formatting = FormattingType.Custom;

                //// Set the 2nd series fill color.
                //leadSourceLine.NSeries[1].Border.Color = Color.Green;
                //leadSourceLine.NSeries[1].Area.Formatting = FormattingType.Custom;

                // Set plot area formatting as none and hide its border.
                leadSourceLine.PlotArea.Area.FillFormat.FillType = FillType.None;
                leadSourceLine.PlotArea.Border.IsVisible = false;

                // Set value axis major tick mark as none and hide axis line. 
                // Also set the color of value axis major grid lines.
                leadSourceLine.ValueAxis.MajorTickMark = TickMarkType.None;
                leadSourceLine.ValueAxis.AxisLine.IsVisible = false;
                leadSourceLine.ValueAxis.MinValue = 0;
                leadSourceLine.ValueAxis.MajorGridLines.Color = Color.FromArgb(217, 217, 217);
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
        public ActionResult CreateExcelGradationCompareLastYear(int year, int month, string reportTypeID)
        {
            WorkbookDesigner designer = new WorkbookDesigner();
            string templatePath = "~/Content/Report/ReportHS/ReportHSTotalHS.xlsx";
            // Get đường dẫn
            templatePath = System.Web.HttpContext.Current.Server.MapPath(templatePath);

            designer.Workbook = new DongA.Core.DongAExcel().OpenExcelFile(templatePath);
            designer.Workbook.CalculateFormula();

            WorksheetCollection workSheets = designer.Workbook.Worksheets;
            Worksheet sheetReport = designer.Workbook.Worksheets[0];

            // Tạo title
            string typeReport = "So sánh - Theo tháng - So với tháng trước";

            string text = string.Format("Tháng: {0}/{1}", month, year);

            // Tạo title
            CreateTitle("A2", "U2", sheetReport, typeReport, 14);

            // Tạo title detailt
            string titleDetailt = text;
            CreateTitle("A3", "U3", sheetReport, titleDetailt, 12);

            // xóa từ/đến
            sheetReport.Cells["L4"].PutValue("");
            sheetReport.Cells["P4"].PutValue("");


            // Set width cho column
            sheetReport.Cells.SetColumnWidthPixel(15, 220);

            List<ReportForTotalPayment> listReportData = new HSReportBL().SearchReportTotalHSForCompareMonth(year, month, reportTypeID);

            // clone Object
            List<ReportForTotalPayment> listReportDataClone = new List<ReportForTotalPayment>(listReportData);

            DataTable dataTable = new DataTable();

            // Theo doanh số chi trả loại hình dịch vụ
            if (listReportData.Count.Equals(3))
            {
                foreach (ReportForTotalPayment item in listReportData)
                {
                    item.ReportID = string.Format("Tháng {0}/{1}", item.Month, item.Year);
                }
                double totalPayment = listReportData[0].Payed - listReportData[1].Payed;
                double totalPaymentLastYear = listReportData[0].Payed - listReportData[2].Payed;

                // Object báo cáo tăng giảm so với tháng trước (%)
                ReportForTotalPayment dataDifference = null;
                dataDifference = new ReportForTotalPayment()
                {
                    ReportID = "Tăng giảm so với tháng trước (%)",
                    Payed = Math.Round(totalPayment / listReportData[1].Payed * 100, 2, MidpointRounding.ToEven),
                };

                listReportData.Add(dataDifference);

                // Object báo cáo tăng giảm so với tháng trước (+/-)
                dataDifference = new ReportForTotalPayment()
                {
                    ReportID = "Tăng giảm so với tháng trước (+/-)",
                    Payed = Math.Round(totalPayment, 2, MidpointRounding.ToEven),
                };

                listReportData.Add(dataDifference);

                // Object báo cáo tăng giảm so với cùng kì năm trước (%)
                dataDifference = new ReportForTotalPayment()
                {
                    ReportID = "Tăng giảm so với cùng kì năm trước (%)",
                    Payed = Math.Round(totalPaymentLastYear / listReportData[2].Payed * 100, 2, MidpointRounding.ToEven),
                };

                listReportData.Add(dataDifference);

                // Object báo cáo tăng giảm so với cùng kì năm trước (+/-)
                dataDifference = new ReportForTotalPayment()
                {
                    ReportID = "Tăng giảm so với cùng kì năm trước (+/-)",
                    Payed = Math.Round(totalPaymentLastYear, 2, MidpointRounding.ToEven),
                };

                listReportData.Add(dataDifference);

                // Add item to datatable
                dataTable = CreateDataTableFormart();

                // Danh sách dataSet của báo cáo ngày
                DataSet dataReport = new DongA.Core.DongAExcel().ConvertListObjectToDataSet(listReportData);

                // Đổ data vào datatble mới
                FillData(dataReport.Tables[0], dataTable);


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
                        int totalCol = 15 + 2;
                        for (int b = 15; b < totalCol; b++)
                        {
                            // Giá trị của value trong table
                            string valueOfTable = dataTable.Rows[stepRow][stepColumn].ToString();

                            // Tô màu cho các dòng có giá trị tăng giảm
                            if (b >= 16 && a >= 10)
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
                            if (a.Equals(totalRow - 2) || a.Equals(totalRow - 4))
                            {
                                // set style cho number
                                style.Custom = "#,##0.00";
                            }
                            else
                            {
                                // set style cho number
                                style.Custom = "#,##";
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
                            if (a.Equals(totalRow - 1) || a.Equals(totalRow - 2))
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


            // Create Chart Line
            //Chart reference
            Aspose.Cells.Charts.Chart leadSourceLine;
            //Add Pie Chart
            int chartIndex = sheetReport.Charts.Add(ChartType.Column3DClustered, 6, 0, 30, 13);
            leadSourceLine = sheetReport.Charts[chartIndex];

            //Chart title
            leadSourceLine.Title.Text = string.Format("Tổng doanh số chi trả {0}/{1} so với tháng trước và so với cùng kì năm trước", month, year);
            leadSourceLine.Title.Font.Color = Color.Silver;

            // get data
            string totalRowData = "Q8:Q10";
            leadSourceLine.NSeries.Add(totalRowData, true);

            // Get row
            string categoryData = "P8:P10";
            leadSourceLine.NSeries.CategoryData = categoryData;
            // Set the names of the chart series taken from cells.
            leadSourceLine.NSeries[0].Name = "=Q7";

            // Set plot area formatting as none and hide its border.
            leadSourceLine.PlotArea.Area.FillFormat.FillType = FillType.None;
            leadSourceLine.PlotArea.Border.IsVisible = false;

            // Set value axis major tick mark as none and hide axis line. 
            // Also set the color of value axis major grid lines.
            leadSourceLine.ValueAxis.MajorTickMark = TickMarkType.None;
            leadSourceLine.ValueAxis.AxisLine.IsVisible = false;
            leadSourceLine.ValueAxis.MajorGridLines.Color = Color.FromArgb(217, 217, 217);

            //// Canh hiển thị CategoryAxis nghiên phù hợp
            //leadSourceLine.CategoryAxis.TickLabels.RotationAngle = 45;

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

            db.Columns.Add("ReportID", typeof(string));
            db.Columns.Add("Payed", typeof(double));
            db.Columns.Add("TongDS", typeof(double));
            return db;
        }

        /// <summary>
        /// Tạo cột cho datatable
        /// </summary>
        /// <returns></returns>
        private DataTable CreateDataTableFormartForYear()
        {
            DataTable db = new DataTable();

            db.Columns.Add("ReportID", typeof(string));
            db.Columns.Add("Year1", typeof(double));
            db.Columns.Add("Year2", typeof(double));
            db.Columns.Add("Year3", typeof(double));
            db.Columns.Add("Year4", typeof(double));
            db.Columns.Add("Year5", typeof(double));
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

                row["ReportID"] = (string)dr["ReportID"];
                row["Payed"] = (double)dr["Payed"];
                fill.Rows.Add(row);
                stt++;
            }
        }

        /// <summary>
        /// Add row cho datatable
        /// </summary>
        /// <param name="mother"></param>
        /// <param name="fill"></param>
        private void FillDataForYear(DataTable mother, DataTable fill)
        {
            int stt = 1;
            foreach (DataRow dr in mother.Rows)
            {
                var row = fill.NewRow();

                row["ReportID"] = (string)dr["ReportID"];
                row["Year1"] = (double)dr["Year1"];
                row["Year2"] = (double)dr["Year2"];
                row["Year3"] = (double)dr["Year3"];
                row["Year4"] = (double)dr["Year4"];
                row["Year5"] = (double)dr["Year5"];
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
            PropertyDescriptorCollection properties =
               TypeDescriptor.GetProperties(typeof(T));
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
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
    public class ReportExcelForTotalMoneyTypeController : Controller
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
        // GET: Admin/ReportExcelForTotalMoneyType
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
            string templatePath = "~/Content/Report/ReportForTotalMoneyType.xlsx";
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
                    sheetReport.Cells["M4"].PutValue(fromDate.ToString("dd/MM/yyyy"));
                    sheetReport.Cells["Q4"].PutValue(toDate.ToString("dd/MM/yyyy"));
                    break;
                // Theo tháng
                case 2:
                    typeReport = "Chi tiết - Theo Tháng";
                    // Set from day and to day
                    sheetReport.Cells["M4"].PutValue(fromDate.ToString("MM/yyyy"));
                    sheetReport.Cells["Q4"].PutValue(toDate.ToString("MM/yyyy"));
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

            //// Trường hợp với report năm
            //if (typeID.Equals(3))
            //{
            //    sheetReport.Cells["Q7"].PutValue("");
            //}

            // Get data report ngày
            List<ReportForTotalMoneyType> listReportData = new List<ReportForTotalMoneyType>();
            List<ReportForTotalMoneyType> listReportDataClone = new List<ReportForTotalMoneyType>();
            List<ReportForTotalMoneyType> listReportDataConvert = new List<ReportForTotalMoneyType>();
            List<ReportForTotalMoneyType> listReportDataConvertClone = new List<ReportForTotalMoneyType>();

            int typeData = 0;

            switch (typeID)
            {
                // Theo ngày
                case 1:
                    // Theo nguyên tệ
                    listReportData = new ReportBL().SearchReportTMTForDay(fromDate, toDate, reportTypeID);
                    // Theo USD
                    listReportDataConvert = new ReportBL().SearchReportTMTForDayConvert(fromDate, toDate, reportTypeID);
                    break;
                // Theo tháng
                case 2:
                    // Theo nguyên tệ
                    listReportData = new ReportBL().SearchReportTMTForMonth(fromDate, toDate, reportTypeID);
                    // Theo USD
                    listReportDataConvert = new ReportBL().SearchReportTMTForMonthConvert(fromDate, toDate, reportTypeID);
                    break;
                // Theo năm
                default:
                    // Theo Nguyên tệ
                    listReportData = new ReportBL().SearchReportTMTForYear(fromDate, toDate, reportTypeID);
                    listReportDataClone = new List<ReportForTotalMoneyType>(listReportData);
                    // Theo USD
                    listReportDataConvert = new ReportBL().SearchReportTMTForYearConvert(fromDate, toDate, reportTypeID);
                    listReportDataConvertClone = new List<ReportForTotalMoneyType>(listReportDataConvert);
                    break;
            }

            // Set border
            Style style = new CellsFactory().CreateStyle();
            style.SetBorder(BorderType.TopBorder, CellBorderType.Thin, Color.Black);
            style.SetBorder(BorderType.BottomBorder, CellBorderType.Thin, Color.Black);
            style.SetBorder(BorderType.LeftBorder, CellBorderType.Thin, Color.Black);
            style.SetBorder(BorderType.RightBorder, CellBorderType.Thin, Color.Black);

            // Nguyên tệ
            DataTable dataTable = new DataTable();
            // Quy USD
            DataTable dataTableConvert = new DataTable();

            // Nguyên tệ
            if (listReportData.Count > 0)
            {
                foreach (ReportForTotalMoneyType item in listReportData)
                {
                    switch (typeID)
                    {
                        // Theo ngày
                        case 1:
                            item.ReportID = string.Concat("Ngày ", item.CreatedDate.Day, "/", item.CreatedDate.Month);
                            item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;
                            break;
                        // Theo tháng
                        case 2:
                            item.ReportID = string.Concat("Tháng ", item.Month, "/", item.Year);
                            item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;
                            break;
                        // Theo năm
                        default:
                            item.ReportID = string.Concat("Năm ", item.Year);
                            item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;
                            break;
                    }
                    item.Type = 0;
                }

                // Add dòng tổng vào list danh sách
                ReportForTotalMoneyType reportData = new ReportForTotalMoneyType()
                {
                    ReportID = "Tổng",
                    VND = listReportData.Sum(item => item.VND),
                    USD = listReportData.Sum(item => item.USD),
                    EUR = listReportData.Sum(item => item.EUR),
                    CAD = listReportData.Sum(item => item.CAD),
                    AUD = listReportData.Sum(item => item.AUD),
                    GBP = listReportData.Sum(item => item.GBP),
                    TongDS = listReportData.Sum(item => item.TongDS)
                };
                listReportData.Add(reportData);

                // Danh sách dataSet của báo cáo ngày
                DataSet dataReport = ConvertListObjectToDataSet(listReportData);
                // Tạo các col cho table
                dataTable = CreateDataTableFormart();
                // Đổ data vào datatble mới
                FillData(dataReport.Tables[0], dataTable);
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
                    int totalCol = 15 + 7;
                    // trường hợp với report chọn năm

                    //if (typeID.Equals(3))
                    //{
                    //    totalCol = 15 + 6;
                    //}

                    for (int b = 15; b < totalCol; b++)
                    {
                        // Giá trị của value trong table
                        string valueOfTable = dataTable.Rows[stepRow][stepColumn].ToString();

                        // Insert vào dòng cột xác định trong Excel
                        sheetReport.Cells[a, b].PutValue(valueOfTable, true);

                        // set style cho number
                        style.Custom = "#,##0.00";

                        // set border
                        sheetReport.Cells[a, b].SetStyle(style);

                        //// Cột tổng cộng
                        //if (b.Equals(totalCol - 1))
                        //{
                        //    sheetReport.Cells[a, b].PutValue(valueOfTable, true, true);
                        //    style.Font.IsBold = true;
                        //    sheetReport.Cells[a, b].SetStyle(style);
                        //}

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

            // quy đổi USD
            // Tạo cột dữ liệu nạp cho table
            string[] arraryMoneyType = { "VND", "USD", "EUR", "CAD", "AUD", "GBP", "Tổng" };
            //string[] arraryMoneyType = { "VND", "USD", "EUR", "CAD", "AUD", "GBP"};
            string[] arraryExcel = { "Q", "R", "S", "T", "U", "V","W" };
            //string[] arraryExcel = { "Q", "R", "S", "T", "U", "V" };

            // Row của table theo quy USD
            int countRow = 8 + listReportData.Count + 2;

            for (int i = 0; i < arraryExcel.Count(); i++)
            {
                sheetReport.Cells[string.Format("{0}{1}", arraryExcel[i], countRow)].PutValue(arraryMoneyType[i]);
                style.Font.IsBold = true;
                sheetReport.Cells[string.Format("{0}{1}", arraryExcel[i], countRow)].SetStyle(style);
            }

            if (listReportDataConvert.Count > 0)
            {
                foreach (ReportForTotalMoneyType item in listReportDataConvert)
                {
                    switch (typeID)
                    {
                        // Theo ngày
                        case 1:
                            item.ReportID = string.Concat("Ngày ", item.CreatedDate.Day, "/", item.CreatedDate.Month);
                            item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;
                            break;
                        // Theo tháng
                        case 2:
                            item.ReportID = string.Concat("Tháng ", item.Month, "/", item.Year);
                            item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;
                            break;
                        // Theo năm
                        default:
                            item.ReportID = string.Concat("Năm ", item.Year);
                            item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;
                            break;
                    }
                    item.Type = 0;
                }

                // Add dòng tổng vào list danh sách
                ReportForTotalMoneyType reportDataConvert = new ReportForTotalMoneyType()
                {
                    ReportID = "Tổng",
                    VND = listReportDataConvert.Sum(item => item.VND),
                    USD = listReportDataConvert.Sum(item => item.USD),
                    EUR = listReportDataConvert.Sum(item => item.EUR),
                    CAD = listReportDataConvert.Sum(item => item.CAD),
                    AUD = listReportDataConvert.Sum(item => item.AUD),
                    GBP = listReportDataConvert.Sum(item => item.GBP),
                    TongDS = listReportDataConvert.Sum(item => item.TongDS)
                };
                listReportDataConvert.Add(reportDataConvert);

                // Quy đổi USD
                dataTableConvert = CreateDataTableFormart();
                DataSet dataReportConvert = ConvertListObjectToDataSet(listReportDataConvert);
                FillData(dataReportConvert.Tables[0], dataTableConvert);

                // Vẽ biểu đồ đường cho table quy đổi USD

                // Create Chart Line
                //Chart reference
                Aspose.Cells.Charts.Chart leadSourceLine;
                //Add Pie Chart
                int chartIndex = sheetReport.Charts.Add(ChartType.Line, 6, 0, 30, 13);
                leadSourceLine = sheetReport.Charts[chartIndex];
                //Chart title
                leadSourceLine.Title.Text = "Tổng hợp - Loại tiền";
                leadSourceLine.Title.Font.Color = Color.Silver;


                // trường hợp với chọn báo theo ngày/tháng
                if (!typeID.Equals(3))
                {
                    // Adding SeriesCollection (chart data source) to the chart ranging from "A1" cell to "B3"
                    //string totalRowData = string.Format("Q{0}:W{1}", countRow, countRow + listReportDataConvert.Count - 2);
                    string totalRowData = string.Format("Q{0}:V{1}", countRow + 1, countRow + listReportDataConvert.Count - 1);
                    leadSourceLine.NSeries.Add(totalRowData, true);

                    // Set the category data covering the range A2:A5.
                    // Tổng số dòng cần hiển thị là số dòng hiện tại  + số dòng trong listData -2 (trong đó 1 là dòng tổng cuối cùng)
                    string categoryData = string.Format("P{0}:P{1}", countRow + 1, countRow + listReportDataConvert.Count - 1);
                    leadSourceLine.NSeries.CategoryData = categoryData;

                    // Canh hiển thị CategoryAxis nghiên phù hợp
                    leadSourceLine.CategoryAxis.TickLabels.RotationAngle = 45;

                    // Set the names of the chart series taken from cells.
                    leadSourceLine.NSeries[0].Name = "=Q7";
                    leadSourceLine.NSeries[1].Name = "=R7";
                    leadSourceLine.NSeries[2].Name = "=S7";
                    leadSourceLine.NSeries[3].Name = "=T7";
                    leadSourceLine.NSeries[4].Name = "=U7";
                    leadSourceLine.NSeries[5].Name = "=V7";
                    //leadSourceLine.NSeries[6].Name = "=W7";

                    // Set the 1st series fill color.
                    leadSourceLine.NSeries[0].Border.Color = Color.Orange;
                    leadSourceLine.NSeries[0].Area.Formatting = FormattingType.Custom;

                    // Set the 2nd series fill color.
                    leadSourceLine.NSeries[1].Border.Color = Color.Green;
                    leadSourceLine.NSeries[1].Area.Formatting = FormattingType.Custom;

                    // Set the 3rd series fill color.
                    leadSourceLine.NSeries[2].Border.Color = Color.Blue;
                    leadSourceLine.NSeries[2].Area.Formatting = FormattingType.Custom;

                    // Set the 4rd series fill color.
                    leadSourceLine.NSeries[3].Border.Color = Color.Red;
                    leadSourceLine.NSeries[3].Area.Formatting = FormattingType.Custom;

                    // Set the 4rd series fill color.
                    leadSourceLine.NSeries[4].Border.Color = Color.Pink;
                    leadSourceLine.NSeries[4].Area.Formatting = FormattingType.Custom;

                    // Set the 4rd series fill color.
                    leadSourceLine.NSeries[5].Border.Color = Color.Maroon;
                    leadSourceLine.NSeries[5].Area.Formatting = FormattingType.Custom;

                    //// Set the 4rd series fill color.
                    //leadSourceLine.NSeries[6].Border.Color = Color.FromArgb(23, 105, 255);
                    //leadSourceLine.NSeries[6].Area.Formatting = FormattingType.Custom;

                    //// Set plot area formatting as none and hide its border.
                    //leadSourceLine.PlotArea.Area.FillFormat.FillType = FillType.None;
                    //leadSourceLine.PlotArea.Border.IsVisible = false;
                }
                else
                {
                    // Tổng số dòng cần hiển thị là số dòng hiện tại  + số dòng trong listData -2 (trong đó 1 là dòng tổng cuối cùng)
                    int count = 0;
                    //string[] str = { "VND", "USD", "EUR", "CAD", "AUD", "GBP", "Tổng" };
                    string[] str = { "VND", "USD", "EUR", "CAD", "AUD", "GBP"};
                    
                    //foreach (ReportForTotalMoneyType item in listReportDataConvertClone)
                    //{
                    //    //string totalRowData = string.Concat("{", string.Format("{0}, {1}, {2}, {3}, {4}, {5}, {6}", item.VND, item.USD, item.EUR, item.CAD, item.AUD, item.GBP, item.TongDS), "}");
                    //    string totalRowData = string.Concat("{", string.Format("{0}, {1}, {2}, {3}, {4}, {5}", item.VND, item.USD, item.EUR, item.CAD, item.AUD, item.GBP), "}");
                    //    leadSourceLine.NSeries.Add(totalRowData, true);

                    //    string categoryLineData = string.Concat("{", string.Join(", ", str), "}");
                    //    leadSourceLine.NSeries.CategoryData = categoryLineData;

                    //    leadSourceLine.NSeries[count].Name = string.Concat("Năm ", item.Year);

                    //    switch (count)
                    //    {
                    //        case 1:
                    //            // Set the 1st series fill color.
                    //            leadSourceLine.NSeries[count].Area.ForegroundColor = Color.Orange;
                    //            leadSourceLine.NSeries[count].Area.Formatting = FormattingType.Custom;
                    //            break;
                    //        case 2:
                    //            // Set the 1st series fill color.
                    //            leadSourceLine.NSeries[count].Area.ForegroundColor = Color.Green;
                    //            leadSourceLine.NSeries[count].Area.Formatting = FormattingType.Custom;
                    //            break;
                    //        case 3:
                    //            // Set the 1st series fill color.
                    //            leadSourceLine.NSeries[count].Area.ForegroundColor = Color.Blue;
                    //            leadSourceLine.NSeries[count].Area.Formatting = FormattingType.Custom;
                    //            break;
                    //        case 4:
                    //            // Set the 1st series fill color.
                    //            leadSourceLine.NSeries[count].Area.ForegroundColor = Color.Red;
                    //            leadSourceLine.NSeries[count].Area.Formatting = FormattingType.Custom;
                    //            break;
                    //        default:
                    //            // Set the 1st series fill color.
                    //            leadSourceLine.NSeries[count].Area.ForegroundColor = Color.Pink;
                    //            leadSourceLine.NSeries[count].Area.Formatting = FormattingType.Custom;
                    //            break;
                    //    }
                    //    count++;
                    //}


                    //string totalRowData = string.Format("Q{0}:W{1}", countRow, countRow + listReportDataConvert.Count - 2);
                    string totalRowData = string.Format("Q{0}:V{1}", countRow, countRow + listReportDataConvert.Count - 1);
                    leadSourceLine.NSeries.Add(totalRowData, true);

                    // Set the category data covering the range A2:A5.
                    // Tổng số dòng cần hiển thị là số dòng hiện tại  + số dòng trong listData -2 (trong đó 1 là dòng tổng cuối cùng)
                    string categoryData = string.Format("P{0}:P{1}", countRow, countRow + listReportDataConvert.Count - 1);
                    leadSourceLine.NSeries.CategoryData = categoryData;

                    //// Canh hiển thị CategoryAxis nghiên phù hợp
                    //leadSourceLine.CategoryAxis.TickLabels.RotationAngle = 45;

                    // Set the names of the chart series taken from cells.
                    leadSourceLine.NSeries[0].Name = "=Q7";
                    leadSourceLine.NSeries[1].Name = "=R7";
                    leadSourceLine.NSeries[2].Name = "=S7";
                    leadSourceLine.NSeries[3].Name = "=T7";
                    leadSourceLine.NSeries[4].Name = "=U7";
                    leadSourceLine.NSeries[5].Name = "=V7";
                    //leadSourceLine.NSeries[6].Name = "=W7";

                    // Set the 1st series fill color.
                    leadSourceLine.NSeries[0].Border.Color = Color.Orange;
                    leadSourceLine.NSeries[0].Area.Formatting = FormattingType.Custom;

                    // Set the 2nd series fill color.
                    leadSourceLine.NSeries[1].Border.Color = Color.Green;
                    leadSourceLine.NSeries[1].Area.Formatting = FormattingType.Custom;

                    // Set the 3rd series fill color.
                    leadSourceLine.NSeries[2].Border.Color = Color.Blue;
                    leadSourceLine.NSeries[2].Area.Formatting = FormattingType.Custom;

                    // Set the 4rd series fill color.
                    leadSourceLine.NSeries[3].Border.Color = Color.Red;
                    leadSourceLine.NSeries[3].Area.Formatting = FormattingType.Custom;

                    // Set the 4rd series fill color.
                    leadSourceLine.NSeries[4].Border.Color = Color.Pink;
                    leadSourceLine.NSeries[4].Area.Formatting = FormattingType.Custom;

                    // Set the 4rd series fill color.
                    leadSourceLine.NSeries[5].Border.Color = Color.Maroon;
                    leadSourceLine.NSeries[5].Area.Formatting = FormattingType.Custom;

                }
                
                // Set plot area formatting as none and hide its border.
                leadSourceLine.PlotArea.Area.FillFormat.FillType = FillType.None;
                leadSourceLine.PlotArea.Border.IsVisible = false;

                // Set value axis major tick mark as none and hide axis line. 
                // Also set the color of value axis major grid lines.
                leadSourceLine.ValueAxis.MajorTickMark = TickMarkType.None;
                leadSourceLine.ValueAxis.AxisLine.IsVisible = false;
                leadSourceLine.ValueAxis.MajorGridLines.Color = Color.FromArgb(217, 217, 217);
                
                // Vẽ biểu đồ cột chồng cho table quy đổi USD
                // Create Chart Line
                //Chart reference
                Aspose.Cells.Charts.Chart leadSourceStackedColumn;
                //Add Pie Chart
                int chartColumnIndex = sheetReport.Charts.Add(ChartType.Column3DStacked, 35, 0, 60, 13);
                leadSourceStackedColumn = sheetReport.Charts[chartColumnIndex];


                //// Canh hiển thị CategoryAxis nghiên phù hợp
                //leadSourceStackedColumn.CategoryAxis.TickLabels.RotationAngle = 45;

                //Chart title
                leadSourceStackedColumn.Title.Text = "Doanh số chi trả theo từng loại tiền";
                leadSourceStackedColumn.Title.Font.Color = Color.Silver;

                // Trường hợp là báo cáo Excel theo năm
                if (!typeID.Equals(3))
                {
                    // Adding SeriesCollection (chart data source) to the chart ranging from "A1" cell to "B3"
                    //string totalRowColumnData = string.Format("Q{0}:W{1}", countRow, countRow + listReportDataConvert.Count - 2);
                    string totalRowColumnData = string.Format("Q{0}:V{1}", countRow, countRow + listReportDataConvert.Count - 1);
                    leadSourceStackedColumn.NSeries.Add(totalRowColumnData, true);

                    // Canh hiển thị CategoryAxis nghiên phù hợp
                    leadSourceStackedColumn.CategoryAxis.TickLabels.RotationAngle = 45;

                    // Set the category data covering the range A2:A5.
                    // Tổng số dòng cần hiển thị là số dòng hiện tại  + số dòng trong listData -2 (trong đó 1 là dòng tổng cuối cùng)
                    string categoryColumnData = string.Format("P{0}:P{1}", countRow, countRow + listReportDataConvert.Count - 1);
                    leadSourceStackedColumn.NSeries.CategoryData = categoryColumnData;

                    // Set the names of the chart series taken from cells.
                    leadSourceStackedColumn.NSeries[0].Name = "=Q7";
                    leadSourceStackedColumn.NSeries[1].Name = "=R7";
                    leadSourceStackedColumn.NSeries[2].Name = "=S7";
                    leadSourceStackedColumn.NSeries[3].Name = "=T7";
                    leadSourceStackedColumn.NSeries[4].Name = "=U7";
                    leadSourceStackedColumn.NSeries[5].Name = "=V7";
                    //leadSourceStackedColumn.NSeries[6].Name = "=W7";

                    // Set the 1st series fill color.
                    leadSourceStackedColumn.NSeries[0].Area.ForegroundColor = Color.Orange;
                    leadSourceStackedColumn.NSeries[0].Area.Formatting = FormattingType.Custom;

                    // Set the 2nd series fill color.
                    leadSourceStackedColumn.NSeries[1].Area.ForegroundColor = Color.Green;
                    leadSourceStackedColumn.NSeries[1].Area.Formatting = FormattingType.Custom;

                    // Set the 3rd series fill color.
                    leadSourceStackedColumn.NSeries[2].Area.ForegroundColor = Color.Blue;
                    leadSourceStackedColumn.NSeries[2].Area.Formatting = FormattingType.Custom;

                    // Set the 4rd series fill color.
                    leadSourceStackedColumn.NSeries[3].Area.ForegroundColor = Color.Red;
                    leadSourceStackedColumn.NSeries[3].Area.Formatting = FormattingType.Custom;

                    // Set the 4rd series fill color.
                    leadSourceStackedColumn.NSeries[4].Area.ForegroundColor = Color.Pink;
                    leadSourceStackedColumn.NSeries[4].Area.Formatting = FormattingType.Custom;

                    // Set the 4rd series fill color.
                    leadSourceStackedColumn.NSeries[5].Area.ForegroundColor = Color.Maroon;
                    leadSourceStackedColumn.NSeries[5].Area.Formatting = FormattingType.Custom;

                    //// Set the 4rd series fill color.
                    //leadSourceStackedColumn.NSeries[6].Area.ForegroundColor = Color.FromArgb(23, 105, 255);
                    //leadSourceStackedColumn.NSeries[6].Area.Formatting = FormattingType.Custom;
                }
                else
                {
                    // Tổng số dòng cần hiển thị là số dòng hiện tại  + số dòng trong listData -2 (trong đó 1 là dòng tổng cuối cùng)

                    //int count = 0;
                    ////string[] str = { "VND", "USD", "EUR", "CAD", "AUD", "GBP", "Tổng" };
                    //string[] str = { "VND", "USD", "EUR", "CAD", "AUD", "GBP" };


                    //foreach (ReportForTotalMoneyType item in listReportDataConvertClone)
                    //{
                    //    //string totalRowData = string.Concat("{", string.Format("{0}, {1}, {2}, {3}, {4}, {5}, {6}", item.VND, item.USD, item.EUR, item.CAD, item.AUD, item.GBP, item.TongDS), "}");
                    //    string totalRowData = string.Concat("{", string.Format("{0}, {1}, {2}, {3}, {4}, {5}", item.VND, item.USD, item.EUR, item.CAD, item.AUD, item.GBP), "}");
                    //    leadSourceStackedColumn.NSeries.Add(totalRowData, true);

                    //    string categoryLineData = string.Concat("{", string.Join(", ", str), "}");
                    //    leadSourceStackedColumn.NSeries.CategoryData = categoryLineData;

                    //    leadSourceStackedColumn.NSeries[count].Name = string.Concat("Năm ", item.Year);

                    //    switch (count)
                    //    {
                    //        case 1:
                    //            // Set the 1st series fill color.
                    //            leadSourceStackedColumn.NSeries[count].Area.ForegroundColor = Color.Orange;
                    //            leadSourceStackedColumn.NSeries[count].Area.Formatting = FormattingType.Custom;
                    //            break;
                    //        case 2:
                    //            // Set the 1st series fill color.
                    //            leadSourceStackedColumn.NSeries[count].Area.ForegroundColor = Color.Green;
                    //            leadSourceStackedColumn.NSeries[count].Area.Formatting = FormattingType.Custom;
                    //            break;
                    //        case 3:
                    //            // Set the 1st series fill color.
                    //            leadSourceStackedColumn.NSeries[count].Area.ForegroundColor = Color.Blue;
                    //            leadSourceStackedColumn.NSeries[count].Area.Formatting = FormattingType.Custom;
                    //            break;
                    //        case 4:
                    //            // Set the 1st series fill color.
                    //            leadSourceStackedColumn.NSeries[count].Area.ForegroundColor = Color.Pink;
                    //            leadSourceStackedColumn.NSeries[count].Area.Formatting = FormattingType.Custom;
                    //            break;
                    //        default:
                    //            // Set the 1st series fill color.
                    //            leadSourceStackedColumn.NSeries[count].Area.ForegroundColor = Color.Maroon;
                    //            leadSourceStackedColumn.NSeries[count].Area.Formatting = FormattingType.Custom;
                    //            break;
                    //    }
                    //    count++;
                    //}

                    // Adding SeriesCollection (chart data source) to the chart ranging from "A1" cell to "B3"
                    //string totalRowColumnData = string.Format("Q{0}:W{1}", countRow, countRow + listReportDataConvert.Count - 2);
                    string totalRowColumnData = string.Format("Q{0}:V{1}", countRow, countRow + listReportDataConvert.Count - 1);
                    leadSourceStackedColumn.NSeries.Add(totalRowColumnData, true);

                    //// Canh hiển thị CategoryAxis nghiên phù hợp
                    //leadSourceStackedColumn.CategoryAxis.TickLabels.RotationAngle = 45;

                    // Set the category data covering the range A2:A5.
                    // Tổng số dòng cần hiển thị là số dòng hiện tại  + số dòng trong listData -2 (trong đó 1 là dòng tổng cuối cùng)
                    string categoryColumnData = string.Format("P{0}:P{1}", countRow, countRow + listReportDataConvert.Count - 1);
                    leadSourceStackedColumn.NSeries.CategoryData = categoryColumnData;

                    // Set the names of the chart series taken from cells.
                    leadSourceStackedColumn.NSeries[0].Name = "=Q7";
                    leadSourceStackedColumn.NSeries[1].Name = "=R7";
                    leadSourceStackedColumn.NSeries[2].Name = "=S7";
                    leadSourceStackedColumn.NSeries[3].Name = "=T7";
                    leadSourceStackedColumn.NSeries[4].Name = "=U7";
                    leadSourceStackedColumn.NSeries[5].Name = "=V7";
                    //leadSourceStackedColumn.NSeries[6].Name = "=W7";

                    // Set the 1st series fill color.
                    leadSourceStackedColumn.NSeries[0].Area.ForegroundColor = Color.Orange;
                    leadSourceStackedColumn.NSeries[0].Area.Formatting = FormattingType.Custom;

                    // Set the 2nd series fill color.
                    leadSourceStackedColumn.NSeries[1].Area.ForegroundColor = Color.Green;
                    leadSourceStackedColumn.NSeries[1].Area.Formatting = FormattingType.Custom;

                    // Set the 3rd series fill color.
                    leadSourceStackedColumn.NSeries[2].Area.ForegroundColor = Color.Blue;
                    leadSourceStackedColumn.NSeries[2].Area.Formatting = FormattingType.Custom;

                    // Set the 4rd series fill color.
                    leadSourceStackedColumn.NSeries[3].Area.ForegroundColor = Color.Red;
                    leadSourceStackedColumn.NSeries[3].Area.Formatting = FormattingType.Custom;

                    // Set the 4rd series fill color.
                    leadSourceStackedColumn.NSeries[4].Area.ForegroundColor = Color.Pink;
                    leadSourceStackedColumn.NSeries[4].Area.Formatting = FormattingType.Custom;

                    // Set the 4rd series fill color.
                    leadSourceStackedColumn.NSeries[5].Area.ForegroundColor = Color.Maroon;
                    leadSourceStackedColumn.NSeries[5].Area.Formatting = FormattingType.Custom;
                }
                

                // Set plot area formatting as none and hide its border.
                leadSourceStackedColumn.PlotArea.Area.FillFormat.FillType = FillType.None;
                leadSourceStackedColumn.PlotArea.Border.IsVisible = false;

                // Set value axis major tick mark as none and hide axis line. 
                // Also set the color of value axis major grid lines.
                leadSourceStackedColumn.ValueAxis.MajorTickMark = TickMarkType.None;
                leadSourceStackedColumn.ValueAxis.AxisLine.IsVisible = false;
                leadSourceStackedColumn.ValueAxis.MajorGridLines.Color = Color.FromArgb(217, 217, 217);
            }

            // Tạo giá trị cho cột Quy USD
            Style st = new CellsFactory().CreateStyle();
            st.Font.IsBold = true;
            sheetReport.Cells[string.Format("P{0}", countRow - 1)].PutValue("Quy USD");
            sheetReport.Cells[string.Format("P{0}", countRow - 1)].SetStyle(st);
            sheetReport.Cells[string.Format("P{0}", countRow)].SetStyle(style);

            // Nạp dữ liệu table vào Excel theo quy đổi USD
            if (dataTableConvert.Rows.Count > 0)
            {
                int stepRow = 0;
                // total row = row start + số row hiện có
                //int totalRow = dataTableConvert.Rows.Count + 7;

                // Vị trí dòng của row Quy đổi USD
                // Countrow + 1
                // Số dòng của row
                int totalRowUSD = countRow + listReportDataConvert.Count;
                for (int a = countRow; a < totalRowUSD; a++)
                {
                    int stepColumn = 0;
                    // Số cột trong báo cáo cần hiển thị
                    // Tổng số cột hiển thị = Số cột hiển thị bắt đầu + tổng số cột cần hiển thị
                    // trường hợp với report chọn tháng
                    int totalCol = 15 + 8;
                    for (int b = 15; b < totalCol; b++)
                    {
                        // Giá trị của value trong table
                        string valueOfTable = dataTableConvert.Rows[stepRow][stepColumn].ToString();

                        // Insert vào dòng cột xác định trong Excel
                        sheetReport.Cells[a, b].PutValue(valueOfTable, true);

                        // set border
                        style.Custom = "#,##0.00";
                        sheetReport.Cells[a, b].SetStyle(style);

                        // Cột tổng cộng
                        if (b.Equals(totalCol - 1))
                        {
                            sheetReport.Cells[a, b].PutValue(valueOfTable, true, true);
                            style.Font.IsBold = true;
                            sheetReport.Cells[a, b].SetStyle(style);
                        }

                        // Trường hợp thuộc dòng cuối Tổng
                        if (a.Equals(totalRowUSD - 1))
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
        public ActionResult CreateExcelForGradationCompare(string gradationID, int year, string reportTypeID)
        {
            WorkbookDesigner designer = new WorkbookDesigner();
            string templatePath = "~/Content/Report/ReportGradationForTotalMoneyType.xlsx";
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

            // Tạo title cho doanh số chi trả loại hình dịch vụ
            string titleDS = "1. Theo doanh số chi trả loại hình dịch vụ";
            CreateTitle("A5", "E5", sheetReport, titleDS, 12);

            // Tạo title chotỷ trọng chi trả loại hình dịch vụ
            string titleTT = "2. Theo tỷ trọng chi trả loại hình dịch vụ";
            CreateTitle("A33", "E33", sheetReport, titleTT, 12);

            // Add tên cột cho bảng báo cáo
            // Set width cho column
            // Theo nguyên tệ
            sheetReport.Cells["P6"].PutValue("Nguyên tệ");

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

            // Theo quy USD
            sheetReport.Cells["P16"].PutValue("Quy USD");

            sheetReport.Cells.SetColumnWidthPixel(15, 80);
            sheetReport.Cells["Q17"].PutValue(string.Format("Lũy kế {0} \n {1}", text, year));
            sheetReport.Cells.SetColumnWidthPixel(16, 170);
            sheetReport.Cells["R17"].PutValue(string.Format("Lũy kế {0} \n {1}", text, year - 1));
            sheetReport.Cells.SetColumnWidthPixel(17, 190);
            sheetReport.Cells["S17"].PutValue(string.Format("Tăng giảm so với cùng kì \n năm {0} (%)", year - 1));
            sheetReport.Cells.SetColumnWidthPixel(18, 200);
            sheetReport.Cells["T17"].PutValue(string.Format("Tăng giảm so với cùng kì \n năm {0} (+/-)", year - 1));
            sheetReport.Cells.SetColumnWidthPixel(19, 200);
            sheetReport.Cells.SetRowHeight(6, 40);

            // Theo quy USD (%)
            sheetReport.Cells["R35"].PutValue("Đơn vị");
            sheetReport.Cells["S35"].PutValue("%");

            sheetReport.Cells.SetColumnWidthPixel(15, 80);
            sheetReport.Cells["Q35"].PutValue(string.Format("Lũy kế {0} \n {1}", text, year));
            sheetReport.Cells.SetColumnWidthPixel(16, 170);
            sheetReport.Cells["R35"].PutValue(string.Format("Lũy kế {0} \n {1}", text, year - 1));
            sheetReport.Cells.SetColumnWidthPixel(17, 190);
            sheetReport.Cells["S35"].PutValue(string.Format("Tăng giảm so với cùng kì \n năm {0} (%)", year - 1));
            sheetReport.Cells.SetColumnWidthPixel(19, 200);
            sheetReport.Cells.SetRowHeight(6, 40);
            // Set style cho các cột đã thay đổi
            string[] cellArray = { "P7", "Q7", "R7", "S7", "T7", "P17", "Q17", "R17", "S17", "T17", "Q35", "R35", "S35", "P35" };
            foreach (string item in cellArray)
            {
                //Get Cell's Style
                Style style = sheetReport.Cells[item].GetStyle();
                //Set Text Wrap property to true
                style.IsTextWrapped = true;
                style.SetBorder(BorderType.TopBorder, CellBorderType.Thin, Color.Black);
                style.SetBorder(BorderType.BottomBorder, CellBorderType.Thin, Color.Black);
                style.SetBorder(BorderType.LeftBorder, CellBorderType.Thin, Color.Black);
                style.SetBorder(BorderType.RightBorder, CellBorderType.Thin, Color.Black);
                style.Font.IsBold = true;
                // Canh giữa cho text trong table
                style.HorizontalAlignment = TextAlignmentType.Center;

                //Set Cell's Style
                sheetReport.Cells[item].SetStyle(style);
            }
            
            // Nguyên tệ
            List<ReportForTotalMoneyType> listReportData = new ReportBL().DataReportTMTForGradationCompare(year, int.Parse(gradationID), reportTypeID);
            // clone Object
            List<ReportForTotalMoneyType> listReportDataClone = new List<ReportForTotalMoneyType>(listReportData);

            // Danh sách doanh số chi trả theo loại hình dịch vụ Quy USD
            List<ReportForTotalMoneyType> listReportDataConvert = new ReportBL().DataReportTMTForGradationCompareConvert(year, int.Parse(gradationID), reportTypeID);
            // clone Object
            List<ReportForTotalMoneyType> listReportDataConvertClone = new List<ReportForTotalMoneyType>(listReportDataConvert);

            // Danh sách doanh số chi trả theo tỉ trọng
            List<ReportForTotalMoneyType> listReportDataPercent = new ReportBL().DataReportTMTForGradationComparePercent(year, int.Parse(gradationID), reportTypeID);
            // clone Object
            List<ReportForTotalMoneyType> listReportDataPercentClone = new List<ReportForTotalMoneyType>(listReportDataPercent);


            DataTable dataTable = new DataTable();
            DataTable dataTableConvert = new DataTable();

            string[] str = { "VND", "USD", "EUR", "CAD", "AUD", "GBP" };
            // Theo doanh số chi trả loại hình dịch vụ
            if (listReportData.Count.Equals(2))
            {
                // Tạo các cột cho datatable
                dataTable.Columns.Add("ReportID", typeof(String));
                dataTable.Columns.Add("AccumulateID1", typeof(double));
                dataTable.Columns.Add("AccumulateID2", typeof(double));
                dataTable.Columns.Add("CompareToIDPercent", typeof(double));
                dataTable.Columns.Add("CompareToID", typeof(double));

                double VNDCompare = listReportData[0].VND - listReportData[1].VND;
                double USDCompare = listReportData[0].USD - listReportData[1].USD;
                double EURCompare = listReportData[0].EUR - listReportData[1].EUR;
                double CADCompare = listReportData[0].CAD - listReportData[1].CAD;
                double AUDCompare = listReportData[0].AUD - listReportData[1].AUD;
                double GBPCompare = listReportData[0].GBP - listReportData[1].GBP;

                // add row vào table
                dataTable.Rows.Add(str[0], listReportData[0].VND, listReportData[1].VND, Math.Round(VNDCompare / listReportData[1].VND * 100, 2, MidpointRounding.ToEven), VNDCompare);
                dataTable.Rows.Add(str[1], listReportData[0].USD, listReportData[1].USD, Math.Round(USDCompare / listReportData[1].USD * 100, 2, MidpointRounding.ToEven), USDCompare);
                dataTable.Rows.Add(str[2], listReportData[0].EUR, listReportData[1].EUR, Math.Round(EURCompare / listReportData[1].EUR * 100, 2, MidpointRounding.ToEven), EURCompare);
                dataTable.Rows.Add(str[3], listReportData[0].CAD, listReportData[1].CAD, Math.Round(CADCompare / listReportData[1].CAD * 100, 2, MidpointRounding.ToEven), CADCompare);
                dataTable.Rows.Add(str[4], listReportData[0].AUD, listReportData[1].AUD, Math.Round(AUDCompare / listReportData[1].AUD * 100, 2, MidpointRounding.ToEven), AUDCompare);
                dataTable.Rows.Add(str[5], listReportData[0].GBP, listReportData[1].GBP, Math.Round(GBPCompare / listReportData[1].GBP * 100, 2, MidpointRounding.ToEven), GBPCompare);

                //DataRow row = dataTable.NewRow();
                //row["ReportID"] = "Tổng";
                //row["AccumulateID1"] = dataTable.Compute("Sum(AccumulateID1)", "");
                //row["AccumulateID2"] = dataTable.Compute("Sum(AccumulateID2)", "");

                //// Sum row tổng compare month
                //double sumCompareMonth = (double)row["AccumulateID1"] - (double)row["AccumulateID2"];
                
                //row["CompareToIDPercent"] = Math.Round(sumCompareMonth / dataTable.Rows.Count, 2, MidpointRounding.ToEven);
                //row["CompareToID"] = Math.Round(sumCompareMonth, 2, MidpointRounding.ToEven);
                //dataTable.Rows.Add(row);
                
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
            if (listReportDataConvertClone.Count.Equals(2))
            {
                // Tạo các cột cho dataTableConvert
                dataTableConvert.Columns.Add("ReportID", typeof(String));
                dataTableConvert.Columns.Add("AccumulateID1", typeof(double));
                dataTableConvert.Columns.Add("AccumulateID2", typeof(double));
                dataTableConvert.Columns.Add("CompareToIDPercent", typeof(double));
                dataTableConvert.Columns.Add("CompareToID", typeof(double));

                double VNDCompare = listReportDataConvertClone[0].VND - listReportDataConvertClone[1].VND;
                double USDCompare = listReportDataConvertClone[0].USD - listReportDataConvertClone[1].USD;
                double EURCompare = listReportDataConvertClone[0].EUR - listReportDataConvertClone[1].EUR;
                double CADCompare = listReportDataConvertClone[0].CAD - listReportDataConvertClone[1].CAD;
                double AUDCompare = listReportDataConvertClone[0].AUD - listReportDataConvertClone[1].AUD;
                double GBPCompare = listReportDataConvertClone[0].GBP - listReportDataConvertClone[1].GBP;

                // add row vào table
                dataTableConvert.Rows.Add(str[0], listReportDataConvertClone[0].VND, listReportDataConvertClone[1].VND, Math.Round(VNDCompare / listReportDataConvertClone[1].VND * 100, 2, MidpointRounding.ToEven), VNDCompare);
                dataTableConvert.Rows.Add(str[1], listReportDataConvertClone[0].USD, listReportDataConvertClone[1].USD, Math.Round(USDCompare / listReportDataConvertClone[1].USD * 100, 2, MidpointRounding.ToEven), USDCompare);
                dataTableConvert.Rows.Add(str[2], listReportDataConvertClone[0].EUR, listReportDataConvertClone[1].EUR, Math.Round(EURCompare / listReportDataConvertClone[1].EUR * 100, 2, MidpointRounding.ToEven), EURCompare);
                dataTableConvert.Rows.Add(str[3], listReportDataConvertClone[0].CAD, listReportDataConvertClone[1].CAD, Math.Round(CADCompare / listReportDataConvertClone[1].CAD * 100, 2, MidpointRounding.ToEven), CADCompare);
                dataTableConvert.Rows.Add(str[4], listReportDataConvertClone[0].AUD, listReportDataConvertClone[1].AUD, Math.Round(AUDCompare / listReportDataConvertClone[1].AUD * 100, 2, MidpointRounding.ToEven), AUDCompare);
                dataTableConvert.Rows.Add(str[5], listReportDataConvertClone[0].GBP, listReportDataConvertClone[1].GBP, Math.Round(GBPCompare / listReportDataConvertClone[1].GBP * 100, 2, MidpointRounding.ToEven), GBPCompare);

                DataRow row = dataTableConvert.NewRow();
                row["ReportID"] = "Tổng";
                row["AccumulateID1"] = dataTableConvert.Compute("Sum(AccumulateID1)", "");
                row["AccumulateID2"] = dataTableConvert.Compute("Sum(AccumulateID2)", "");

                // Sum row tổng compare month
                double sumCompareMonth = (double)row["AccumulateID1"] - (double)row["AccumulateID2"];
                
                row["CompareToIDPercent"] = Math.Round((double)dataTableConvert.Compute("AVG(CompareToIDPercent)", ""), 2, MidpointRounding.ToEven);

                row["CompareToID"] = Math.Round(sumCompareMonth, 2, MidpointRounding.ToEven);
                dataTableConvert.Rows.Add(row);

                // Create Chart column
                //Chart reference
                Aspose.Cells.Charts.Chart leadSourceColumn;
                //Add Pie Chart
                int chartIndex = sheetReport.Charts.Add(ChartType.Column3DClustered, 6, 0, 30, 13);
                leadSourceColumn = sheetReport.Charts[chartIndex];
                

                //Chart title
                leadSourceColumn.Title.Text = string.Format("Doanh số loại tiền chi trả \n Giai đoạn: {0}", text);
                leadSourceColumn.Title.Font.Color = Color.Silver;

                // Adding SeriesCollection (chart data source) to the chart ranging from "A1" cell to "B3"
                string totalRowData = "Q18:R23";
                leadSourceColumn.NSeries.Add(totalRowData, true);

                // Set the category data covering the range A2:A5.
                // Tổng số dòng cần hiển thị là số dòng hiện tại  + số dòng trong listData -2 (trong đó 1 là dòng tổng cuối cùng)
                string categoryData = "P8:P23";
                leadSourceColumn.NSeries.CategoryData = categoryData;


                // Set the names of the chart series taken from cells.
                leadSourceColumn.NSeries[0].Name = "=Q7";
                leadSourceColumn.NSeries[1].Name = "=R7";

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

                // Set border
                Style style = new CellsFactory().CreateStyle();
                style.SetBorder(BorderType.TopBorder, CellBorderType.Thin, Color.Black);
                style.SetBorder(BorderType.BottomBorder, CellBorderType.Thin, Color.Black);
                style.SetBorder(BorderType.LeftBorder, CellBorderType.Thin, Color.Black);
                style.SetBorder(BorderType.RightBorder, CellBorderType.Thin, Color.Black);

                if (dataTableConvert.Rows.Count > 0)
                {
                    int stepRow = 0;
                    // total row = row start + số row hiện có
                    int totalRow = dataTableConvert.Rows.Count + 17;
                    // Số dòng của row
                    for (int a = 17; a < totalRow; a++)
                    {
                        int stepColumn = 0;
                        // Số cột trong báo cáo cần hiển thị
                        // Tổng số cột hiển thị = Số cột hiển thị bắt đầu + tổng số cột cần hiển thị
                        int totalCol = 15 + 5;
                        for (int b = 15; b < totalCol; b++)
                        {
                            // Giá trị của value trong table
                            string valueOfTable = dataTableConvert.Rows[stepRow][stepColumn].ToString();

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
                foreach (ReportForTotalMoneyType item in listReportDataPercentClone)
                {
                    item.ReportID = string.Concat("Lũy kế ", text, " ", year);
                    if (!check)
                    {
                        item.ReportID = string.Concat("Lũy kế ", text, " ", year - 1);
                    }
                    item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;
                    item.Type = 0;
                    // Set lại giá trị cho check để lấy giá trị của năm trước
                    check = false;
                }

                ReportForTotalMoneyType dataPieYear = null;
                ReportForTotalMoneyType dataPieLastYear = null;
                // Data report năm hiện tại nhập vào
                dataPieYear = listReportDataPercentClone.Find(x => x.Year == year.ToString());
                // Data report năm ngoái so với năm hiện tại nhập vào
                dataPieLastYear = listReportDataPercentClone.Find(x => x.Year == (year - 1).ToString());

                if (dataPieYear != null)
                {
                    //Add Pie Chart
                    Aspose.Cells.Charts.Chart leadSourcePie;
                    int chartIndex = sheetReport.Charts.Add(ChartType.Pie3D, 34, 0, 49, 5);
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
                    string totalRowData = string.Concat("{", string.Format("{0}, {1}, {2}, {3}, {4}, {5}", dataPieYear.VND, dataPieYear.USD, dataPieYear.EUR, dataPieYear.CAD, dataPieYear.AUD, dataPieYear.GBP), "}");
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
                    int chartIndex = sheetReport.Charts.Add(ChartType.Pie3D, 34, 8, 49, 13);
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
                    string totalRowData = string.Concat("{", string.Format("{0}, {1}, {2}, {3}, {4}, {5}", dataPieYear.VND, dataPieYear.USD, dataPieYear.EUR, dataPieYear.CAD, dataPieYear.AUD, dataPieYear.GBP), "}");
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
                double VNDCompare = listReportDataPercentClone[0].VND - listReportDataPercentClone[1].VND;
                double USDCompare = listReportDataPercentClone[0].USD - listReportDataPercentClone[1].USD;
                double EURCompare = listReportDataPercentClone[0].EUR - listReportDataPercentClone[1].EUR;
                double CADCompare = listReportDataPercentClone[0].CAD - listReportDataPercentClone[1].CAD;
                double AUDCompare = listReportDataPercentClone[0].AUD - listReportDataPercentClone[1].AUD;
                double GBPCompare = listReportDataPercentClone[0].GBP - listReportDataPercentClone[1].GBP;

                // Tạo các cột cho datatable
                DataTable dataTablePie = new DataTable();

                dataTablePie.Columns.Add("ReportID", typeof(String));
                dataTablePie.Columns.Add("AccumulateID1", typeof(double));
                dataTablePie.Columns.Add("AccumulateID2", typeof(double));
                dataTablePie.Columns.Add("CompareToIDPercent", typeof(double));

                // add row vào table
                dataTablePie.Rows.Add(str[0], listReportDataPercentClone[0].VND, listReportDataPercentClone[1].VND, VNDCompare);
                dataTablePie.Rows.Add(str[1], listReportDataPercentClone[0].USD, listReportDataPercentClone[1].USD, USDCompare);
                dataTablePie.Rows.Add(str[2], listReportDataPercentClone[0].EUR, listReportDataPercentClone[1].EUR, EURCompare);
                dataTablePie.Rows.Add(str[3], listReportDataPercentClone[0].CAD, listReportDataPercentClone[1].CAD, CADCompare);
                dataTablePie.Rows.Add(str[4], listReportDataPercentClone[0].AUD, listReportDataPercentClone[1].AUD, AUDCompare);
                dataTablePie.Rows.Add(str[5], listReportDataPercentClone[0].GBP, listReportDataPercentClone[1].GBP, GBPCompare);

                DataRow row = dataTablePie.NewRow();
                row["ReportID"] = "Tổng";
                row["AccumulateID1"] = dataTablePie.Compute("Sum(AccumulateID1)", "");
                row["AccumulateID2"] = dataTablePie.Compute("Sum(AccumulateID2)", "");

                // Sum row tổng compare month
                double sumCompareMonth = (double)row["AccumulateID1"] - (double)row["AccumulateID2"];

                row["CompareToIDPercent"] = Math.Round(sumCompareMonth, 2, MidpointRounding.ToEven);
                dataTablePie.Rows.Add(row);

                // Set border
                Style style = new CellsFactory().CreateStyle();
                style.SetBorder(BorderType.TopBorder, CellBorderType.Thin, Color.Black);
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
            db.Columns.Add("VND", typeof(double));
            db.Columns.Add("USD", typeof(double));
            db.Columns.Add("EUR", typeof(double));
            db.Columns.Add("CAD", typeof(double));
            db.Columns.Add("AUD", typeof(double));
            db.Columns.Add("GBP", typeof(double));
            db.Columns.Add("TongDS", typeof(double));
            return db;
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
            string templatePath = "~/Content/Report/ReportForTotalMoneyType.xlsx";
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

            // Tạo title cho doanh số chi trả loại hình dịch vụ
            string titleDS = "1. Theo doanh số chi trả loại hình dịch vụ";
            CreateTitle("A5", "E5", sheetReport, titleDS, 12);

            // Tạo title chotỷ trọng chi trả loại hình dịch vụ
            string titleTT = "2. Theo tỷ trọng chi trả loại hình dịch vụ";
            CreateTitle("A33", "E33", sheetReport, titleTT, 12);

            // Add tên cột cho bảng báo cáo
            // Set width cho column
            // Theo doanh số
            //sheetReport.Cells.SetColumnWidthPixel(15, 80);
            sheetReport.Cells["S6"].PutValue("");
            sheetReport.Cells["T6"].PutValue("");

            // Nguyên tệ
            sheetReport.Cells["Q7"].PutValue(string.Format("Tháng {0}/{1}", month, year));
            //sheetReport.Cells.SetColumnWidthPixel(16, 170);

            sheetReport.Cells["R7"].PutValue(string.Format("Tháng {0}/{1}", month - 1, year));
            if (month == 1)
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

            // Quy USD
            // In đậm
            sheetReport.Cells["P16"].PutValue("Quy USD");
            Style st = sheetReport.Cells["P16"].GetStyle();
            st.Font.IsBold = true;
            sheetReport.Cells["P16"].SetStyle(st);

            sheetReport.Cells["Q17"].PutValue(string.Format("Tháng {0}/{1}", month, year));
            //sheetReport.Cells.SetColumnWidthPixel(16, 170);
            sheetReport.Cells["R17"].PutValue(string.Format("Tháng {0}/{1}", month - 1, year));
            if (month == 1)
            {
                sheetReport.Cells["R17"].PutValue(string.Format("Tháng {0}/{1}", 12, year - 1));
            }

            //sheetReport.Cells.SetColumnWidthPixel(17, 190);
            sheetReport.Cells["S17"].PutValue(string.Format("Tháng {0}/{1}", month, year - 1));
            sheetReport.Cells.SetColumnWidthPixel(19, 200);
            sheetReport.Cells["T17"].PutValue("Tăng giảm so với tháng (+/-)");
            sheetReport.Cells.SetColumnWidthPixel(20, 200);
            sheetReport.Cells["U17"].PutValue("Tăng giảm so với tháng (%)");
            sheetReport.Cells.SetColumnWidthPixel(21, 200);
            sheetReport.Cells["V17"].PutValue("Tăng giảm so với cùng kì \n năm trước (+/-)");
            sheetReport.Cells.SetColumnWidthPixel(22, 200);
            sheetReport.Cells["W17"].PutValue("Tăng giảm so với cùng kì \n năm trước (%)");
            sheetReport.Cells.SetRowHeight(16, 40);

            // Tạo cột cho tỉ trọng %
            //sheetReport.Cells.SetColumnWidthPixel(15, 200);
            sheetReport.Cells["Q35"].PutValue(string.Format("Tháng {0}/{1}", month, year));
            //sheetReport.Cells.SetColumnWidthPixel(16, 200);

            sheetReport.Cells["R35"].PutValue(string.Format("Tháng {0}/{1}", month - 1, year));
            if (month == 1)
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

            // Set style cho các cột đã thay đổi
            string[] cellArray = { "Q7", "R7", "S7", "T7", "U7", "V7", "W7", "P35", "Q35", "R35", "S35", "T35", "U35", "P17", "Q17", "R17", "S17", "T17", "U17", "V17", "W17" };
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
                style.SetBorder(BorderType.LeftBorder, CellBorderType.Thin, Color.Black);
                style.SetBorder(BorderType.RightBorder, CellBorderType.Thin, Color.Black);
                style.SetBorder(BorderType.BottomBorder, CellBorderType.Thin, Color.Black);
                style.Font.IsBold = true;
                //Set Cell's Style
                sheetReport.Cells[item].SetStyle(style);
            }

            // Data theo nguyên tệ
            List<ReportForTotalMoneyType> listReportData = new ReportBL().DataReportTMTCompareForMonth(year, month, reportTypeID);
            // clone Object
            List<ReportForTotalMoneyType> listReportDataClone = new List<ReportForTotalMoneyType>(listReportData);

            // Data theo USD
            List<ReportForTotalMoneyType> listReportDataConvert = new ReportBL().DataReportTMTCompareForMonthConvert(year, month, reportTypeID);
            // clone Object
            List<ReportForTotalMoneyType> listReportDataConvertClone = new List<ReportForTotalMoneyType>(listReportDataConvert);

            // Danh sách tỉ trọng chi trả theo loại hình dịch vụ
            List<ReportForTotalMoneyType> listReportDataPercent = new ReportBL().DataReportTMTCompareForMonthPercent(year, month, reportTypeID);
            // clone Object
            List<ReportForTotalMoneyType> listReportDataPercentClone = new List<ReportForTotalMoneyType>(listReportDataPercent);

            DataTable dataTable = new DataTable();
            DataTable dataTableConvert = new DataTable();

            string[] str = { "VND", "USD", "EUR", "CAD", "AUD", "GBP" };

            // Theo nguyên tệ
            if (listReportData.Count.Equals(3))
            {
                // Tạo các cột cho datatable
                dataTable.Columns.Add("ReportID", typeof(String));
                dataTable.Columns.Add("AccumulateID1", typeof(double));
                dataTable.Columns.Add("AccumulateID2", typeof(double));
                dataTable.Columns.Add("AccumulateID3", typeof(double));
                dataTable.Columns.Add("CompareToMonth", typeof(double));
                dataTable.Columns.Add("CompareToMonthPercent", typeof(double));
                dataTable.Columns.Add("CompareToMonthLastYear", typeof(double));
                dataTable.Columns.Add("CompareToMonthLastYearPercent", typeof(double));
                dataTable.PrimaryKey = new DataColumn[] { dataTable.Columns["ReportID"] };

                // tháng hiện tại
                double VNDCompareMonth = listReportData[0].VND - listReportData[1].VND;
                double USDCompareMonth = listReportData[0].USD - listReportData[1].USD;
                double EURCompareMonth = listReportData[0].EUR - listReportData[1].EUR;
                double CADCompareMonth = listReportData[0].CAD - listReportData[1].CAD;
                double AUDCompareMonth = listReportData[0].AUD - listReportData[1].AUD;
                double GBPCompareMonth = listReportData[0].GBP - listReportData[1].GBP;

                // tháng cùng kì năm trước
                double VNDCompareMonthLastYear = listReportData[0].VND - listReportData[2].VND;
                double USDCompareMonthLastYear = listReportData[0].USD - listReportData[2].USD;
                double EURCompareMonthLastYear = listReportData[0].EUR - listReportData[2].EUR;
                double CADCompareMonthLastYear = listReportData[0].CAD - listReportData[2].CAD;
                double AUDCompareMonthLastYear = listReportData[0].AUD - listReportData[2].AUD;
                double GBPCompareMonthLastYear = listReportData[0].GBP - listReportData[2].GBP;

                // add row vào table
                dataTable.Rows.Add(str[0], listReportData[0].VND, listReportData[1].VND, listReportData[2].VND
                    , VNDCompareMonth
                    , listReportData[1].VND == 0 ? listReportData[1].VND : Math.Round(VNDCompareMonth / listReportData[1].VND * 100, 2, MidpointRounding.ToEven)
                    , VNDCompareMonthLastYear
                    , listReportData[2].VND == 0 ? listReportData[2].VND : Math.Round(VNDCompareMonthLastYear / listReportData[2].VND * 100, 2, MidpointRounding.ToEven));

                dataTable.Rows.Add(str[1], listReportData[0].USD, listReportData[1].USD, listReportData[2].USD
                    , USDCompareMonth
                    , listReportData[1].USD == 0 ? listReportData[1].USD : Math.Round(USDCompareMonth / listReportData[1].USD * 100, 2, MidpointRounding.ToEven)
                    , USDCompareMonthLastYear
                    , listReportData[2].USD == 0 ? listReportData[2].USD : Math.Round(USDCompareMonthLastYear / listReportData[2].USD * 100, 2, MidpointRounding.ToEven));

                dataTable.Rows.Add(str[2], listReportData[0].EUR, listReportData[1].EUR, listReportData[2].EUR
                    , EURCompareMonth
                    , listReportData[1].EUR == 0 ? listReportData[1].EUR : Math.Round(EURCompareMonth / listReportData[1].EUR * 100, 2, MidpointRounding.ToEven)
                    , EURCompareMonthLastYear
                    , listReportData[2].EUR == 0 ? listReportData[2].EUR : Math.Round(EURCompareMonthLastYear / listReportData[2].EUR * 100, 2, MidpointRounding.ToEven));

                dataTable.Rows.Add(str[3], listReportData[0].CAD, listReportData[1].CAD, listReportData[2].CAD
                    , CADCompareMonth
                    , listReportData[1].CAD == 0 ? listReportData[1].CAD : Math.Round(CADCompareMonth / listReportData[1].CAD * 100, 2, MidpointRounding.ToEven)
                    , CADCompareMonthLastYear
                    , listReportData[2].CAD == 0 ? listReportData[2].CAD : Math.Round(CADCompareMonthLastYear / listReportData[2].CAD * 100, 2, MidpointRounding.ToEven));

                dataTable.Rows.Add(str[4], listReportData[0].AUD, listReportData[1].AUD, listReportData[2].AUD
                    , AUDCompareMonth
                    , listReportData[1].AUD == 0 ? listReportData[1].AUD : Math.Round(AUDCompareMonth / listReportData[1].AUD * 100, 2, MidpointRounding.ToEven)
                    , AUDCompareMonthLastYear
                    , listReportData[2].AUD == 0 ? listReportData[2].AUD : Math.Round(AUDCompareMonthLastYear / listReportData[2].AUD * 100, 2, MidpointRounding.ToEven));

                dataTable.Rows.Add(str[5], listReportData[0].GBP, listReportData[1].GBP, listReportData[2].GBP
                    , GBPCompareMonth
                    , listReportData[1].GBP == 0 ? listReportData[1].GBP : Math.Round(GBPCompareMonth / listReportData[1].GBP * 100, 2, MidpointRounding.ToEven)
                    , GBPCompareMonthLastYear
                    , listReportData[2].GBP == 0 ? listReportData[2].GBP : Math.Round(GBPCompareMonthLastYear / listReportData[2].GBP * 100, 2, MidpointRounding.ToEven));
                
                
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

                            // Các cột tăng giảm theo (+/-) và (%)
                            if (b.Equals(totalCol - 1) || b.Equals(totalCol - 2) 
                                || b.Equals(totalCol - 3) || b.Equals(totalCol - 4))
                            {
                                sheetReport.Cells[a, b].PutValue(valueOfTable, true, true);
                                style.Font.IsBold = true;
                                sheetReport.Cells[a, b].SetStyle(style);
                            }

                            // Trường hợp thuộc dòng cuối
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

            // Quy theo USD
            if (listReportDataConvert.Count.Equals(3))
            {
                // Tạo các cột cho dataTableConvert
                dataTableConvert.Columns.Add("ReportID", typeof(String));
                dataTableConvert.Columns.Add("AccumulateID1", typeof(double));
                dataTableConvert.Columns.Add("AccumulateID2", typeof(double));
                dataTableConvert.Columns.Add("AccumulateID3", typeof(double));
                dataTableConvert.Columns.Add("CompareToMonth", typeof(double));
                dataTableConvert.Columns.Add("CompareToMonthPercent", typeof(double));
                dataTableConvert.Columns.Add("CompareToMonthLastYear", typeof(double));
                dataTableConvert.Columns.Add("CompareToMonthLastYearPercent", typeof(double));
                dataTableConvert.PrimaryKey = new DataColumn[] { dataTableConvert.Columns["ReportID"] };

                // tháng hiện tại
                double VNDCompareMonth = listReportDataConvert[0].VND - listReportDataConvert[1].VND;
                double USDCompareMonth = listReportDataConvert[0].USD - listReportDataConvert[1].USD;
                double EURCompareMonth = listReportDataConvert[0].EUR - listReportDataConvert[1].EUR;
                double CADCompareMonth = listReportDataConvert[0].CAD - listReportDataConvert[1].CAD;
                double AUDCompareMonth = listReportDataConvert[0].AUD - listReportDataConvert[1].AUD;
                double GBPCompareMonth = listReportDataConvert[0].GBP - listReportDataConvert[1].GBP;

                // tháng cùng kì năm trước
                double VNDCompareMonthLastYear = listReportDataConvert[0].VND - listReportDataConvert[2].VND;
                double USDCompareMonthLastYear = listReportDataConvert[0].USD - listReportDataConvert[2].USD;
                double EURCompareMonthLastYear = listReportDataConvert[0].EUR - listReportDataConvert[2].EUR;
                double CADCompareMonthLastYear = listReportDataConvert[0].CAD - listReportDataConvert[2].CAD;
                double AUDCompareMonthLastYear = listReportDataConvert[0].AUD - listReportDataConvert[2].AUD;
                double GBPCompareMonthLastYear = listReportDataConvert[0].GBP - listReportDataConvert[2].GBP;

                // add row vào table
                dataTableConvert.Rows.Add(str[0], listReportDataConvert[0].VND, listReportDataConvert[1].VND, listReportDataConvert[2].VND
                    , VNDCompareMonth
                    , listReportDataConvert[1].VND == 0 ? listReportDataConvert[1].VND : Math.Round(VNDCompareMonth / listReportDataConvert[1].VND * 100, 2, MidpointRounding.ToEven)
                    , VNDCompareMonthLastYear
                    , listReportDataConvert[2].VND == 0 ? listReportDataConvert[2].VND : Math.Round(VNDCompareMonthLastYear / listReportDataConvert[2].VND * 100, 2, MidpointRounding.ToEven));

                dataTableConvert.Rows.Add(str[1], listReportDataConvert[0].USD, listReportDataConvert[1].USD, listReportDataConvert[2].USD
                    , USDCompareMonth
                    , listReportDataConvert[1].USD == 0 ? listReportDataConvert[1].USD : Math.Round(USDCompareMonth / listReportDataConvert[1].USD * 100, 2, MidpointRounding.ToEven)
                    , USDCompareMonthLastYear
                    , listReportDataConvert[2].USD == 0 ? listReportDataConvert[2].USD : Math.Round(USDCompareMonthLastYear / listReportDataConvert[2].USD * 100, 2, MidpointRounding.ToEven));

                dataTableConvert.Rows.Add(str[2], listReportDataConvert[0].EUR, listReportDataConvert[1].EUR, listReportDataConvert[2].EUR
                    , EURCompareMonth
                    , listReportDataConvert[1].EUR == 0 ? listReportDataConvert[1].EUR : Math.Round(EURCompareMonth / listReportDataConvert[1].EUR * 100, 2, MidpointRounding.ToEven)
                    , EURCompareMonthLastYear
                    , listReportDataConvert[2].EUR == 0 ? listReportDataConvert[2].EUR : Math.Round(EURCompareMonthLastYear / listReportDataConvert[2].EUR * 100, 2, MidpointRounding.ToEven));

                dataTableConvert.Rows.Add(str[3], listReportDataConvert[0].CAD, listReportDataConvert[1].CAD, listReportDataConvert[2].CAD
                    , CADCompareMonth
                    , listReportDataConvert[1].CAD == 0 ? listReportDataConvert[1].CAD : Math.Round(CADCompareMonth / listReportDataConvert[1].CAD * 100, 2, MidpointRounding.ToEven)
                    , CADCompareMonthLastYear
                    , listReportDataConvert[2].CAD == 0 ? listReportDataConvert[2].CAD : Math.Round(CADCompareMonthLastYear / listReportDataConvert[2].CAD * 100, 2, MidpointRounding.ToEven));

                dataTableConvert.Rows.Add(str[4], listReportDataConvert[0].AUD, listReportDataConvert[1].AUD, listReportDataConvert[2].AUD
                    , AUDCompareMonth
                    , listReportDataConvert[1].AUD == 0 ? listReportDataConvert[1].AUD : Math.Round(AUDCompareMonth / listReportDataConvert[1].AUD * 100, 2, MidpointRounding.ToEven)
                    , AUDCompareMonthLastYear
                    , listReportDataConvert[2].AUD == 0 ? listReportDataConvert[2].AUD : Math.Round(AUDCompareMonthLastYear / listReportDataConvert[2].AUD * 100, 2, MidpointRounding.ToEven));

                dataTableConvert.Rows.Add(str[5], listReportDataConvert[0].GBP, listReportDataConvert[1].GBP, listReportDataConvert[2].GBP
                    , GBPCompareMonth
                    , listReportDataConvert[1].GBP == 0 ? listReportDataConvert[1].GBP : Math.Round(GBPCompareMonth / listReportDataConvert[1].GBP * 100, 2, MidpointRounding.ToEven)
                    , GBPCompareMonthLastYear
                    , listReportDataConvert[2].GBP == 0 ? listReportDataConvert[2].GBP : Math.Round(GBPCompareMonthLastYear / listReportDataConvert[2].GBP * 100, 2, MidpointRounding.ToEven));

                DataRow row = dataTableConvert.NewRow();
                row["ReportID"] = "Tổng";
                row["AccumulateID1"] = dataTableConvert.Compute("Sum(AccumulateID1)", "");
                row["AccumulateID2"] = dataTableConvert.Compute("Sum(AccumulateID2)", "");
                row["AccumulateID3"] = dataTableConvert.Compute("Sum(AccumulateID3)", "");

                row["CompareToMonth"] = dataTableConvert.Compute("Sum(CompareToMonth)", "");
                row["CompareToMonthPercent"] = Math.Round((double)dataTableConvert.Compute("Sum(CompareToMonthPercent)", "") /6, 2, MidpointRounding.ToEven);
                row["CompareToMonthLastYear"] = dataTableConvert.Compute("Sum(CompareToMonthLastYear)", "");
                // Với 6 loại tiền
                row["CompareToMonthLastYearPercent"] = Math.Round((double)dataTableConvert.Compute("Sum(CompareToMonthLastYearPercent)", "") / 6, 2, MidpointRounding.ToEven);

                dataTableConvert.Rows.Add(row);

                // Create Chart Line
                //Chart reference
                Aspose.Cells.Charts.Chart leadSourceColumn;
                //Add Pie Chart
                int chartIndex = sheetReport.Charts.Add(ChartType.Column3DClustered, 6, 0, 30, 13);
                leadSourceColumn = sheetReport.Charts[chartIndex];
                

                //Chart title
                leadSourceColumn.Title.Text = string.Format("Doanh số loại tiền chi trả (Quy USD) tháng {0}/{1} \n so với tháng trước và so với cùng kì năm trước", month, year);
                leadSourceColumn.Title.Font.Color = Color.Silver;

                // Set width cho column
                sheetReport.Cells.SetColumnWidthPixel(15, 220);

                int count = 0;
                foreach (ReportForTotalMoneyType item in listReportDataConvertClone)
                {
                    item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;
                    //string totalRowData = string.Concat("{", string.Format("{0}, {1}, {2}, {3}, {4}, {5}, {6}", item.VND, item.USD, item.EUR, item.CAD, item.AUD, item.GBP, item.TongDS), "}");
                    string totalRowData = string.Concat("{", string.Format("{0}, {1}, {2}, {3}, {4}, {5}", item.VND, item.USD, item.EUR, item.CAD, item.AUD, item.GBP), "}");
                    leadSourceColumn.NSeries.Add(totalRowData, true);

                    string categoryData = string.Concat("{", string.Join(", ", str), "}");
                    leadSourceColumn.NSeries.CategoryData = categoryData;

                    leadSourceColumn.NSeries[count].Name = string.Format("Tháng {0}/{1}", item.Month, item.Year);

                    // Set the 2nd series fill color.
                    leadSourceColumn.NSeries[count].Area.ForegroundColor = Color.Orange;
                    leadSourceColumn.NSeries[count].Area.Formatting = FormattingType.Custom;

                    if (count.Equals(1))
                    {
                        // Set the 1st series fill color.
                        leadSourceColumn.NSeries[count].Area.ForegroundColor = Color.Green;
                        leadSourceColumn.NSeries[count].Area.Formatting = FormattingType.Custom;
                    }

                    if (count.Equals(2))
                    {
                        // Set the 1st series fill color.
                        leadSourceColumn.NSeries[count].Area.ForegroundColor = Color.Blue;
                        leadSourceColumn.NSeries[count].Area.Formatting = FormattingType.Custom;
                    }
                    count++;
                }

                // Set plot area formatting as none and hide its border.
                leadSourceColumn.PlotArea.Area.FillFormat.FillType = FillType.None;
                leadSourceColumn.PlotArea.Border.IsVisible = false;

                // Set value axis major tick mark as none and hide axis line. 
                // Also set the color of value axis major grid lines.
                leadSourceColumn.ValueAxis.MajorTickMark = TickMarkType.None;
                leadSourceColumn.ValueAxis.AxisLine.IsVisible = false;
                leadSourceColumn.ValueAxis.MajorGridLines.Color = Color.FromArgb(217, 217, 217);

                // Set border
                Style style = new CellsFactory().CreateStyle();
                style.SetBorder(BorderType.TopBorder, CellBorderType.Thin, Color.Black);
                style.SetBorder(BorderType.BottomBorder, CellBorderType.Thin, Color.Black);
                style.SetBorder(BorderType.LeftBorder, CellBorderType.Thin, Color.Black);
                style.SetBorder(BorderType.RightBorder, CellBorderType.Thin, Color.Black);

                if (dataTableConvert.Rows.Count > 0)
                {
                    int stepRow = 0;
                    // total row = row start + số row hiện có. 2 là dòng tổng
                    int totalRow = dataTableConvert.Rows.Count + 7 + 2 + dataTableConvert.Rows.Count + 1;
                    // Số dòng của row
                    for (int a = 17; a < totalRow; a++)
                    {
                        int stepColumn = 0;
                        // Số cột trong báo cáo cần hiển thị
                        // Tổng số cột hiển thị = Số cột hiển thị bắt đầu + tổng số cột cần hiển thị
                        int totalCol = 15 + 8;
                        for (int b = 15; b < totalCol; b++)
                        {
                            // Giá trị của value trong table
                            string valueOfTable = dataTableConvert.Rows[stepRow][stepColumn].ToString();

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
                            if (b.Equals(totalCol - 1) || b.Equals(totalCol - 2)
                                || b.Equals(totalCol - 3) || b.Equals(totalCol - 4))
                            {
                                sheetReport.Cells[a, b].PutValue(valueOfTable, true, true);
                                style.Font.IsBold = true;
                                sheetReport.Cells[a, b].SetStyle(style);
                            }

                            // Trường hợp thuộc dòng cuối
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
                bool check = true;
                foreach (ReportForTotalMoneyType item in listReportDataPercentClone)
                {
                    item.ReportID = string.Format("Tháng {0}/{1}", item.Month, item.Year);
                }

                ReportForTotalMoneyType dataPieYear = null;
                ReportForTotalMoneyType dataPieLastMonth = null;
                ReportForTotalMoneyType dataPieLastYear = null;
                // Data report năm hiện tại nhập vào
                dataPieYear = listReportDataPercentClone.Find(x => x.Year == year.ToString());
                // Data report năm hiện tại, tháng hiện tại nhập vào
                dataPieLastMonth = listReportDataPercentClone.Find(x => x.Year == year.ToString() && x.Month == (month - 1).ToString());

                if ( month == 1)
                {
                    dataPieLastMonth = listReportDataPercentClone.Find(x => x.Year == (year -1).ToString() && x.Month == "12");
                }
                // Data report năm ngoái so với năm hiện tại nhập vào
                dataPieLastYear = listReportDataPercentClone.Find(x => x.Year == (year - 1).ToString());

                if (dataPieYear != null)
                {
                    //Add Pie Chart
                    Aspose.Cells.Charts.Chart leadSourcePie;
                    int chartIndex = sheetReport.Charts.Add(ChartType.Pie3D, 34, 0, 49, 6);
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
                    string totalRowData = string.Concat("{", string.Format("{0}, {1}, {2}, {3}, {4}, {5}", dataPieYear.VND, dataPieYear.USD, dataPieYear.EUR, dataPieYear.CAD, dataPieYear.AUD, dataPieYear.GBP), "}");
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
                    int chartIndex = sheetReport.Charts.Add(ChartType.Pie3D, 34, 8, 49, 14);
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
                    string totalRowData = string.Concat("{", string.Format("{0}, {1}, {2}, {3}, {4}, {5}", dataPieYear.VND, dataPieYear.USD, dataPieYear.EUR, dataPieYear.CAD, dataPieYear.AUD, dataPieYear.GBP), "}");
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
                    int chartIndex = sheetReport.Charts.Add(ChartType.Pie3D, 51, 0, 65, 6);
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
                    string totalRowData = string.Concat("{", string.Format("{0}, {1}, {2}, {3}, {4}, {5}", dataPieYear.VND, dataPieYear.USD, dataPieYear.EUR, dataPieYear.CAD, dataPieYear.AUD, dataPieYear.GBP), "}");
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
                double VNDCompare = listReportDataPercentClone[0].VND - listReportDataPercentClone[1].VND;
                double USDCompare = listReportDataPercentClone[0].USD - listReportDataPercentClone[1].USD;
                double EURCompare = listReportDataPercentClone[0].EUR - listReportDataPercentClone[1].EUR;
                double CADCompare = listReportDataPercentClone[0].CAD - listReportDataPercentClone[1].CAD;
                double AUDCompare = listReportDataPercentClone[0].AUD - listReportDataPercentClone[1].AUD;
                double GBPCompare = listReportDataPercentClone[0].GBP - listReportDataPercentClone[1].GBP;

                // tháng cùng kì năm trước
                double VNDCompareMonthLastYear = listReportDataPercentClone[0].VND - listReportDataPercentClone[2].VND;
                double USDCompareMonthLastYear = listReportDataPercentClone[0].USD - listReportDataPercentClone[2].USD;
                double EURCompareMonthLastYear = listReportDataPercentClone[0].EUR - listReportDataPercentClone[2].EUR;
                double CADCompareMonthLastYear = listReportDataPercentClone[0].CAD - listReportDataPercentClone[2].CAD;
                double AUDCompareMonthLastYear = listReportDataPercentClone[0].AUD - listReportDataPercentClone[2].AUD;
                double GBPCompareMonthLastYear = listReportDataPercentClone[0].GBP - listReportDataPercentClone[2].GBP;

                // Tạo các cột cho datatable
                DataTable dataTablePie = new DataTable();

                dataTablePie.Columns.Add("ReportID", typeof(String));
                dataTablePie.Columns.Add("AccumulateID1", typeof(double));
                dataTablePie.Columns.Add("AccumulateID2", typeof(double));
                dataTablePie.Columns.Add("AccumulateID3", typeof(double));
                dataTablePie.Columns.Add("CompareToMonthPercent", typeof(double));
                dataTablePie.Columns.Add("CompareToMonthLastYearPercent", typeof(double));

                // add row vào table
                dataTablePie.Rows.Add(str[0], listReportDataPercentClone[0].VND, listReportDataPercentClone[1].VND
                    , listReportDataPercentClone[2].VND
                    , listReportDataPercentClone[1].VND == 0 ? listReportDataPercentClone[1].VND : Math.Round(VNDCompare / listReportDataPercentClone[1].VND * 100, 2, MidpointRounding.ToEven)
                    , listReportDataPercentClone[2].VND == 0 ? listReportDataPercentClone[2].VND : Math.Round(VNDCompareMonthLastYear / listReportDataPercentClone[2].VND * 100, 2, MidpointRounding.ToEven));

                dataTablePie.Rows.Add(str[1], listReportDataPercentClone[0].USD, listReportDataPercentClone[1].USD
                    , listReportDataPercentClone[2].USD
                    , listReportDataPercentClone[1].USD == 0 ? listReportDataPercentClone[1].USD : Math.Round(USDCompare / listReportDataPercentClone[1].USD * 100, 2, MidpointRounding.ToEven)
                    , listReportDataPercentClone[2].USD == 0 ? listReportDataPercentClone[2].USD : Math.Round(USDCompareMonthLastYear / listReportDataPercentClone[2].USD * 100, 2, MidpointRounding.ToEven));

                dataTablePie.Rows.Add(str[2], listReportDataPercentClone[0].EUR, listReportDataPercentClone[1].EUR
                    , listReportDataPercentClone[2].EUR
                    , listReportDataPercentClone[1].EUR == 0 ? listReportDataPercentClone[1].EUR : Math.Round(EURCompare / listReportDataPercentClone[1].EUR * 100, 2, MidpointRounding.ToEven)
                    , listReportDataPercentClone[2].EUR == 0 ? listReportDataPercentClone[2].EUR : Math.Round(EURCompareMonthLastYear / listReportDataPercentClone[2].EUR * 100, 2, MidpointRounding.ToEven));

                dataTablePie.Rows.Add(str[3], listReportDataPercentClone[0].CAD, listReportDataPercentClone[1].CAD
                    , listReportDataPercentClone[2].CAD
                    , listReportDataPercentClone[1].CAD == 0 ? listReportDataPercentClone[1].CAD : Math.Round(CADCompare / listReportDataPercentClone[1].CAD * 100, 2, MidpointRounding.ToEven)
                    , listReportDataPercentClone[2].CAD == 0 ? listReportDataPercentClone[2].CAD : Math.Round(CADCompareMonthLastYear / listReportDataPercentClone[2].CAD * 100, 2, MidpointRounding.ToEven));

                dataTablePie.Rows.Add(str[4], listReportDataPercentClone[0].AUD, listReportDataPercentClone[1].AUD
                    , listReportDataPercentClone[2].AUD
                    , listReportDataPercentClone[1].AUD == 0 ? listReportDataPercentClone[1].AUD : Math.Round(AUDCompare / listReportDataPercentClone[1].AUD * 100, 2, MidpointRounding.ToEven)
                    , listReportDataPercentClone[2].AUD == 0 ? listReportDataPercentClone[2].AUD : Math.Round(AUDCompareMonthLastYear / listReportDataPercentClone[2].AUD * 100, 2, MidpointRounding.ToEven));

                dataTablePie.Rows.Add(str[5], listReportDataPercentClone[0].GBP, listReportDataPercentClone[1].GBP
                    , listReportDataPercentClone[2].GBP
                    , listReportDataPercentClone[1].GBP == 0 ? listReportDataPercentClone[1].GBP : Math.Round(GBPCompare / listReportDataPercentClone[1].GBP * 100, 2, MidpointRounding.ToEven)
                    , listReportDataPercentClone[2].GBP == 0 ? listReportDataPercentClone[2].GBP : Math.Round(GBPCompareMonthLastYear / listReportDataPercentClone[2].GBP * 100, 2, MidpointRounding.ToEven));

                // Add dòng tổng
                DataRow row = dataTablePie.NewRow();
                row["ReportID"] = "Tổng";
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

                            // dòng tổng cộng
                            // Cột tổng cộng
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
                row["VND"] = (double)dr["VND"];
                row["USD"] = (double)dr["USD"];
                row["EUR"] = (double)dr["EUR"];
                row["CAD"] = (double)dr["CAD"];
                row["AUD"] = (double)dr["AUD"];
                row["GBP"] = (double)dr["GBP"];
                row["TongDS"] = (double)dr["TongDS"];
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
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
    public class ReportHSTypeExcelController : Controller
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

        // GET: Admin/RepportHSTypeExcel
        public ActionResult Index()
        {
            return View();
        }


        public ActionResult CreateExcel(DateTime fromDate, DateTime toDate, string typeID, string reportTypeID)
        {
            WorkbookDesigner designer = new WorkbookDesigner();
            string templatePath = "~/Content/Report/ReportHS/ReportHSForDMY.xlsx";
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
                case "1":
                    typeReport = "Chi tiết - Theo ngày";
                    break;
                // Theo tháng
                case "2":
                    typeReport = "Chi tiết - Theo Tháng";
                    break;
                // Theo năm
                default:
                    typeReport = "Chi tiết - Theo Năm";
                    break;
            }

            CreateTitle("A2", "U2", sheetReport, typeReport, 14);

            switch (typeID)
            {
                // Theo ngày
                case "1":
                    // Set from day and to day
                    sheetReport.Cells["M4"].PutValue(fromDate.ToString("dd/MM/yyyy"));
                    sheetReport.Cells["Q4"].PutValue(toDate.ToString("dd/MM/yyyy"));
                    break;
                // Theo tháng
                case "2":

                    // Set from day and to day
                    sheetReport.Cells["M4"].PutValue(fromDate.ToString("MM/yyyy"));
                    sheetReport.Cells["Q4"].PutValue(toDate.ToString("MM/yyyy"));
                    break;
                // Theo năm
                default:

                    // Set from day and to day
                    sheetReport.Cells["M4"].PutValue(fromDate.Year.ToString());
                    sheetReport.Cells["Q4"].PutValue(toDate.Year.ToString());
                    break;
            }

            // Create Chart Line
            //Chart reference
            Aspose.Cells.Charts.Chart leadSourceLine;
            //Add Pie Chart
            int chartIndex = sheetReport.Charts.Add(ChartType.Line, 6, 0, 30, 13);
            leadSourceLine = sheetReport.Charts[chartIndex];

            // Canh hiển thị CategoryAxis nghiên phù hợp
            leadSourceLine.CategoryAxis.TickLabels.RotationAngle = 45;

            //Chart title
            leadSourceLine.Title.Text = "Hồ sơ chi trả theo từng dịch vụ";
            leadSourceLine.Title.Font.Color = Color.Silver;


            List<Report> listReportData = new List<Report>();
            int typeIDForDate = int.Parse(typeID);
            // Danh sách ngày
            switch (typeID)
            {
                // Theo ngày
                case "1":
                    listReportData = new HSReportBL().SearchDay(fromDate, toDate, reportTypeID);
                    break;
                // Theo tháng
                case "2":
                    listReportData = new HSReportBL().SearchMonth(fromDate, toDate, reportTypeID);
                    break;
                // Theo năm
                default:
                    listReportData = new HSReportBL().SearchYear(fromDate, toDate, reportTypeID);
                    break;
            }

            DataTable dataTable = new DataTable();

            if (listReportData.Count > 0)
            {
                foreach (Report item in listReportData)
                {
                    switch (typeID)
                    {
                        // Theo ngày
                        case "1":
                            item.ReportID = string.Concat("Ngày ", item.CreatedDate.Day, "/", item.CreatedDate.Month);
                            break;
                        // Theo tháng
                        case "2":
                            item.ReportID = string.Concat("Tháng ", item.Month, "/", item.Year);
                            break;
                        // Theo năm
                        default:
                            item.ReportID = string.Concat("Năm ", item.Year);
                            break;
                    }
                    item.TongDS = item.DSChiNha + item.DSChiQuay + item.DSCK;
                    item.Type = 0;
                }

                // Add dòng tổng vào list danh sách
                Report report = new Report()
                {
                    ReportID = "Tổng",
                    DSChiNha = listReportData.Sum(item => item.DSChiNha),
                    DSChiQuay = listReportData.Sum(item => item.DSChiQuay),
                    DSCK = listReportData.Sum(item => item.DSCK),
                    TongDS = listReportData.Sum(item => item.TongDS)
                };
                listReportData.Add(report);

                // Tạo các col cho table
                dataTable = CreateDataTableFormart(typeID);

                // Danh sách dataSet của báo cáo ngày
                DataSet dataReport = ConvertListObjectToDataSet(listReportData);

                // Đổ data vào datatble mới
                FillData(dataReport.Tables[0], dataTable, typeID);

                // Adding SeriesCollection (chart data source) to the chart ranging from "A1" cell to "B3"
                string totalRowData = string.Format("Q8:S{0}", listReportData.Count + 8 - 2);
                leadSourceLine.NSeries.Add(totalRowData, true);

                // Set the category data covering the range A2:A5.
                // Tổng số dòng cần hiển thị là số dòng hiện tại  + số dòng trong listData -2 (trong đó 1 là dòng tổng cuối cùng)
                string categoryData = string.Format("P8:P{0}", listReportData.Count + 8 - 2);
                leadSourceLine.NSeries.CategoryData = categoryData;

                // Set the names of the chart series taken from cells.
                leadSourceLine.NSeries[0].Name = "=Q7";
                leadSourceLine.NSeries[1].Name = "=R7";
                leadSourceLine.NSeries[2].Name = "=S7";
                //leadSourceLine.NSeries[3].Name = "=T7";

                // Set the 1st series fill color.
                leadSourceLine.NSeries[0].Border.Color = Color.Orange;
                leadSourceLine.NSeries[0].Area.Formatting = FormattingType.Custom;

                // Set the 2nd series fill color.
                leadSourceLine.NSeries[1].Border.Color = Color.Green;
                leadSourceLine.NSeries[1].Area.Formatting = FormattingType.Custom;

                // Set the 3rd series fill color.
                leadSourceLine.NSeries[2].Border.Color = Color.Blue;
                leadSourceLine.NSeries[2].Area.Formatting = FormattingType.Custom;

                //// Set the 4rd series fill color.
                //leadSourceLine.NSeries[3].Border.Color = Color.Purple;
                //leadSourceLine.NSeries[3].Area.Formatting = FormattingType.Custom;


                // Set plot area formatting as none and hide its border.
                leadSourceLine.PlotArea.Area.FillFormat.FillType = FillType.None;
                leadSourceLine.PlotArea.Border.IsVisible = false;

                // Set value axis major tick mark as none and hide axis line. 
                // Also set the color of value axis major grid lines.
                leadSourceLine.ValueAxis.MajorTickMark = TickMarkType.None;
                leadSourceLine.ValueAxis.AxisLine.IsVisible = false;
                leadSourceLine.ValueAxis.MajorGridLines.Color = Color.FromArgb(217, 217, 217);
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
                    int totalCol = 15 + 5;
                    for (int b = 15; b < totalCol; b++)
                    {
                        // Giá trị của value trong table
                        string valueOfTable = dataTable.Rows[stepRow][stepColumn].ToString();

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
                case "1":
                    return ExportReport("ReportDay", designer);
                    break;
                // Theo tháng
                case "2":
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
        public ActionResult CreateExcelGradationCompare(string gradationID, int year, string reportTypeID)
        {
            WorkbookDesigner designer = new WorkbookDesigner();
            string templatePath = "~/Content/Report/ReportHS/ReportHSCompare.xlsx";
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


            // Create Chart Line
            //Chart reference
            Aspose.Cells.Charts.Chart leadSourceLine;
            //Add Pie Chart
            int chartIndex = sheetReport.Charts.Add(ChartType.Column3DClustered, 6, 0, 30, 13);
            leadSourceLine = sheetReport.Charts[chartIndex];


            //Chart title
            leadSourceLine.Title.Text = "Doanh số theo loại hình dịch vụ";
            leadSourceLine.Title.Font.Color = Color.Silver;

            // Set width cho column
            sheetReport.Cells.SetColumnWidthPixel(15, 250);

            // Danh sách doanh số chi trả theo loại hình dịch vụ
            List<Report> listReportData = new HSReportBL().SearchGradationCompare(year, int.Parse(gradationID), reportTypeID);
            // clone Object
            List<Report> listReportDataClone = new List<Report>(listReportData);

            // Danh sách tỉ trọng chi trả theo loại hình dịch vụ
            List<Report> listReportDataPercent = new HSReportBL().ListDataGradationComparePercent(int.Parse(gradationID), year, reportTypeID);
            // clone Object
            List<Report> listReportDataPercentClone = new List<Report>(listReportDataPercent);

            DataTable dataTable = new DataTable();

            // Theo doanh số chi trả loại hình dịch vụ
            if (listReportData.Count.Equals(2))
            {
                bool check = true;
                foreach (Report item in listReportData)
                {
                    item.GradationID = string.Concat("Lũy kế ", text, " ", year);
                    if (!check)
                    {
                        item.GradationID = string.Concat("Lũy kế ", text, " ", year - 1);
                    }
                    item.TongDS = item.DSChiNha + item.DSChiQuay + item.DSCK;
                    item.Type = 0;
                    // Set lại giá trị cho check để lấy giá trị của năm trước
                    check = false;
                }

                // Object báo cáo tăng giảm so với cùng kỳ (+/-)
                // ds so với tháng trước
                double dsChiQuaySum = (listReportData[0].DSChiQuay - listReportData[1].DSChiQuay);
                double dsChiNhaSum = (listReportData[0].DSChiNha - listReportData[1].DSChiNha);
                double dsCKSum = (listReportData[0].DSCK - listReportData[1].DSCK);
                double tongDSSum = (listReportData[0].TongDS - listReportData[1].TongDS);

                // Object báo cáo tăng giảm so với cùng kỳ (%)
                Report dataDifferencePercent = new Report()
                {
                    GradationID = string.Format("Tăng giảm so với cùng kì {0} (%)", year - 1),
                    DSChiQuay = listReportData[1].DSChiQuay == 0 ? 0 : Math.Round(dsChiQuaySum / listReportData[1].DSChiQuay * 100, 2, MidpointRounding.ToEven),
                    DSChiNha = listReportData[1].DSChiNha == 0 ? 0 : Math.Round(dsChiNhaSum / listReportData[1].DSChiNha * 100, 2, MidpointRounding.ToEven),
                    DSCK = listReportData[1].DSCK == 0 ? 0 : Math.Round(dsCKSum / listReportData[1].DSCK * 100, 2, MidpointRounding.ToEven),
                    TongDS = listReportData[1].TongDS == 0 ? 0 : Math.Round(tongDSSum / listReportData[1].TongDS * 100, 2, MidpointRounding.ToEven),
                };

                listReportData.Add(dataDifferencePercent);

                Report dataDifference = new Report()
                {
                    GradationID = string.Format("Tăng giảm so với cùng kì {0} (+/-)", year - 1),
                    DSChiQuay = Math.Round(dsChiQuaySum, 2, MidpointRounding.ToEven),
                    DSChiNha = Math.Round(dsChiNhaSum, 2, MidpointRounding.ToEven),
                    DSCK = Math.Round(dsCKSum, 2, MidpointRounding.ToEven),
                    TongDS = Math.Round(tongDSSum, 2, MidpointRounding.ToEven),
                };

                listReportData.Add(dataDifference);

                dataTable = new DataTable();
                dataTable.Columns.Add("GradationID", typeof(string));
                dataTable.Columns.Add("DSChiQuay", typeof(double));
                dataTable.Columns.Add("DSChiNha", typeof(double));
                dataTable.Columns.Add("DSCK", typeof(double));
                dataTable.Columns.Add("TongDS", typeof(double));
                

                // Danh sách dataSet của báo cáo ngày
                DataSet dataReport = new DongA.Core.DongAExcel().ConvertListObjectToDataSet(listReportData);

                // Đổ data vào datatble mới
                FillDataGradationCompare(dataReport.Tables[0], dataTable);

                int count = 0;
                bool checkColor = true;
                foreach (Report item in listReportDataClone)
                {
                    //string totalRowData = string.Concat("{", string.Format("{0}, {1}, {2}, {3}", item.DSChiQuay, item.DSChiNha, item.DSCK, item.TongDS), "}");
                    string totalRowData = string.Concat("{", string.Format("{0}, {1}, {2}", item.DSChiQuay, item.DSChiNha, item.DSCK), "}");
                    leadSourceLine.NSeries.Add(totalRowData, true);

                    //string categoryData = "{Doanh số chi quầy, Doanh số chi nhà, Doanh số chi chuyển khoản, Tổng doanh số}";
                    string categoryData = "{Doanh số chi quầy, Doanh số chi nhà, Doanh số chi chuyển khoản}";
                    leadSourceLine.NSeries.CategoryData = categoryData;

                    leadSourceLine.NSeries[count].Name = string.Concat("Lũy kế ", text, " ", item.Year);

                    // Set the 2nd series fill color.
                    leadSourceLine.NSeries[count].Area.ForegroundColor = Color.Orange;
                    leadSourceLine.NSeries[count].Area.Formatting = FormattingType.Custom;

                    if (!checkColor)
                    {
                        // Set the 1st series fill color.
                        leadSourceLine.NSeries[count].Area.ForegroundColor = Color.Blue;
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

                            if (a == 9)
                            {
                                // set style cho number
                                style.Custom = "0.00";
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

            // Biểu đồ Doanh số theo loại hình chi trả của năm hiện tại
            if (listReportDataPercentClone.Count.Equals(2))
            {
                bool check = true;
                foreach (Report item in listReportDataPercentClone)
                {
                    item.GradationID = string.Concat("Lũy kế ", text, " ", year);
                    if (!check)
                    {
                        item.GradationID = string.Concat("Lũy kế ", text, " ", year - 1);
                    }
                    //item.TongDS = item.DSChiNha + item.DSChiQuay + item.DSCK;
                    //item.Type = 0;
                    // Set lại giá trị cho check để lấy giá trị của năm trước
                    check = false;
                }

                Report dataPieYear = null;
                Report dataPieLastYear = null;
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
                    string totalRowData = string.Concat("{", string.Format("{0}, {1}, {2}", dataPieYear.DSChiQuay, dataPieYear.DSChiNha, dataPieYear.DSCK), "}");
                    leadSourcePie.NSeries.Add(totalRowData, true);

                    string categoryData = "{chi quầy, chi nhà, chi chuyển khoản}";
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
                    string totalRowData = string.Concat("{", string.Format("{0}, {1}, {2}", dataPieLastYear.DSChiQuay, dataPieLastYear.DSChiNha, dataPieLastYear.DSCK), "}");
                    leadSourcePieLasYear.NSeries.Add(totalRowData, true);

                    string categoryData = "{chi quầy, chi nhà, chi chuyển khoản}";
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
                double dsChiQuaySum = (listReportDataPercentClone[0].DSChiQuay - listReportDataPercentClone[1].DSChiQuay);
                double dsChiNhaSum = (listReportDataPercentClone[0].DSChiNha - listReportDataPercentClone[1].DSChiNha);
                double dsCKSum = (listReportDataPercentClone[0].DSCK - listReportDataPercentClone[1].DSCK);
                double tongDSSum = (listReportDataPercentClone[0].TongDS - listReportDataPercentClone[1].TongDS);

                // Object báo cáo tăng giảm so với cùng kỳ (%)
                Report dataDifferencePercent = new Report()
                {
                    GradationID = string.Format("Tăng giảm so với cùng kì {0} (%)", year - 1),
                    DSChiQuay = listReportDataPercentClone[1].DSChiQuay == 0 ? 0 : Math.Round(dsChiQuaySum / listReportDataPercentClone[1].DSChiQuay * 100, 2, MidpointRounding.ToEven),
                    DSChiNha = listReportDataPercentClone[1].DSChiNha == 0 ? 0 : Math.Round(dsChiNhaSum / listReportDataPercentClone[1].DSChiNha * 100, 2, MidpointRounding.ToEven),
                    DSCK = listReportDataPercentClone[1].DSCK == 0 ? 0 : Math.Round(dsCKSum / listReportDataPercentClone[1].DSCK * 100, 2, MidpointRounding.ToEven),
                    TongDS = 0,
                };

                listReportDataPercentClone.Add(dataDifferencePercent);

                DataTable dataTablePie = new DataTable();
                // Add item to datatable
                dataTablePie = CreateDataTableFormartGradation();

                // Danh sách dataSet của báo cáo ngày
                DataSet dataReportPie = new DongA.Core.DongAExcel().ConvertListObjectToDataSet(listReportDataPercentClone);
                // Đổ data vào datatble mới
                FillDataGradationCompare(dataReportPie.Tables[0], dataTablePie);
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
                        int totalCol = 15 + 5;
                        for (int b = 15; b < totalCol; b++)
                        {
                            // Giá trị của value trong table
                            string valueOfTable = dataTablePie.Rows[stepRow][stepColumn].ToString();


                            // Tô màu cho các dòng có giá trị tăng giảm
                            if (b >= 16 && a >= 37)
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
        public ActionResult CreateExcelCompareMonth(int year, int month, string reportTypeID)
        {
            WorkbookDesigner designer = new WorkbookDesigner();
            string templatePath = "~/Content/Report/ReportHS/ReportHSCompare.xlsx";
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


            // Create Chart Line
            //Chart reference
            Aspose.Cells.Charts.Chart leadSourceLine;
            //Add Pie Chart
            int chartIndex = sheetReport.Charts.Add(ChartType.Column3DClustered, 6, 0, 30, 13);
            leadSourceLine = sheetReport.Charts[chartIndex];
            //Chart title
            leadSourceLine.Title.Text = string.Format("Doanh số theo loại hình dịch vụ tháng {0}/{1} so với tháng trước và so với cùng kì năm trước", month, year);
            leadSourceLine.Title.Font.Color = Color.Silver;

            // Set width cho column
            sheetReport.Cells.SetColumnWidthPixel(15, 220);

            // Danh sách doanh số chi trả theo loại hình dịch vụ
            List<Report> listReportData = new HSReportBL().SearchMonthCompareGrid(year, month, reportTypeID);
            // clone Object
            List<Report> listReportDataClone = new List<Report>(listReportData);

            List<Report> listReportDataPercent = new HSReportBL().ListDataLastMonthCompareProportionPercent(year, month, reportTypeID);
            // clone Object
            List<Report> listReportDataPercentClone = new List<Report>(listReportDataPercent);

            DataTable dataTable = new DataTable();

            // Theo doanh số chi trả loại hình dịch vụ
            if (listReportData.Count.Equals(3))
            {
                foreach (Report item in listReportData)
                {
                    item.GradationID = string.Format("Tháng {0}/{1}", item.Month, item.Year);
                    item.TongDS = item.DSChiQuay + item.DSChiNha + item.DSCK;
                }

                // ds so với tháng trước
                double dsChiQuaySum = (listReportData[0].DSChiQuay - listReportData[1].DSChiQuay);
                double dsChiNhaSum = (listReportData[0].DSChiNha - listReportData[1].DSChiNha);
                double dsCKSum = (listReportData[0].DSCK - listReportData[1].DSCK);
                double tongDSSum = (listReportData[0].TongDS - listReportData[1].TongDS);

                // ds so với cùng kì năm trước
                double dsChiQuaySumlastYear = (listReportData[0].DSChiQuay - listReportData[2].DSChiQuay);
                double dsChiNhaSumlastYear = (listReportData[0].DSChiNha - listReportData[2].DSChiNha);
                double dsCKSumlastYear = (listReportData[0].DSCK - listReportData[2].DSCK);
                double tongDSSumlastYear = (listReportData[0].TongDS - listReportData[2].TongDS);

                List<Report> listReport = new List<Report>()
                {
                    // So sánh Tháng trước theo %
                    new Report()
                    {
                        GradationID = "Tăng giảm so với tháng trước (%)",
                        DSChiQuay = listReportData[1].DSChiQuay == 0 ? 0 : Math.Round(dsChiQuaySum / listReportData[1].DSChiQuay *100, 2, MidpointRounding.ToEven),
                        DSChiNha = listReportData[1].DSChiNha == 0 ? 0 : Math.Round(dsChiNhaSum / listReportData[1].DSChiNha *100, 2, MidpointRounding.ToEven),
                        DSCK = listReportData[1].DSCK == 0 ? 0 : Math.Round(dsCKSum / listReportData[1].DSCK *100, 2, MidpointRounding.ToEven),
                        TongDS = listReportData[1].TongDS == 0 ? 0 :  Math.Round(tongDSSum / listReportData[1].TongDS *100, 2, MidpointRounding.ToEven)
                    },

                    // So sánh Tháng trước theo tăng giảm
                    new Report()
                    {
                        GradationID = "Tăng giảm so với tháng trước (+/-)",
                        DSChiQuay = Math.Round(dsChiQuaySum, 2, MidpointRounding.ToEven),
                        DSChiNha = Math.Round(dsChiNhaSum, 2, MidpointRounding.ToEven),
                        DSCK = Math.Round(dsCKSum, 2,MidpointRounding.ToEven),
                        TongDS = Math.Round(tongDSSum, 2,MidpointRounding.ToEven),
                    },

                    // So sánh cùng kỳ năm trước %
                    new Report()
                    {
                        GradationID = "Tăng giảm so với cùng kỳ năm trước (%)",
                        DSChiQuay = listReportData[2].DSChiQuay == 0 ? 0 : Math.Round(dsChiQuaySumlastYear / listReportData[2].DSChiQuay *100, 2, MidpointRounding.ToEven),
                        DSChiNha = listReportData[2].DSChiNha == 0 ? 0 : Math.Round(dsChiNhaSumlastYear / listReportData[2].DSChiNha *100, 2, MidpointRounding.ToEven),
                        DSCK = listReportData[2].DSCK == 0 ? 0 : Math.Round(dsCKSumlastYear / listReportData[2].DSCK *100, 2, MidpointRounding.ToEven),
                        TongDS = listReportData[2].TongDS == 0 ? 0 : Math.Round(tongDSSumlastYear / listReportData[2].TongDS *100, 2, MidpointRounding.ToEven)
                    },

                    // So sánh cùng kỳ năm trước theo tăng giảm
                    new Report()
                    {
                        GradationID = "Tăng giảm so với cùng kỳ năm trước (+/-)",
                        DSChiQuay = Math.Round(dsChiQuaySumlastYear, 2, MidpointRounding.ToEven),
                        DSChiNha = Math.Round(dsChiNhaSumlastYear, 2, MidpointRounding.ToEven),
                        DSCK = Math.Round(dsCKSumlastYear, 2,MidpointRounding.ToEven),
                        TongDS = Math.Round(tongDSSumlastYear, 2,MidpointRounding.ToEven),
                    },
                };

                // Merger
                listReportData.AddRange(listReport);

                // Add item to datatable
                dataTable = CreateDataTableFormartGradation();

                // Danh sách dataSet của báo cáo ngày
                DataSet dataReport = new DongA.Core.DongAExcel().ConvertListObjectToDataSet(listReportData);

                // Đổ data vào datatble mới
                FillDataGradationCompare(dataReport.Tables[0], dataTable);

                int count = 0;
                foreach (Report item in listReportDataClone)
                {
                    //string totalRowData = string.Concat("{", string.Format("{0}, {1}, {2}, {3}", item.DSChiQuay, item.DSChiNha, item.DSCK, item.TongDS), "}");
                    string totalRowData = string.Concat("{", string.Format("{0}, {1}, {2}", item.DSChiQuay, item.DSChiNha, item.DSCK), "}");
                    leadSourceLine.NSeries.Add(totalRowData, true);

                    //string categoryData = "{Doanh số chi quầy, Doanh số chi nhà, Doanh số chi chuyển khoản, Tổng doanh số}";
                    string categoryData = "{Doanh số chi quầy, Doanh số chi nhà, Doanh số chi chuyển khoản}";
                    leadSourceLine.NSeries.CategoryData = categoryData;

                    leadSourceLine.NSeries[count].Name = string.Format("Tháng {0}/{1}", item.Month, item.Year);

                    // Set the 2nd series fill color.
                    leadSourceLine.NSeries[count].Area.ForegroundColor = Color.Orange;
                    leadSourceLine.NSeries[count].Area.Formatting = FormattingType.Custom;

                    if (count.Equals(1))
                    {
                        // Set the 1st series fill color.
                        leadSourceLine.NSeries[count].Area.ForegroundColor = Color.Blue;
                        leadSourceLine.NSeries[count].Area.Formatting = FormattingType.Custom;
                    }

                    if (count.Equals(2))
                    {
                        // Set the 1st series fill color.
                        leadSourceLine.NSeries[count].Area.ForegroundColor = Color.Green;
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
                        int totalCol = 15 + 5;
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
                            if (a > 9)
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
                foreach (Report item in listReportDataPercentClone)
                {
                    item.GradationID = string.Format("Tháng {0}/{1}", item.Month, item.Year);
                    item.TongDS = 100;
                }

                Report dataPieYear = null;
                Report dataPieLastMonth = null;
                Report dataPieLastYear = null;
                // Data report năm hiện tại nhập vào
                dataPieYear = listReportDataPercentClone.Find(x => x.Year == year.ToString());
                // Data report năm hiện tại, tháng hiện tại nhập vào
                dataPieLastMonth = listReportDataPercentClone.Find(x => x.Year == year.ToString() && x.Month == (month - 1).ToString());
                // Data report năm ngoái so với năm hiện tại nhập vào
                dataPieLastYear = listReportDataPercentClone.Find(x => x.Year == (year - 1).ToString());

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
                    leadSourcePie.Title.Text = dataPieYear.GradationID;
                    leadSourcePie.Title.Font.Color = Color.Silver;
                    leadSourcePie.Title.Font.IsBold = true;
                    leadSourcePie.Title.Font.Size = 12;

                    // Set properties of nseries
                    string totalRowData = string.Concat("{", string.Format("{0}, {1}, {2}", dataPieYear.DSChiQuay, dataPieYear.DSChiNha, dataPieYear.DSCK), "}");
                    leadSourcePie.NSeries.Add(totalRowData, true);

                    string categoryData = "{chi quầy, chi nhà, chi chuyển khoản}";
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
                    leadSourcePieLasYear.Title.Text = dataPieLastYear.GradationID;
                    leadSourcePieLasYear.Title.Font.Color = Color.Silver;
                    leadSourcePieLasYear.Title.Font.IsBold = true;
                    leadSourcePieLasYear.Title.Font.Size = 12;

                    // Set properties of nseries
                    string totalRowData = string.Concat("{", string.Format("{0}, {1}, {2}", dataPieLastYear.DSChiQuay, dataPieLastYear.DSChiNha, dataPieLastYear.DSCK), "}");
                    leadSourcePieLasYear.NSeries.Add(totalRowData, true);

                    string categoryData = "{chi quầy, chi nhà, chi chuyển khoản}";
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
                    leadSourcePie.Title.Text = dataPieLastMonth.GradationID;
                    leadSourcePie.Title.Font.Color = Color.Silver;
                    leadSourcePie.Title.Font.IsBold = true;
                    leadSourcePie.Title.Font.Size = 12;

                    // Set properties of nseries
                    string totalRowData = string.Concat("{", string.Format("{0}, {1}, {2}", dataPieLastMonth.DSChiQuay, dataPieLastMonth.DSChiNha, dataPieLastMonth.DSCK), "}");
                    leadSourcePie.NSeries.Add(totalRowData, true);

                    string categoryData = "{chi quầy, chi nhà, chi chuyển khoản}";
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
                double dsChiQuaySum = (listReportDataPercentClone[0].DSChiQuay - listReportDataPercentClone[1].DSChiQuay);
                double dsChiNhaSum = (listReportDataPercentClone[0].DSChiNha - listReportDataPercentClone[1].DSChiNha);
                double dsCKSum = (listReportDataPercentClone[0].DSCK - listReportDataPercentClone[1].DSCK);
                double tongDSSum = (listReportDataPercentClone[0].TongDS - listReportDataPercentClone[1].TongDS);

                // Object báo cáo tăng giảm so với tháng trước (%)
                Report dataDifferencePercentLastMonth = new Report()
                {
                    GradationID = "Tăng giảm so với tháng trước (%)",
                    DSChiQuay = listReportDataPercentClone[1].DSChiQuay == 0 ? 0 : Math.Round(dsChiQuaySum / listReportDataPercentClone[1].DSChiQuay * 100, 2, MidpointRounding.ToEven),
                    DSChiNha = listReportDataPercentClone[1].DSChiNha == 0 ? 0 : Math.Round(dsChiNhaSum / listReportDataPercentClone[1].DSChiNha * 100, 2, MidpointRounding.ToEven),
                    DSCK = listReportDataPercentClone[1].DSCK == 0 ? 0 : Math.Round(dsCKSum / listReportDataPercentClone[1].DSCK * 100, 2, MidpointRounding.ToEven),
                    TongDS = listReportDataPercentClone[1].TongDS == 0 ? 0 : Math.Round(tongDSSum / listReportDataPercentClone[1].TongDS * 100, 2, MidpointRounding.ToEven),
                };

                listReportDataPercentClone.Add(dataDifferencePercentLastMonth);

                // ds so với tháng trước
                double dsChiQuaySumLastYear = (listReportDataPercentClone[0].DSChiQuay - listReportDataPercentClone[2].DSChiQuay);
                double dsChiNhaSumLastYear = (listReportDataPercentClone[0].DSChiNha - listReportDataPercentClone[2].DSChiNha);
                double dsCKSumLastYear = (listReportDataPercentClone[0].DSCK - listReportDataPercentClone[2].DSCK);
                double tongDSSumLastYear = (listReportDataPercentClone[0].TongDS - listReportDataPercentClone[2].TongDS);

                // Object báo cáo tăng giảm so với cùng kỳ (%)
                Report dataDifferencePercentLastYear = new Report()
                {
                    GradationID = "Tăng giảm so với cùng kì năm trước (%)",
                    DSChiQuay = listReportDataPercentClone[2].DSChiQuay == 0 ? 0 : Math.Round(dsChiQuaySumLastYear / listReportDataPercentClone[2].DSChiQuay * 100, 2, MidpointRounding.ToEven),
                    DSChiNha = listReportDataPercentClone[2].DSChiNha == 0 ? 0 : Math.Round(dsChiNhaSumLastYear / listReportDataPercentClone[2].DSChiNha * 100, 2, MidpointRounding.ToEven),
                    DSCK = listReportDataPercentClone[2].DSCK == 0 ? 0 : Math.Round(dsCKSumLastYear / listReportDataPercentClone[2].DSCK * 100, 2, MidpointRounding.ToEven),
                    TongDS = listReportDataPercentClone[2].TongDS == 0 ? 0 : Math.Round(tongDSSumLastYear / listReportDataPercentClone[2].TongDS * 100, 2, MidpointRounding.ToEven),
                };

                listReportDataPercentClone.Add(dataDifferencePercentLastYear);

                DataTable dataTablePie = new DataTable();
                // Add item to datatable
                dataTablePie = CreateDataTableFormartGradation();

                // Danh sách dataSet của báo cáo ngày
                DataSet dataReportPie = new DongA.Core.DongAExcel().ConvertListObjectToDataSet(listReportDataPercentClone);
                // Đổ data vào datatble mới
                FillDataGradationCompare(dataReportPie.Tables[0], dataTablePie);
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
                        int totalCol = 15 + 5;
                        for (int b = 15; b < totalCol; b++)
                        {
                            // Giá trị của value trong table
                            string valueOfTable = dataTablePie.Rows[stepRow][stepColumn].ToString();

                            // Tô màu cho các dòng có giá trị tăng giảm
                            if (b >= 16 && a >= 38)
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
        /// Tạo cột cho datatable
        /// </summary>
        /// <returns></returns>
        private DataTable CreateDataTableFormart(string typeID)
        {
            DataTable db = new DataTable();
            db.Columns.Add("ReportID", typeof(string));

            db.Columns.Add("DSChiQuay", typeof(double));
            db.Columns.Add("DSChiNha", typeof(double));
            db.Columns.Add("DSCK", typeof(double));
            db.Columns.Add("TongDS", typeof(double));
            return db;
        }

        /// <summary>
        /// Tạo cột cho datatable
        /// </summary>
        /// <returns></returns>
        private DataTable CreateDataTableFormartGradation()
        {
            DataTable db = new DataTable();
            db.Columns.Add("GradationID", typeof(string));
            db.Columns.Add("DSChiQuay", typeof(double));
            db.Columns.Add("DSChiNha", typeof(double));
            db.Columns.Add("DSCK", typeof(double));
            db.Columns.Add("TongDS", typeof(double));
            return db;
        }

        /// <summary>
        /// Add row cho datatable
        /// </summary>
        /// <param name="mother"></param>
        /// <param name="fill"></param>
        private void FillData(DataTable mother, DataTable fill, string typeID)
        {
            int stt = 1;
            foreach (DataRow dr in mother.Rows)
            {
                var row = fill.NewRow();
                row["ReportID"] = (string)dr["ReportID"];

                row["DSChiQuay"] = (double)dr["DSChiQuay"];
                row["DSChiNha"] = (double)dr["DSChiNha"];
                row["DSCK"] = (double)dr["DSCK"];
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
        private void FillDataGradationCompare(DataTable mother, DataTable fill)
        {
            int stt = 1;
            foreach (DataRow dr in mother.Rows)
            {
                var row = fill.NewRow();

                row["GradationID"] = (string)dr["GradationID"];
                row["DSChiQuay"] = (double)dr["DSChiQuay"];
                row["DSChiNha"] = (double)dr["DSChiNha"];
                row["DSCK"] = (double)dr["DSCK"];
                row["TongDS"] = (double)dr["TongDS"];
                fill.Rows.Add(row);
                stt++;
            }
        }
    }
}
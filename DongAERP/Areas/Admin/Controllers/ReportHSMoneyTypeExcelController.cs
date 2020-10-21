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
    public class ReportHSMoneyTypeExcelController : Controller
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

        // GET: Admin/ReportHSMoneyTypeExcel
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
            string templatePath = "~/Content/Report/ReportHS/ReportHSTotalMoneyType.xlsx";
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
            List<ReportForTotalMoneyType> listReportDataConvertClone = new List<ReportForTotalMoneyType>();
            
            switch (typeID)
            {
                // Theo ngày
                case 1:
                    listReportData = new HSReportBL().SearchReportHSTotalMoneyTypeForDay(fromDate, toDate, reportTypeID);
                    break;
                // Theo tháng
                case 2:

                    listReportData = new HSReportBL().SearchReportHSTotalMoneyTypeForMonth(fromDate, toDate, reportTypeID);
                    break;
                // Theo năm
                default:

                    listReportData = new HSReportBL().SearchReportHSTotalMoneyTypeForYear(fromDate, toDate, reportTypeID);
                    listReportDataClone = new List<ReportForTotalMoneyType>(listReportData);
                    break;
            }

            // Set border
            Style style = new CellsFactory().CreateStyle();
            style.SetBorder(BorderType.TopBorder, CellBorderType.Thin, Color.Black);
            style.SetBorder(BorderType.BottomBorder, CellBorderType.Thin, Color.Black);
            style.SetBorder(BorderType.LeftBorder, CellBorderType.Thin, Color.Black);
            style.SetBorder(BorderType.RightBorder, CellBorderType.Thin, Color.Black);


            DataTable dataTable = new DataTable();
            
            if (listReportData.Count > 0)
            {
                foreach (ReportForTotalMoneyType item in listReportData)
                {
                    switch (typeID)
                    {
                        // Theo ngày
                        case 1:
                            item.ReportID = string.Format("Ngày {0}", item.CreatedDate.ToString("dd/MM/yyyy"));
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

            int countRow = 0;
            if (dataTable.Rows.Count > 0)
            {
                int stepRow = 0;
                // total row = row start + số row hiện có
                int totalRow = dataTable.Rows.Count + 7;
                countRow = totalRow;
                // Số dòng của row
                for (int a = 7; a < totalRow; a++)
                {
                    int stepColumn = 0;
                    // Số cột trong báo cáo cần hiển thị
                    // Tổng số cột hiển thị = Số cột hiển thị bắt đầu + tổng số cột cần hiển thị
                    // trường hợp với report chọn tháng
                    int totalCol = 15 + 8;
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
                        style.Custom = "#,##";

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

            // Adding SeriesCollection (chart data source) to the chart ranging from "A1" cell to "B3"
            //string totalRowData = string.Format("Q{0}:W{1}", countRow, countRow + listReportDataConvert.Count - 2);
            string totalRowData = string.Format("Q8:V{0}", countRow - 1);
            leadSourceLine.NSeries.Add(totalRowData, true);

            // Set the category data covering the range A2:A5.
            // Tổng số dòng cần hiển thị là số dòng hiện tại  + số dòng trong listData -2 (trong đó 1 là dòng tổng cuối cùng)
            string categoryData = string.Format("P8:P{0}", countRow - 1);
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
            string templatePath = "~/Content/Report/ReportHS/ReportHSTotalMoneyTypeGradation.xlsx";
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
            string titleDS = "1. Theo hồ sơ chi trả loại hình dịch vụ";
            CreateTitle("A5", "E5", sheetReport, titleDS, 12);

            // Tạo title chotỷ trọng chi trả loại hình dịch vụ
            string titleTT = "2. Theo tỷ trọng chi trả loại hình dịch vụ";
            CreateTitle("A33", "E33", sheetReport, titleTT, 12);

            // Add tên cột cho bảng báo cáo
            // Set width cho column

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
            string[] cellArray = { "P7", "Q7", "R7", "S7", "T7", "Q35", "R35", "S35", "P35" };
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


            List<ReportForTotalMoneyType> listReportData = new HSReportBL().SearchReportHSTotalMoneyTypeForGradationCompare(year, int.Parse(gradationID), reportTypeID);
            foreach(ReportForTotalMoneyType item in listReportData)
            {
                item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;
            }
            // clone Object
            List<ReportForTotalMoneyType> listReportDataClone = new List<ReportForTotalMoneyType>(listReportData);
            
            // Danh sách doanh số chi trả theo tỉ trọng
            List<ReportForTotalMoneyType> listReportDataPercent = new HSReportBL().SearchReportHSTotalMoneyTypeForGradationComparePercent(year, int.Parse(gradationID), reportTypeID);
            // clone Object
            List<ReportForTotalMoneyType> listReportDataPercentClone = new List<ReportForTotalMoneyType>(listReportDataPercent);


            DataTable dataTable = new DataTable();

            string[] str = { "VND", "USD", "EUR", "CAD", "AUD", "GBP", "Tổng" };
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

                double TongDSCompare = listReportData[0].TongDS - listReportData[1].TongDS;
                
                // add row vào table
                dataTable.Rows.Add(str[0], listReportData[0].VND, listReportData[1].VND, Math.Round(VNDCompare / listReportData[1].VND * 100, 2, MidpointRounding.ToEven), VNDCompare);
                dataTable.Rows.Add(str[1], listReportData[0].USD, listReportData[1].USD, Math.Round(USDCompare / listReportData[1].USD * 100, 2, MidpointRounding.ToEven), USDCompare);
                dataTable.Rows.Add(str[2], listReportData[0].EUR, listReportData[1].EUR, Math.Round(EURCompare / listReportData[1].EUR * 100, 2, MidpointRounding.ToEven), EURCompare);
                dataTable.Rows.Add(str[3], listReportData[0].CAD, listReportData[1].CAD, Math.Round(CADCompare / listReportData[1].CAD * 100, 2, MidpointRounding.ToEven), CADCompare);
                dataTable.Rows.Add(str[4], listReportData[0].AUD, listReportData[1].AUD, Math.Round(AUDCompare / listReportData[1].AUD * 100, 2, MidpointRounding.ToEven), AUDCompare);
                dataTable.Rows.Add(str[5], listReportData[0].GBP, listReportData[1].GBP, Math.Round(GBPCompare / listReportData[1].GBP * 100, 2, MidpointRounding.ToEven), GBPCompare);

                dataTable.Rows.Add(str[6], listReportData[0].TongDS, listReportData[1].TongDS, Math.Round(TongDSCompare / listReportData[1].TongDS * 100, 2, MidpointRounding.ToEven), TongDSCompare);
                
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

                            if (b == 18)
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


            // Create Chart column
            //Chart reference
            Aspose.Cells.Charts.Chart leadSourceColumn;
            //Add Pie Chart
            int chartIndex = sheetReport.Charts.Add(ChartType.Column3DClustered, 6, 0, 30, 13);
            leadSourceColumn = sheetReport.Charts[chartIndex];


            //Chart title
            leadSourceColumn.Title.Text = string.Format("Hồ sơ loại tiền chi trả \n Giai đoạn: {0}", text);
            leadSourceColumn.Title.Font.Color = Color.Silver;

            // Adding SeriesCollection (chart data source) to the chart ranging from "A1" cell to "B3"
            string totalRowData = "Q8:R13";
            leadSourceColumn.NSeries.Add(totalRowData, true);

            // Set the category data covering the range A2:A5.
            // Tổng số dòng cần hiển thị là số dòng hiện tại  + số dòng trong listData -2 (trong đó 1 là dòng tổng cuối cùng)
            string categoryData = "P8:P13";
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
                    totalRowData = string.Concat("{", string.Format("{0}, {1}, {2}, {3}, {4}, {5}", dataPieYear.VND, dataPieYear.USD, dataPieYear.EUR, dataPieYear.CAD, dataPieYear.AUD, dataPieYear.GBP), "}");
                    leadSourcePie.NSeries.Add(totalRowData, true);

                    categoryData = string.Concat("{", string.Join(", ", str), "}");
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
                    totalRowData = string.Concat("{", string.Format("{0}, {1}, {2}, {3}, {4}, {5}", dataPieYear.VND, dataPieYear.USD, dataPieYear.EUR, dataPieYear.CAD, dataPieYear.AUD, dataPieYear.GBP), "}");
                    leadSourcePieLasYear.NSeries.Add(totalRowData, true);

                    categoryData = string.Concat("{", string.Join(", ", str), "}");
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
                double VNDCompare = dataPieYear.VND - dataPieLastYear.VND;
                double USDCompare = dataPieYear.USD - dataPieLastYear.USD;
                double EURCompare = dataPieYear.EUR - dataPieLastYear.EUR;
                double CADCompare = dataPieYear.CAD - dataPieLastYear.CAD;
                double AUDCompare = dataPieYear.AUD - dataPieLastYear.AUD;
                double GBPCompare = dataPieYear.GBP - dataPieLastYear.GBP;

                // Tạo các cột cho datatable
                DataTable dataTablePie = new DataTable();

                dataTablePie.Columns.Add("ReportID", typeof(String));
                dataTablePie.Columns.Add("AccumulateID1", typeof(double));
                dataTablePie.Columns.Add("AccumulateID2", typeof(double));
                dataTablePie.Columns.Add("CompareToIDPercent", typeof(double));

                // add row vào table
                dataTablePie.Rows.Add(str[0], dataPieYear.VND, dataPieLastYear.VND, dataPieLastYear.VND == 0 ? 0 : Math.Round((VNDCompare / dataPieLastYear.VND) * 100, 2, MidpointRounding.ToEven));
                dataTablePie.Rows.Add(str[1], dataPieYear.USD, dataPieLastYear.USD, dataPieLastYear.USD == 0 ? 0 : Math.Round((USDCompare / dataPieLastYear.USD) * 100, 2, MidpointRounding.ToEven));
                dataTablePie.Rows.Add(str[2], dataPieYear.EUR, dataPieLastYear.EUR, dataPieLastYear.EUR == 0 ? 0 : Math.Round((EURCompare / dataPieLastYear.EUR) * 100, 2, MidpointRounding.ToEven));
                dataTablePie.Rows.Add(str[3], dataPieYear.CAD, dataPieLastYear.CAD, dataPieLastYear.CAD == 0 ? 0 : Math.Round((CADCompare / dataPieLastYear.CAD) * 100, 2, MidpointRounding.ToEven));
                dataTablePie.Rows.Add(str[4], dataPieYear.AUD, dataPieLastYear.AUD, dataPieLastYear.AUD == 0 ? 0 : Math.Round((AUDCompare / dataPieLastYear.AUD) * 100, 2, MidpointRounding.ToEven));
                dataTablePie.Rows.Add(str[5], dataPieYear.GBP, dataPieLastYear.GBP, dataPieLastYear.GBP == 0 ? 0 : Math.Round((GBPCompare / dataPieLastYear.GBP) * 100, 2, MidpointRounding.ToEven));

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
        /// Tạo mẫu cho Excel cho so sánh theo giai đoạn
        /// </summary>
        /// <param name="gradationID"></param>
        /// <param name="year"></param>
        /// <param name="typeID"></param>
        /// <returns></returns>
        public ActionResult CreateExcelForCompareYear(int year, int month, string reportTypeID)
        {
            WorkbookDesigner designer = new WorkbookDesigner();
            string templatePath = "~/Content/Report/ReportHS/ReportHSTotalMoneyType.xlsx";
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


            sheetReport.Cells["Q7"].PutValue(string.Format("Tháng {0}/{1}", month, year));
            //sheetReport.Cells.SetColumnWidthPixel(16, 170);
            sheetReport.Cells["R7"].PutValue(string.Format("Tháng {0}/{1}", month - 1, year));
            if(month == 1)
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
            string[] cellArray = { "Q7", "R7", "S7", "T7", "U7", "V7", "W7", "P35", "Q35", "R35", "S35", "T35", "U35" };
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
            List<ReportForTotalMoneyType> listReportData = new HSReportBL().SearchReportHSTotalMoneyTypeForCompareMonth(year, month, reportTypeID);
            foreach(ReportForTotalMoneyType item in listReportData)
            {
                item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;
            }
            // clone Object
            List<ReportForTotalMoneyType> listReportDataClone = new List<ReportForTotalMoneyType>(listReportData);
           
            // Danh sách tỉ trọng chi trả theo loại hình dịch vụ
            List<ReportForTotalMoneyType> listReportDataPercent = new HSReportBL().SearchReportHSTotalMoneyTypeForCompareMonthPercent(year, month, reportTypeID);
            // clone Object
            List<ReportForTotalMoneyType> listReportDataPercentClone = new List<ReportForTotalMoneyType>(listReportDataPercent);

            DataTable dataTable = new DataTable();

            dataTable.Columns.Add("ReportID", typeof(String));
            dataTable.Columns.Add("AccumulateID1", typeof(double));
            dataTable.Columns.Add("AccumulateID2", typeof(double));
            dataTable.Columns.Add("AccumulateID3", typeof(double));
            dataTable.Columns.Add("CompareToMonth", typeof(double));
            dataTable.Columns.Add("CompareToMonthPercent", typeof(double));
            dataTable.Columns.Add("CompareToMonthLastYear", typeof(double));
            dataTable.Columns.Add("CompareToMonthLastYearPercent", typeof(double));

            string[] str = { "VND", "USD", "EUR", "CAD", "AUD", "GBP", "Tổng" };

            // Theo nguyên tệ
            if (listReportData.Count.Equals(3))
            {

                // tháng hiện tại
                double VNDCompareMonth = listReportData[0].VND - listReportData[1].VND;
                double USDCompareMonth = listReportData[0].USD - listReportData[1].USD;
                double EURCompareMonth = listReportData[0].EUR - listReportData[1].EUR;
                double CADCompareMonth = listReportData[0].CAD - listReportData[1].CAD;
                double AUDCompareMonth = listReportData[0].AUD - listReportData[1].AUD;
                double GBPCompareMonth = listReportData[0].GBP - listReportData[1].GBP;

                double TongDSCompareMonth = listReportData[0].TongDS - listReportData[1].TongDS;

                // tháng cùng kì năm trước
                double VNDCompareLastMonth = listReportData[0].VND - listReportData[2].VND;
                double USDCompareLastMonth = listReportData[0].USD - listReportData[2].USD;
                double EURCompareLastMonth = listReportData[0].EUR - listReportData[2].EUR;
                double CADCompareLastMonth = listReportData[0].CAD - listReportData[2].CAD;
                double AUDCompareLastMonth = listReportData[0].AUD - listReportData[2].AUD;
                double GBPCompareLastMonth = listReportData[0].GBP - listReportData[2].GBP;

                double TongDSCompareLastMonth = listReportData[0].TongDS - listReportData[2].TongDS;

                // add row vào table
                dataTable.Rows.Add(str[0], listReportData[0].VND, listReportData[1].VND, listReportData[2].VND
                    , VNDCompareMonth, listReportData[1].VND == 0 ? 0 : Math.Round(VNDCompareMonth / listReportData[1].VND * 100, 2, MidpointRounding.ToEven)
                    , VNDCompareLastMonth, listReportData[2].VND == 0 ? 0 : Math.Round(VNDCompareLastMonth / listReportData[2].VND * 100, 2, MidpointRounding.ToEven));

                dataTable.Rows.Add(str[1], listReportData[0].USD, listReportData[1].USD, listReportData[2].USD
                    , USDCompareMonth, listReportData[1].USD == 0 ? 0 : Math.Round(USDCompareMonth / listReportData[1].USD * 100, 2, MidpointRounding.ToEven)
                    , USDCompareLastMonth, listReportData[2].USD == 0 ? 0 : Math.Round(USDCompareLastMonth / listReportData[2].USD * 100, 2, MidpointRounding.ToEven));

                dataTable.Rows.Add(str[2], listReportData[0].EUR, listReportData[1].EUR, listReportData[2].EUR
                    , EURCompareMonth, listReportData[1].EUR == 0 ? 0 : Math.Round(EURCompareMonth / listReportData[1].EUR * 100, 2, MidpointRounding.ToEven)
                    , EURCompareLastMonth, listReportData[2].EUR == 0 ? 0 : Math.Round(EURCompareLastMonth / listReportData[2].EUR * 100, 2, MidpointRounding.ToEven));

                dataTable.Rows.Add(str[3], listReportData[0].CAD, listReportData[1].CAD, listReportData[2].CAD
                    , CADCompareMonth, listReportData[1].CAD == 0 ? 0 : Math.Round(CADCompareMonth / listReportData[1].CAD * 100, 2, MidpointRounding.ToEven)
                    , CADCompareLastMonth, listReportData[2].CAD == 0 ? 0 : Math.Round(CADCompareLastMonth / listReportData[2].CAD * 100, 2, MidpointRounding.ToEven));

                dataTable.Rows.Add(str[4], listReportData[0].AUD, listReportData[1].AUD, listReportData[2].AUD
                    , AUDCompareMonth, listReportData[1].AUD == 0 ? 0 : Math.Round(AUDCompareMonth / listReportData[1].AUD * 100, 2, MidpointRounding.ToEven)
                    , AUDCompareLastMonth, listReportData[2].AUD == 0 ? 0 : Math.Round(AUDCompareLastMonth / listReportData[2].AUD * 100, 2, MidpointRounding.ToEven));

                dataTable.Rows.Add(str[5], listReportData[0].GBP, listReportData[1].GBP, listReportData[2].GBP
                    , GBPCompareMonth, listReportData[1].GBP == 0 ? 0 : Math.Round(GBPCompareMonth / listReportData[1].GBP * 100, 2, MidpointRounding.ToEven)
                    , GBPCompareLastMonth, listReportData[2].GBP == 0 ? 0 : Math.Round(GBPCompareLastMonth / listReportData[2].GBP * 100, 2, MidpointRounding.ToEven));

                dataTable.Rows.Add(str[6], listReportData[0].TongDS, listReportData[1].TongDS, listReportData[2].TongDS
                    , TongDSCompareMonth, listReportData[1].TongDS == 0 ? 0 : Math.Round(TongDSCompareMonth / listReportData[1].TongDS * 100, 2, MidpointRounding.ToEven)
                    , TongDSCompareLastMonth, listReportData[2].TongDS == 0 ? 0 : Math.Round(TongDSCompareLastMonth / listReportData[2].TongDS * 100, 2, MidpointRounding.ToEven));

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
                            // Trường hợp với cột là %
                            if (b.Equals(totalCol - 1) || b.Equals(totalCol - 3))
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
            foreach (ReportForTotalMoneyType item in listReportDataClone)
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
                if (month == 1)
                {
                    dataPieLastMonth = listReportDataPercentClone.Find(x => x.Year == (year - 1).ToString() && x.Month == "12");
                }
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
            return ExportReport("ReportCompareMonth", designer);
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
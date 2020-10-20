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
    public class ReportLayoutController : Controller
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

        // GET: Admin/ReportLayout
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult CreateExcel(DateTime fromDate, DateTime toDate, string typeID, string reportTypeID)
        {
            WorkbookDesigner designer = new WorkbookDesigner();
            string templatePath = "~/Content/Report/ReportDetail.xlsx";
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
            
            CreateTitle("A2", "U2", sheetReport, typeReport);

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
            leadSourceLine.Title.Text = "Doanh số chi trả theo từng dịch vụ";
            leadSourceLine.Title.Font.Color = Color.Silver;


            List<Report> listReportData = new List<Report>();
            int typeIDForDate = int.Parse(typeID);
            // Danh sách ngày
            switch (typeID)
            {
                // Theo ngày
                case "1":
                    listReportData = new ReportBL().SearchDay(fromDate, toDate, reportTypeID);
                    break;
                // Theo tháng
                case "2":
                    listReportData = new ReportBL().SearchMonth(fromDate, toDate, reportTypeID);
                    break;
                // Theo năm
                default:
                    listReportData = new ReportBL().SearchYear(fromDate, toDate, reportTypeID);
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
        /// Hàm tạo tile cho xuất Excell
        /// </summary>
        /// <param name="upperLeft"></param>
        /// <param name="upperRight"></param>
        /// <param name="sheetReport"></param>
        /// <param name="Title"></param>
        private void CreateTitle(string upperLeft, string upperRight, Worksheet sheetReport, string Title)
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
            styleTitle.Font.Size = 14;
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
    }
}
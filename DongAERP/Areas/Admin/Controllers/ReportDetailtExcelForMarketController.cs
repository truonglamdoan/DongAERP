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
    public class ReportDetailtExcelForMarketController : Controller
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

        // GET: Admin/ReportDetailtExcelForMarket
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult CreateExcelForMarket(DateTime fromDate, DateTime toDate, string typeID, string reportTypeID, string marketID)
        {
            WorkbookDesigner designer = new WorkbookDesigner();
            string templatePath = "~/Content/Report/ReportDetailForMarket.xlsx";
            // Get đường dẫn
            templatePath = System.Web.HttpContext.Current.Server.MapPath(templatePath);

            designer.Workbook = new DongA.Core.DongAExcel().OpenExcelFile(templatePath);
            designer.Workbook.CalculateFormula();

            WorksheetCollection workSheets = designer.Workbook.Worksheets;
            Worksheet sheetReport = designer.Workbook.Worksheets[0];

            // Tạo title
            string typeReport = "Tất cả thị trường";
            // Trường hợp typeID = 1 là chọn thị trường tất cả
            // typeID = 2 là chọn từng thị trường
            if (marketID.Equals("1"))
            {
                typeReport = "Thị trường: Châu Á";
            }

            CreateTitle("A2", "K2", sheetReport, typeReport, 14);

            // Set from day and to day
            sheetReport.Cells["F3"].PutValue(fromDate.ToString("dd/MM/yyyy"));
            sheetReport.Cells["H3"].PutValue(toDate.ToString("dd/MM/yyyy"));

            List<ReportDetailtSTMarket> listReportData = new List<ReportDetailtSTMarket>();
            // Danh sách ngày
            switch (typeID)
            {
                // Theo ngày
                case "1":
                    listReportData = new ReportBL().SearchReportDetailtForDay(fromDate, toDate, reportTypeID, marketID);
                    break;
                // Theo tháng
                // Theo năm
                default:
                    //listReportData = new ReportBL().SearchMarketForOne(fromDate, toDate, reportTypeID, marketID);
                    break;
            }

            DataTable dataTable = new DataTable();

            if (listReportData.Count > 0)
            {
                int count = 1;
                foreach (ReportDetailtSTMarket item in listReportData)
                {
                    item.STT = count.ToString();
                    item.TongDS = item.DSChiNha + item.DSChiQuay + item.DSCK;
                    count++;
                }

                ReportDetailtSTMarket dataItem = new ReportDetailtSTMarket()
                {
                    STT = "",
                    MarketName = "Tổng",
                    DSChiQuay = listReportData.Sum(x => x.DSChiQuay),
                    DSChiNha = listReportData.Sum(x => x.DSChiNha),
                    DSCK = listReportData.Sum(x => x.DSCK),
                    TongDS = listReportData.Sum(x => x.TongDS)
                };

                listReportData.Add(dataItem);

                // Tạo các col cho table
                dataTable = CreateDataTableFormart(typeID);

                // Danh sách dataSet của báo cáo ngày
                DataSet dataReport = ConvertListObjectToDataSet(listReportData);

                // Đổ data vào datatble mới
                FillData(dataReport.Tables[0], dataTable, typeID);

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
                int totalRow = dataTable.Rows.Count + 6;
                // Số dòng của row
                for (int a = 6; a < totalRow; a++)
                {
                    int stepColumn = 0;
                    // Số cột trong báo cáo cần hiển thị
                    // Tổng số cột hiển thị = Số cột hiển thị bắt đầu + tổng số cột cần hiển thị
                    int totalCol = 3 + 6;
                    for (int b = 3; b < totalCol; b++)
                    {
                        // Giá trị của value trong table
                        string valueOfTable = dataTable.Rows[stepRow][stepColumn].ToString();

                        // Insert vào dòng cột xác định trong Excel
                        sheetReport.Cells[a, b].PutValue(valueOfTable, true);
                        // set style số cho các cột khác cột STT
                        if (!b.Equals(3))
                        {
                            // set style cho number
                            style.Custom = "#,##0.00";
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
            return ExportReport("ReportDetailtForMarket", designer);
        }

        public ActionResult CreateExcelMarketForOne(DateTime fromDate, DateTime toDate, string typeID, string reportTypeID, string marketID)
        {
            WorkbookDesigner designer = new WorkbookDesigner();
            string templatePath = "~/Content/Report/ReportDetailForMarket.xlsx";
            // Get đường dẫn
            templatePath = System.Web.HttpContext.Current.Server.MapPath(templatePath);

            designer.Workbook = new DongA.Core.DongAExcel().OpenExcelFile(templatePath);
            designer.Workbook.CalculateFormula();

            WorksheetCollection workSheets = designer.Workbook.Worksheets;
            Worksheet sheetReport = designer.Workbook.Worksheets[0];

            // Tạo title
            string typeReport = "Tất cả thị trường";
            // Trường hợp typeID = 1 là chọn thị trường tất cả
            // typeID = 2 là chọn từng thị trường
            if (marketID.Equals("1"))
            {
                typeReport = "Thị trường: Châu Á";
            }

            CreateTitle("A2", "K2", sheetReport, typeReport, 14);

            // Set from day and to day
            sheetReport.Cells["F3"].PutValue(fromDate.ToString("dd/MM/yyyy"));
            sheetReport.Cells["H3"].PutValue(toDate.ToString("dd/MM/yyyy"));

            List<ReportDetailtServiceType> listReportData = new List<ReportDetailtServiceType>();
            listReportData = new ReportBL().SearchMarketForOne(fromDate, toDate, reportTypeID, marketID);

            DataTable dataTable = new DataTable();

            if (listReportData.Count > 0)
            {
                // List thị trường
                int count = 1;
                // Danh sách các thị trường thuộc Châu Á
                List<string> listMarketCode = new List<string>();

                foreach (ReportDetailtServiceType item in listReportData)
                {
                    if (string.IsNullOrEmpty(item.MarketName))
                    {
                        item.MarketName = "Tất cả Thị trường";
                        item.STT = (count).ToString();
                        count++;
                    }
                    else
                    {
                        // Trường hợp không có code thị trường trong list tạm thì add vào list tạm
                        if (listMarketCode.Contains(item.MarketCode))
                        {
                            count++;
                        }
                        else
                        {
                            count = 1;
                            listMarketCode.Add(item.MarketCode);
                        }

                        item.STT = (count).ToString();
                    }
                    item.TongDS = item.DSChiNha + item.DSChiQuay + item.DSCK;
                }

                ReportDetailtServiceType dataItem = new ReportDetailtServiceType()
                {
                    STT = "Tổng",
                    MarketName = "",
                    PartnerName = "",
                    DSChiQuay = listReportData.Sum(x => x.DSChiQuay),
                    DSChiNha = listReportData.Sum(x => x.DSChiNha),
                    DSCK = listReportData.Sum(x => x.DSCK),
                    TongDS = listReportData.Sum(x => x.TongDS)
                };

                listReportData.Add(dataItem);

                // Tạo các col cho table
                dataTable = CreateDataTableFormart(typeID);

                // Danh sách dataSet của báo cáo ngày
                DataSet dataReport = ConvertListObjectToDataSet(listReportData);

                // Đổ data vào datatble mới
                FillData(dataReport.Tables[0], dataTable, typeID);
            }

            // Set border
            Style style = new CellsFactory().CreateStyle();
            style.SetBorder(BorderType.BottomBorder, CellBorderType.Thin, Color.Black);
            style.SetBorder(BorderType.LeftBorder, CellBorderType.Thin, Color.Black);
            style.SetBorder(BorderType.RightBorder, CellBorderType.Thin, Color.Black);
            style.SetBorder(BorderType.TopBorder, CellBorderType.Thin, Color.Black);

            if (dataTable.Rows.Count > 0)
            {
                int stepRow = 0;
                // total row = row start + số row hiện có
                int totalRow = dataTable.Rows.Count + 6;

                // Danh sách các thị trường thuộc Châu Á
                List<string> listMarketCode = new List<string>();
                // Số dòng của row

                int rowNumber = 6;
                for (int a = 6; a < totalRow; a++)
                {
                    int stepColumn = 0;
                    // Số cột trong báo cáo cần hiển thị
                    // Tổng số cột hiển thị = Số cột hiển thị bắt đầu + tổng số cột cần hiển thị
                    int totalCol = 3 + 6;

                    for (int b = 3; b < totalCol; b++)
                    {
                        string valueCheck = dataTable.Rows[stepRow][1].ToString();

                        // Check tồn tại trong listTemp
                        if (!listMarketCode.Contains(valueCheck) && !string.IsNullOrEmpty(valueCheck))
                        {
                            listMarketCode.Add(valueCheck);

                            // Insert vào dòng cột xác định trong Excel
                            sheetReport.Cells[rowNumber, b].PutValue(valueCheck, true);

                            // Set isBold cho dòng xác định là thị trường
                            style.Font.IsBold = true;
                            sheetReport.Cells[rowNumber, b].SetStyle(style);
                            // chuyển lại chế độ mặt định để không ảnh hưởng đến các thành phần khác
                            style.Font.IsBold = false;

                            //sheetReport.Cells
                            // Tăng số dòng lên
                            rowNumber++;
                        }

                        if (dataTable.Columns[stepColumn].ColumnName.Contains("MarketName"))
                        {
                            stepColumn++;
                        }

                        // Giá trị của value trong table
                        string valueOfTable = dataTable.Rows[stepRow][stepColumn].ToString();

                        // Insert vào dòng cột xác định trong Excel
                        sheetReport.Cells[rowNumber, b].PutValue(valueOfTable, true);
                        // set style số cho các cột khác cột STT
                        if (!b.Equals(3))
                        {
                            // set style cho number
                            style.Custom = "#,##0.00";
                        }
                        else
                        {
                            style.Custom = "#,##0";
                        }

                        // set border
                        sheetReport.Cells[rowNumber, b].SetStyle(style);

                        // Cột tổng cộng
                        if (b.Equals(totalCol - 1))
                        {
                            sheetReport.Cells[rowNumber, b].PutValue(valueOfTable, true, true);
                            style.Font.IsBold = true;
                            sheetReport.Cells[rowNumber, b].SetStyle(style);
                        }

                        // Set lại giá trị mặt định
                        style.Font.IsBold = false;
                        // Tăng cột theo dòng của table
                        stepColumn++;
                    }
                    // Tăng dòng của table lên
                    stepRow++;
                    rowNumber++;
                }

                // Trường hợp thuộc dòng cuối Tổng
                //Bold style flag options

                StyleFlag boldStyleFlag = new StyleFlag();

                boldStyleFlag.HorizontalAlignment = true;

                boldStyleFlag.FontBold = true;

                style.Font.IsBold = true;
                sheetReport.Cells.Rows[rowNumber - 1].ApplyStyle(style, boldStyleFlag);
            }
            else
            {
                sheetReport.Cells["D10"].PutValue("Không có dữ liệu");
            }

            //sheetReport.Cells.ConvertStringToNumericValue();
            // Chạy process
            designer.Process();
            return ExportReport("ReportDetailMarketForOne", designer);
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
            string templatePath = "~/Content/Report/ReportDetailMarketForGaration.xlsx";
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
            CreateTitle("A2", "K2", sheetReport, typeReport, 14);
            // Tạo title detailt
            string titleDetailt = string.Format("Giai đoạn: {0}", text);
            CreateTitle("A3", "K3", sheetReport, titleDetailt, 12);

            // Nguyên tệ
            List<ReportDetailtServiceType> listReportData = new ReportBL().ReportDetailtGradationCompareForAll(year, int.Parse(gradationID), reportTypeID);
            // clone Object
            List<ReportDetailtServiceType> listReportDataClone = new List<ReportDetailtServiceType>(listReportData);

            foreach (ReportDetailtServiceType item in listReportData)
            {
                item.TongDS = item.DSChiQuay + item.DSChiNha + item.DSCK;
            }

            // Danh sách mã thị trường của Tất cả
            List<string> listMarket = new List<string>() { "003", "005", "001", "002", "014", "004" };

            DataTable dataTable = new DataTable();

            // Tạo các cột cho datatable
            dataTable.Columns.Add("MarketName", typeof(String));
            dataTable.Columns.Add("CQ1", typeof(double));
            dataTable.Columns.Add("CN1", typeof(double));
            dataTable.Columns.Add("CK1", typeof(double));
            dataTable.Columns.Add("TDS1", typeof(double));

            dataTable.Columns.Add("CQ2", typeof(double));
            dataTable.Columns.Add("CN2", typeof(double));
            dataTable.Columns.Add("CK2", typeof(double));
            dataTable.Columns.Add("TDS2", typeof(double));

            dataTable.Columns.Add("CQ3", typeof(double));
            dataTable.Columns.Add("CN3", typeof(double));
            dataTable.Columns.Add("CK3", typeof(double));
            dataTable.Columns.Add("TDS3", typeof(double));

            if (listReportData.Count > 0)
            {
                foreach (string item in listMarket)
                {
                    // Cùng kì
                    ReportDetailtServiceType dataItemLastYear = listReportData.Find(x => x.MarketCode == item && x.Year == (year - 1).ToString());
                    ReportDetailtServiceType dataItemYear = listReportData.Find(x => x.MarketCode == item && x.Year == year.ToString());

                    // Trường hợp năm không có đối tác
                    if (dataItemLastYear == null && dataItemYear != null)
                    {
                        dataItemLastYear = new ReportDetailtServiceType();
                        dataItemLastYear.MarketName = dataItemYear.MarketName;
                        dataItemLastYear.DSChiQuay = 0;
                        dataItemLastYear.DSChiNha = 0;
                        dataItemLastYear.DSCK = 0;
                        dataItemLastYear.Year = (year - 1).ToString();
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
                    }

                    double sumDSChiQuay = dataItemYear.DSChiQuay - dataItemLastYear.DSChiQuay;
                    double sumDSChiNha = dataItemYear.DSChiNha - dataItemLastYear.DSChiNha;
                    double sumDSCK = dataItemYear.DSCK - dataItemLastYear.DSCK;

                    double sumTongDS = dataItemYear.TongDS - dataItemLastYear.TongDS;


                    dataTable.Rows.Add(dataItemLastYear.MarketName, dataItemYear.DSChiQuay, dataItemYear.DSChiNha, dataItemYear.DSCK, dataItemYear.TongDS
                        , dataItemLastYear.DSChiQuay, dataItemLastYear.DSChiNha, dataItemLastYear.DSCK, dataItemLastYear.TongDS
                        , sumDSChiQuay, sumDSChiNha, sumDSCK, sumTongDS);
                }

                DataRow row = dataTable.NewRow();
                row["MarketName"] = "Tổng";
                row["CQ1"] = dataTable.Compute("Sum(CQ1)", "");
                row["CN1"] = dataTable.Compute("Sum(CN1)", "");
                row["CK1"] = dataTable.Compute("Sum(CK1)", "");
                row["TDS1"] = dataTable.Compute("Sum(TDS1)", "");

                row["CQ2"] = dataTable.Compute("Sum(CQ2)", "");
                row["CN2"] = dataTable.Compute("Sum(CN2)", "");
                row["CK2"] = dataTable.Compute("Sum(CK2)", "");
                row["TDS2"] = dataTable.Compute("Sum(TDS2)", "");

                row["CQ3"] = dataTable.Compute("Sum(CQ3)", "");
                row["CN3"] = dataTable.Compute("Sum(CN3)", "");
                row["CK3"] = dataTable.Compute("Sum(CK3)", "");
                row["TDS3"] = dataTable.Compute("Sum(TDS3)", "");
                dataTable.Rows.Add(row);


            }

            // Set border
            Style style = new CellsFactory().CreateStyle();
            style.SetBorder(BorderType.BottomBorder, CellBorderType.Thin, Color.Black);
            style.SetBorder(BorderType.LeftBorder, CellBorderType.Thin, Color.Black);
            style.SetBorder(BorderType.RightBorder, CellBorderType.Thin, Color.Black);
            style.SetBorder(BorderType.TopBorder, CellBorderType.Thin, Color.Black);

            if (dataTable.Rows.Count > 0)
            {
                int stepRow = 0;
                // total row = row start + số row hiện có
                int totalRow = dataTable.Rows.Count + 62;
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
                        string valueOfTable = dataTable.Rows[stepRow][stepColumn].ToString();

                        // Tô màu cho các dòng có giá trị tăng giảm
                        if (b >= 10)
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

            // vẻ biểu đồ cột
            // Create Chart Line
            //Chart reference
            Aspose.Cells.Charts.Chart leadSourceColumn;
            //Add Pie Chart
            int chartIndex = sheetReport.Charts.Add(ChartType.Column3DClustered, 6, 0, 30, 13);
            leadSourceColumn = sheetReport.Charts[chartIndex];


            //Chart title
            leadSourceColumn.Title.Text = string.Format("Doanh số từng dịch vụ từng thị trường \n Giai đoạn: {0}", text);
            leadSourceColumn.Title.Font.Color = Color.Silver;

            // Set width cho column
            sheetReport.Cells.SetColumnWidthPixel(15, 220);

            // List loại chi trả
            string[] str = {string.Format("Chi Quầy {0}", year), string.Format("Chi Nhà {0}", year), string.Format("Chuyển khoản {0}", year)
                    , string.Format("Chi Quầy {0}", year - 1), string.Format("Chi Nhà {0}", year - 1), string.Format("Chuyển khoản {0}", year - 1)};

            int count = 0;
            foreach (ReportDetailtServiceType item in listReportDataClone)
            {
                
                // Loại bỏ sự lập lại của Market name
                if (count < 6)
                {
                    //item.TongDS = item.DSChiQuay + item.DSChiNha + item.DSCK;

                    // Cùng kì
                    ReportDetailtServiceType dataItemLastYear = listReportData.Find(x => x.MarketCode == item.MarketCode && x.Year == (year - 1).ToString());
                    ReportDetailtServiceType dataItemYear = listReportData.Find(x => x.MarketCode == item.MarketCode && x.Year == year.ToString());

                    // Trường hợp năm không có đối tác
                    if (dataItemLastYear == null && dataItemYear != null)
                    {
                        dataItemLastYear = new ReportDetailtServiceType();
                        dataItemLastYear.MarketName = dataItemYear.MarketName;
                        dataItemLastYear.DSChiQuay = 0;
                        dataItemLastYear.DSChiNha = 0;
                        dataItemLastYear.DSCK = 0;
                        dataItemLastYear.Year = (year - 1).ToString();
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
                    }

                    string totalRowData = string.Concat("{", string.Format("{0}, {1}, {2}, {3}, {4}, {5}", dataItemYear.DSChiQuay, dataItemYear.DSChiNha, dataItemYear.DSCK, dataItemLastYear.DSChiQuay, dataItemLastYear.DSChiNha, dataItemLastYear.DSCK), "}");

                    leadSourceColumn.NSeries.Add(totalRowData, true);

                    string categoryData = string.Concat("{", string.Join(", ", str), "}");
                    leadSourceColumn.NSeries.CategoryData = categoryData;

                    leadSourceColumn.NSeries[count].Name = dataItemYear.MarketName;

                    switch (count)
                    {
                        case 1:
                            // Set the 1st series fill color.
                            leadSourceColumn.NSeries[count].Area.ForegroundColor = Color.Orange;
                            leadSourceColumn.NSeries[count].Area.Formatting = FormattingType.Custom;
                            break;
                        case 2:
                            // Set the 1st series fill color.
                            leadSourceColumn.NSeries[count].Area.ForegroundColor = Color.Green;
                            leadSourceColumn.NSeries[count].Area.Formatting = FormattingType.Custom;
                            break;
                        case 3:
                            // Set the 1st series fill color.
                            leadSourceColumn.NSeries[count].Area.ForegroundColor = Color.Blue;
                            leadSourceColumn.NSeries[count].Area.Formatting = FormattingType.Custom;
                            break;
                        case 4:
                            // Set the 1st series fill color.
                            leadSourceColumn.NSeries[count].Area.ForegroundColor = Color.Red;
                            leadSourceColumn.NSeries[count].Area.Formatting = FormattingType.Custom;
                            break;
                        case 5:
                            // Set the 1st series fill color.
                            leadSourceColumn.NSeries[count].Area.ForegroundColor = Color.Silver;
                            leadSourceColumn.NSeries[count].Area.Formatting = FormattingType.Custom;
                            break;
                        default:
                            // Set the 1st series fill color.
                            leadSourceColumn.NSeries[count].Area.ForegroundColor = Color.Pink;
                            leadSourceColumn.NSeries[count].Area.Formatting = FormattingType.Custom;
                            break;
                    }

                    count++;
                }
            }
            
            // Set plot area formatting as none and hide its border.
            leadSourceColumn.PlotArea.Area.FillFormat.FillType = FillType.None;
            leadSourceColumn.PlotArea.Border.IsVisible = false;


            // vẻ biểu đồ cột Stack
            // Create Chart Line
            //Chart reference
            Aspose.Cells.Charts.Chart leadSourceColumnStack;
            //Add Pie Chart
            int chartIndexStack = sheetReport.Charts.Add(ChartType.Column3D100PercentStacked, 32, 0, 56, 13);
            leadSourceColumnStack = sheetReport.Charts[chartIndexStack];


            //Chart title
            leadSourceColumnStack.Title.Text = string.Format("Doanh số từng dịch vụ từng thị trường \n Giai đoạn: {0}", text);
            leadSourceColumnStack.Title.Font.Color = Color.Silver;

            // Set width cho column
            sheetReport.Cells.SetColumnWidthPixel(15, 220);

            // List loại chi trả
            string[] strStack = {string.Format("Chi Quầy {0}", year), string.Format("Chi Nhà {0}", year), string.Format("Chuyển khoản {0}", year)
                    , string.Format("Chi Quầy {0}", year - 1), string.Format("Chi Nhà {0}", year - 1), string.Format("Chuyển khoản {0}", year - 1)};

            count = 0;

            List<ReportDetailtServiceType> resultConvert = new List<ReportDetailtServiceType>();

            if (listReportDataClone.Count > 0)
            {
                double sumDSChiQuayYear = 0;
                double sumDSChiNhaYear = 0;
                double sumDSCKYear = 0;

                double sumDSChiQuayLastYear = 0;
                double sumDSChiNhaLastYear = 0;
                double sumDSCKLastYear = 0;

                if (listReportDataClone.Count > 0)
                {
                    // Last Year
                    sumDSChiQuayLastYear = listReportDataClone.Where(x => x.Year == (year - 1).ToString()).Sum(y => y.DSChiQuay);
                    sumDSChiNhaLastYear = listReportDataClone.Where(x => x.Year == (year - 1).ToString()).Sum(y => y.DSChiNha);
                    sumDSCKLastYear = listReportDataClone.Where(x => x.Year == (year - 1).ToString()).Sum(y => y.DSCK);

                    // Year hiện tại
                    sumDSChiQuayYear = listReportDataClone.Where(x => x.Year == year.ToString()).Sum(y => y.DSChiQuay);
                    sumDSChiNhaYear = listReportDataClone.Where(x => x.Year == year.ToString()).Sum(y => y.DSChiNha);
                    sumDSCKYear = listReportDataClone.Where(x => x.Year == year.ToString()).Sum(y => y.DSCK);
                }

                foreach (ReportDetailtServiceType item in listReportDataClone)
                {
                    item.TongDS = item.DSChiQuay + item.DSChiNha + item.DSCK;
                    ReportDetailtServiceType itemDetailtPercent = new ReportDetailtServiceType();

                    // Last year
                    if (item.Year == (year - 1).ToString())
                    {
                        itemDetailtPercent = new ReportDetailtServiceType()
                        {
                            MarketCode = item.MarketCode,
                            MarketName = item.MarketName,
                            DSChiQuay = item.TongDS == 0 ? 0 : Math.Round(item.DSChiQuay / sumDSChiQuayLastYear * 100, 2, MidpointRounding.ToEven),
                            DSChiNha = item.TongDS == 0 ? 0 : Math.Round(item.DSChiNha / sumDSChiNhaLastYear * 100, 2, MidpointRounding.ToEven),
                            DSCK = item.TongDS == 0 ? 0 : Math.Round(item.DSCK / sumDSCKLastYear * 100, 2, MidpointRounding.ToEven),
                            Year = item.Year
                        };
                    }

                    // year hien tai
                    if (item.Year == year.ToString())
                    {
                        itemDetailtPercent = new ReportDetailtServiceType()
                        {
                            MarketCode = item.MarketCode,
                            MarketName = item.MarketName,
                            DSChiQuay = item.TongDS == 0 ? 0 : Math.Round(item.DSChiQuay / sumDSChiQuayYear * 100, 2, MidpointRounding.ToEven),
                            DSChiNha = item.TongDS == 0 ? 0 : Math.Round(item.DSChiNha / sumDSChiNhaYear * 100, 2, MidpointRounding.ToEven),
                            DSCK = item.TongDS == 0 ? 0 : Math.Round(item.DSCK / sumDSCKYear * 100, 2, MidpointRounding.ToEven),
                            Year = item.Year
                        };
                    }
                    resultConvert.Add(itemDetailtPercent);
                }
            }

            foreach (ReportDetailtServiceType item in resultConvert)
            {

                // Loại bỏ sự lập lại của Market name
                if (count < 6)
                {
                    //item.TongDS = item.DSChiQuay + item.DSChiNha + item.DSCK;

                    // Cùng kì
                    ReportDetailtServiceType dataItemLastYear = listReportData.Find(x => x.MarketCode == item.MarketCode && x.Year == (year - 1).ToString());
                    ReportDetailtServiceType dataItemYear = listReportData.Find(x => x.MarketCode == item.MarketCode && x.Year == year.ToString());

                    // Trường hợp năm không có đối tác
                    if (dataItemLastYear == null && dataItemYear != null)
                    {
                        dataItemLastYear = new ReportDetailtServiceType();
                        dataItemLastYear.MarketName = dataItemYear.MarketName;
                        dataItemLastYear.DSChiQuay = 0;
                        dataItemLastYear.DSChiNha = 0;
                        dataItemLastYear.DSCK = 0;
                        dataItemLastYear.Year = (year - 1).ToString();
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
                    }

                    string totalRowData = string.Concat("{", string.Format("{0}, {1}, {2}, {3}, {4}, {5}", dataItemYear.DSChiQuay, dataItemYear.DSChiNha, dataItemYear.DSCK, dataItemLastYear.DSChiQuay, dataItemLastYear.DSChiNha, dataItemLastYear.DSCK), "}");

                    leadSourceColumnStack.NSeries.Add(totalRowData, true);

                    string categoryData = string.Concat("{", string.Join(", ", strStack), "}");
                    leadSourceColumnStack.NSeries.CategoryData = categoryData;

                    leadSourceColumnStack.NSeries[count].Name = dataItemYear.MarketName;

                    switch (count)
                    {
                        case 1:
                            // Set the 1st series fill color.
                            leadSourceColumnStack.NSeries[count].Area.ForegroundColor = Color.Orange;
                            leadSourceColumnStack.NSeries[count].Area.Formatting = FormattingType.Custom;
                            break;
                        case 2:
                            // Set the 1st series fill color.
                            leadSourceColumnStack.NSeries[count].Area.ForegroundColor = Color.Green;
                            leadSourceColumnStack.NSeries[count].Area.Formatting = FormattingType.Custom;
                            break;
                        case 3:
                            // Set the 1st series fill color.
                            leadSourceColumnStack.NSeries[count].Area.ForegroundColor = Color.Blue;
                            leadSourceColumnStack.NSeries[count].Area.Formatting = FormattingType.Custom;
                            break;
                        case 4:
                            // Set the 1st series fill color.
                            leadSourceColumnStack.NSeries[count].Area.ForegroundColor = Color.Red;
                            leadSourceColumnStack.NSeries[count].Area.Formatting = FormattingType.Custom;
                            break;
                        case 5:
                            // Set the 1st series fill color.
                            leadSourceColumnStack.NSeries[count].Area.ForegroundColor = Color.Silver;
                            leadSourceColumnStack.NSeries[count].Area.Formatting = FormattingType.Custom;
                            break;
                        default:
                            // Set the 1st series fill color.
                            leadSourceColumnStack.NSeries[count].Area.ForegroundColor = Color.Pink;
                            leadSourceColumnStack.NSeries[count].Area.Formatting = FormattingType.Custom;
                            break;
                    }

                    count++;
                }
            }


            // Set plot area formatting as none and hide its border.
            leadSourceColumnStack.PlotArea.Area.FillFormat.FillType = FillType.None;
            leadSourceColumnStack.PlotArea.Border.IsVisible = false;

            //sheetReport.Cells.ConvertStringToNumericValue();
            // Chạy process
            designer.Process();
            return ExportReport("ReportDetailMarketForGradation", designer);

        }

        /// <summary>
        /// Tạo mẫu cho Excel cho so sánh theo giai đoạn
        /// </summary>
        /// <param name="gradationID"></param>
        /// <param name="year"></param>
        /// <param name="typeID"></param>
        /// <returns></returns>
        public ActionResult CreateExcelGradationCompareForOne(string gradationID, int year, string reportTypeID, string marketID)
        {
            WorkbookDesigner designer = new WorkbookDesigner();
            string templatePath = "~/Content/Report/ReportDetailMarketGarationForOne.xlsx";
            // Get đường dẫn
            templatePath = System.Web.HttpContext.Current.Server.MapPath(templatePath);

            designer.Workbook = new DongA.Core.DongAExcel().OpenExcelFile(templatePath);
            designer.Workbook.CalculateFormula();

            WorksheetCollection workSheets = designer.Workbook.Worksheets;
            Worksheet sheetReport = designer.Workbook.Worksheets[0];


            // Tạo title
            string typeReport = "So sánh - Theo giai đoạn - Từng thị trường";

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
            CreateTitle("A2", "K2", sheetReport, typeReport, 14);

            string stringMarket = string.Empty;

            switch (marketID)
            {
                case "001":
                    stringMarket = "Thị trường: Châu Âu";
                    break;
                case "002":
                    stringMarket = "Thị trường: Mỹ";
                    break;
                case "003":
                    stringMarket = "Thị trường: Canada";
                    break;
                case "004":
                    stringMarket = "Thị trường: Úc";
                    break;
                case "005":
                    stringMarket = "Thị trường: Châu Á";
                    break;
                default:
                    stringMarket = "Thị trường: Toàn cầu";
                    break;
            }
            // Tạo title
            CreateTitle("A3", "K3", sheetReport, stringMarket, 14);
            // Tạo title detailt
            string titleDetailt = string.Format("Giai đoạn: {0}", text);
            CreateTitle("A4", "K4", sheetReport, titleDetailt, 12);

            
            List<ReportDetailtServiceType> listReportData = new ReportBL().ReportDetailtGradationCompareForOne(year, int.Parse(gradationID), reportTypeID, marketID);
            // clone Object
            List<ReportDetailtServiceType> listReportDataClone = new List<ReportDetailtServiceType>(listReportData);

            foreach (ReportDetailtServiceType item in listReportData)
            {
                item.TongDS = item.DSChiQuay + item.DSChiNha + item.DSCK;
            }

            DataTable dataTable = new DataTable();

            // Tạo các cột cho datatable
            dataTable.Columns.Add("PartnerName", typeof(String));
            dataTable.Columns.Add("CQ1", typeof(double));
            dataTable.Columns.Add("CN1", typeof(double));
            dataTable.Columns.Add("CK1", typeof(double));
            dataTable.Columns.Add("TDS1", typeof(double));

            dataTable.Columns.Add("CQ2", typeof(double));
            dataTable.Columns.Add("CN2", typeof(double));
            dataTable.Columns.Add("CK2", typeof(double));
            dataTable.Columns.Add("TDS2", typeof(double));

            dataTable.Columns.Add("CQ3", typeof(double));
            dataTable.Columns.Add("CN3", typeof(double));
            dataTable.Columns.Add("CK3", typeof(double));
            dataTable.Columns.Add("TDS3", typeof(double));

            foreach (ReportDetailtServiceType item in listReportDataClone)
            {
                // Cùng kì
                ReportDetailtServiceType dataItemLastYear = listReportDataClone.Find(x => x.PartnerCode == item.PartnerCode && x.Year == (year - 1).ToString());
                ReportDetailtServiceType dataItemYear = listReportDataClone.Find(x => x.PartnerCode == item.PartnerCode && x.Year == year.ToString());

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

                double sumDSChiQuay = dataItemYear.DSChiQuay - dataItemLastYear.DSChiQuay;
                double sumDSChiNha = dataItemYear.DSChiNha - dataItemLastYear.DSChiNha;
                double sumDSCK = dataItemYear.DSCK - dataItemLastYear.DSCK;

                double sumTongDS = dataItemYear.TongDS - dataItemLastYear.TongDS;


                // Check tồn tại của item
                string value = string.Format("PartnerName='{0}'", item.PartnerName);
                DataRow[] foundRows = dataTable.Select(value);
                if (dataItemLastYear != null && dataItemYear != null && foundRows.Count() == 0)
                {
                    // add item vào table
                    dataTable.Rows.Add(dataItemLastYear.PartnerName, dataItemYear.DSChiQuay, dataItemYear.DSChiNha, dataItemYear.DSCK, dataItemYear.TongDS
                        , dataItemLastYear.DSChiQuay, dataItemLastYear.DSChiNha, dataItemLastYear.DSCK, dataItemLastYear.TongDS
                        , sumDSChiQuay, sumDSChiNha, sumDSCK, sumTongDS);
                }
            }

            DataRow row = dataTable.NewRow();
            row["PartnerName"] = "Tổng";
            row["CQ1"] = dataTable.Compute("Sum(CQ1)", "");
            row["CN1"] = dataTable.Compute("Sum(CN1)", "");
            row["CK1"] = dataTable.Compute("Sum(CK1)", "");
            row["TDS1"] = dataTable.Compute("Sum(TDS1)", "");

            row["CQ2"] = dataTable.Compute("Sum(CQ2)", "");
            row["CN2"] = dataTable.Compute("Sum(CN2)", "");
            row["CK2"] = dataTable.Compute("Sum(CK2)", "");
            row["TDS2"] = dataTable.Compute("Sum(TDS2)", "");

            row["CQ3"] = dataTable.Compute("Sum(CQ3)", "");
            row["CN3"] = dataTable.Compute("Sum(CN3)", "");
            row["CK3"] = dataTable.Compute("Sum(CK3)", "");
            row["TDS3"] = dataTable.Compute("Sum(TDS3)", "");
            dataTable.Rows.Add(row);

            // Set border
            Style style = new CellsFactory().CreateStyle();
            style.SetBorder(BorderType.BottomBorder, CellBorderType.Thin, Color.Black);
            style.SetBorder(BorderType.LeftBorder, CellBorderType.Thin, Color.Black);
            style.SetBorder(BorderType.RightBorder, CellBorderType.Thin, Color.Black);
            style.SetBorder(BorderType.TopBorder, CellBorderType.Thin, Color.Black);

            if (dataTable.Rows.Count > 0)
            {
                int stepRow = 0;
                // total row = row start + số row hiện có
                int totalRow = dataTable.Rows.Count + 62;
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
                        string valueOfTable = dataTable.Rows[stepRow][stepColumn].ToString();

                        // Tô màu cho các dòng có giá trị tăng giảm
                        if (b >= 10)
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

            // vẻ biểu đồ cột
            // Create Chart Line
            //Chart reference
            Aspose.Cells.Charts.Chart leadSourceColumn;
            //Add Pie Chart
            int chartIndex = sheetReport.Charts.Add(ChartType.Column3DClustered, 6, 0, 30, 13);
            leadSourceColumn = sheetReport.Charts[chartIndex];


            //Chart title
            leadSourceColumn.Title.Text = string.Format("Doanh số từng dịch vụ từng thị trường \n Giai đoạn: {0}", text);
            leadSourceColumn.Title.Font.Color = Color.Black;

            // Set width cho column
            sheetReport.Cells.SetColumnWidthPixel(15, 220);

            // List loại chi trả
            string[] str = {string.Format("Chi Quầy {0}", year), string.Format("Chi Nhà {0}", year), string.Format("Chuyển khoản {0}", year)
                    , string.Format("Chi Quầy {0}", year - 1), string.Format("Chi Nhà {0}", year - 1), string.Format("Chuyển khoản {0}", year - 1)};

            List<string> Listdulicate = new List<string>();
            int count = 0;
            foreach (ReportDetailtServiceType item in listReportDataClone)
            {
                // pass qua nếu đã tồn tại trong list
                if (Listdulicate.Contains(item.PartnerCode))
                {
                    continue;
                }

                // Add item vào Listdulicate
                Listdulicate.Add(item.PartnerCode);

                // Cùng kì
                ReportDetailtServiceType dataItemLastYear = listReportData.Find(x => x.PartnerCode == item.PartnerCode && x.Year == (year - 1).ToString());
                ReportDetailtServiceType dataItemYear = listReportData.Find(x => x.PartnerCode == item.PartnerCode && x.Year == year.ToString());

                // Trường hợp năm không có đối tác
                if (dataItemLastYear == null && dataItemYear != null)
                {
                    dataItemLastYear = new ReportDetailtServiceType();
                    dataItemLastYear.PartnerName = dataItemYear.PartnerName;
                    dataItemLastYear.DSChiQuay = 0;
                    dataItemLastYear.DSChiNha = 0;
                    dataItemLastYear.DSCK = 0;
                    dataItemLastYear.Year = (year - 1).ToString();
                }

                // Trường hợp năm hiện tại không có đối tác
                if (dataItemYear == null && dataItemLastYear != null)
                {
                    dataItemYear = new ReportDetailtServiceType();
                    dataItemYear.PartnerName = dataItemLastYear.PartnerName;
                    dataItemYear.DSChiQuay = 0;
                    dataItemYear.DSChiNha = 0;
                    dataItemYear.DSCK = 0;
                    dataItemYear.Year = year.ToString();
                }

                string totalRowData = string.Concat("{", string.Format("{0}, {1}, {2}, {3}, {4}, {5}", dataItemYear.DSChiQuay, dataItemYear.DSChiNha, dataItemYear.DSCK, dataItemLastYear.DSChiQuay, dataItemLastYear.DSChiNha, dataItemLastYear.DSCK), "}");

                leadSourceColumn.NSeries.Add(totalRowData, true);

                string categoryData = string.Concat("{", string.Join(", ", str), "}");
                leadSourceColumn.NSeries.CategoryData = categoryData;

                leadSourceColumn.NSeries[count].Name = dataItemYear.PartnerName;

                switch (count)
                {
                    case 1:
                        // Set the 1st series fill color.
                        leadSourceColumn.NSeries[count].Area.ForegroundColor = Color.Orange;
                        leadSourceColumn.NSeries[count].Area.Formatting = FormattingType.Custom;
                        break;
                    case 2:
                        // Set the 1st series fill color.
                        leadSourceColumn.NSeries[count].Area.ForegroundColor = Color.Green;
                        leadSourceColumn.NSeries[count].Area.Formatting = FormattingType.Custom;
                        break;
                    case 3:
                        // Set the 1st series fill color.
                        leadSourceColumn.NSeries[count].Area.ForegroundColor = Color.Blue;
                        leadSourceColumn.NSeries[count].Area.Formatting = FormattingType.Custom;
                        break;
                    case 4:
                        // Set the 1st series fill color.
                        leadSourceColumn.NSeries[count].Area.ForegroundColor = Color.Red;
                        leadSourceColumn.NSeries[count].Area.Formatting = FormattingType.Custom;
                        break;
                    case 5:
                        // Set the 1st series fill color.
                        leadSourceColumn.NSeries[count].Area.ForegroundColor = Color.Silver;
                        leadSourceColumn.NSeries[count].Area.Formatting = FormattingType.Custom;
                        break;
                    default:
                        // Set the 1st series fill color.
                        leadSourceColumn.NSeries[count].Area.ForegroundColor = Color.Pink;
                        leadSourceColumn.NSeries[count].Area.Formatting = FormattingType.Custom;
                        break;
                }

                count++;
            }

            // Set plot area formatting as none and hide its border.
            leadSourceColumn.PlotArea.Area.FillFormat.FillType = FillType.None;
            leadSourceColumn.PlotArea.Border.IsVisible = false;

            // vẻ biểu đồ cột Stack
            // Create Chart Line
            //Chart reference
            Aspose.Cells.Charts.Chart leadSourceColumnStack;
            //Add Pie Chart
            int chartIndexStack = sheetReport.Charts.Add(ChartType.Column3DStacked, 32, 0, 56, 13);
            leadSourceColumnStack = sheetReport.Charts[chartIndexStack];


            //Chart title
            leadSourceColumnStack.Title.Text = string.Format("Doanh số từng dịch vụ từng thị trường \n Giai đoạn: {0}", text);
            leadSourceColumnStack.Title.Font.Color = Color.Silver;

            // Set width cho column
            sheetReport.Cells.SetColumnWidthPixel(15, 220);

            // List loại chi trả
            List<string> strStack = new List<string>();

            count = 0;

            // Get list mới lưu trữ
            Listdulicate = new List<string>();
            List<string> listTongDSYear = new List<string>();
            List<string> listTongDSLastYear = new List<string>();

            foreach (ReportDetailtServiceType item in listReportDataClone)
            {
                // pass qua nếu đã tồn tại trong list
                if (Listdulicate.Contains(item.PartnerCode))
                {
                    continue;
                }

                // Cùng kì
                ReportDetailtServiceType dataItemLastYear = listReportData.Find(x => x.PartnerCode == item.PartnerCode && x.Year == (year - 1).ToString());
                ReportDetailtServiceType dataItemYear = listReportData.Find(x => x.PartnerCode == item.PartnerCode && x.Year == year.ToString());

                // Trường hợp năm không có đối tác
                if (dataItemLastYear == null && dataItemYear != null)
                {
                    dataItemLastYear = new ReportDetailtServiceType();
                    dataItemLastYear.PartnerName = dataItemYear.PartnerName;
                    dataItemLastYear.DSChiQuay = 0;
                    dataItemLastYear.DSChiNha = 0;
                    dataItemLastYear.DSCK = 0;
                    dataItemLastYear.TongDS = 0;
                    dataItemLastYear.Year = (year - 1).ToString();
                }

                // Trường hợp năm hiện tại không có đối tác
                if (dataItemYear == null && dataItemLastYear != null)
                {
                    dataItemYear = new ReportDetailtServiceType();
                    dataItemYear.PartnerName = dataItemLastYear.PartnerName;
                    dataItemYear.DSChiQuay = 0;
                    dataItemYear.DSChiNha = 0;
                    dataItemYear.DSCK = 0;
                    dataItemYear.TongDS = 0;
                    dataItemYear.Year = year.ToString();
                }

                listTongDSYear.Add(dataItemYear.TongDS.ToString());
                listTongDSLastYear.Add(dataItemLastYear.TongDS.ToString());

                // Add item vào Listdulicate
                Listdulicate.Add(item.PartnerCode);
                item.TongDS = item.DSChiQuay + item.DSChiNha + item.DSCK;

                strStack.Add(item.PartnerName);
            }

            // Get list mới lưu trữ
            Listdulicate = new List<string>();

            // chạy 2 lần
            for (int i = 0; i< 2;i++)
            {
                if (i == 0)
                {
                    string totalRowData = string.Concat("{", string.Join(", ", listTongDSYear), "}");

                    leadSourceColumnStack.NSeries.Add(totalRowData, true);
                }
                else
                {
                    string totalRowData = string.Concat("{", string.Join(", ", listTongDSLastYear), "}");

                    leadSourceColumnStack.NSeries.Add(totalRowData, true);
                }


                string categoryDataStack = string.Concat("{", string.Join(", ", strStack), "}");
                leadSourceColumnStack.NSeries.CategoryData = categoryDataStack;

            }

            leadSourceColumnStack.NSeries[0].Name = string.Format("Tổng 3T {0}", year);
            leadSourceColumnStack.NSeries[1].Name = string.Format("Tổng 3T {0}", year - 1);

            // Set plot area formatting as none and hide its border.
            leadSourceColumnStack.PlotArea.Area.FillFormat.FillType = FillType.None;
            leadSourceColumnStack.PlotArea.Border.IsVisible = false;

            typeReport = "2. Theo tỷ trọng doanh số theo thị trường - loại hình chi trả";
            CreateTitle("B73", "F73", sheetReport, typeReport, 14);

            List<ReportDetailtServiceType> listDataGradationConvert = new List<ReportDetailtServiceType>(listReportData);
            // clone Object
            List<ReportDetailtServiceType> listDataGradationClone = new List<ReportDetailtServiceType>(listDataGradationConvert);
            foreach (ReportDetailtServiceType item in listDataGradationClone)
            {
                // Get dữ liệu của năm hiện tại
                ReportDetailtServiceType listDataYear = listDataGradationConvert.Find(x => x.Year == year.ToString() && x.PartnerCode == item.PartnerCode);
                // Get dữ liệu của năm trước
                ReportDetailtServiceType listDataLastYear = listDataGradationConvert.Find(x => x.Year == (year - 1).ToString() && x.PartnerCode == item.PartnerCode);

                // Trường hợp năm không có đối tác
                if (listDataLastYear == null && listDataYear != null)
                {
                    listDataLastYear = new ReportDetailtServiceType();
                    listDataLastYear.PartnerCode = listDataYear.PartnerCode;
                    listDataLastYear.PartnerName = listDataYear.PartnerName;
                    listDataLastYear.DSChiQuay = 0;
                    listDataLastYear.DSChiNha = 0;
                    listDataLastYear.DSCK = 0;
                    listDataLastYear.TongDS = 0;
                    listDataLastYear.Year = (year - 1).ToString();
                    listDataGradationConvert.Add(listDataLastYear);
                }

                // Trường hợp năm hiện tại không có đối tác
                if (listDataYear == null && listDataLastYear != null)
                {
                    listDataYear = new ReportDetailtServiceType();
                    listDataYear.PartnerCode = listDataLastYear.PartnerCode;
                    listDataYear.PartnerName = listDataLastYear.PartnerName;
                    listDataYear.DSChiQuay = 0;
                    listDataYear.DSChiNha = 0;
                    listDataYear.DSCK = 0;
                    listDataYear.DSCK = 0;
                    listDataYear.TongDS = 0;
                    listDataYear.Year = year.ToString();
                    listDataGradationConvert.Add(listDataYear);
                }
            }

            foreach (ReportDetailtServiceType item in listDataGradationConvert)
            {
                item.TongDS = item.DSChiQuay + item.DSChiNha + item.DSCK;
            }

            double sumTongDSYear = listDataGradationConvert.Where(x => x.Year == year.ToString()).Sum(x => x.TongDS);
            double sumTongDSLastYear = listDataGradationConvert.Where(x => x.Year == (year - 1).ToString()).Sum(x => x.TongDS);

            List<ReportDetailtServiceType> listDataConvert = new List<ReportDetailtServiceType>();

            // Khởi tạo datatable
            DataTable table = new DataTable();
            // Tạo các cột cho datatable
            table.Columns.Add("STT", typeof(String));
            table.Columns.Add("PartnerName", typeof(String));
            table.Columns.Add("LK1", typeof(double));
            table.Columns.Add("LK2", typeof(double));

            count = 1;
            foreach (ReportDetailtServiceType item in listDataGradationConvert)
            {
                // Get dữ liệu của năm hiện tại
                ReportDetailtServiceType listDataYear = listDataGradationConvert.Find(x => x.Year == year.ToString() && x.PartnerCode == item.PartnerCode);
                // Get dữ liệu của năm trước
                ReportDetailtServiceType listDataLastYear = listDataGradationConvert.Find(x => x.Year == (year - 1).ToString() && x.PartnerCode == item.PartnerCode);

                // Check tồn tại của item
                string value = string.Format("PartnerName='{0}'", item.PartnerName);
                DataRow[] foundRows = table.Select(value);

                if (listDataYear != null && listDataLastYear != null && foundRows.Count() == 0)
                {
                    double valueYear = Math.Round((listDataYear.TongDS / sumTongDSYear) * 100, 2, MidpointRounding.ToEven);
                    double valueLastYear = Math.Round((listDataLastYear.TongDS / sumTongDSLastYear) * 100, 2, MidpointRounding.ToEven);
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

            // Tạo bảng dữ liệu
            int countRowPercent = (dataTable.Rows.Count + 62) + 25;

            if (table.Rows.Count > 0)
            {
                int stepRow = 0;
                // total row = row start + số row hiện có
                int totalRow = table.Rows.Count + countRowPercent;

                // Add header cho bảng dữ liệu phần trăm
                sheetReport.Cells[countRowPercent - 1, 0].PutValue("STT", true, true);
                style.Font.IsBold = true;
                sheetReport.Cells[countRowPercent - 1, 0].SetStyle(style);

                sheetReport.Cells[countRowPercent - 1, 1].PutValue("Đối tác", true, true);
                style.Font.IsBold = true;
                sheetReport.Cells[countRowPercent - 1, 1].SetStyle(style);

                sheetReport.Cells[countRowPercent - 1, 2].PutValue(string.Format("Lũy kế {0} {1}", text, year), true, true);
                style.Font.IsBold = true;
                sheetReport.Cells[countRowPercent - 1, 2].SetStyle(style);

                sheetReport.Cells[countRowPercent - 1, 3].PutValue(string.Format("Lũy kế {0} {1}", text, year - 1), true, true);
                style.Font.IsBold = true;
                sheetReport.Cells[countRowPercent - 1, 3].SetStyle(style);

                // Số dòng của row
                for (int a = countRowPercent; a < totalRow; a++)
                {
                    int stepColumn = 0;
                    // Số cột trong báo cáo cần hiển thị
                    // Tổng số cột hiển thị = Số cột hiển thị bắt đầu + tổng số cột cần hiển thị
                    int totalCol = 0 + 4;
                    for (int b = 0; b < totalCol; b++)
                    {
                        // Giá trị của value trong table
                        string valueOfTable = table.Rows[stepRow][stepColumn].ToString();

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

            List<string> dataPieYear = new List<string>();
            List<string> dataPieLastYear = new List<string>();
            List<string> dataPartner = new List<string>();
            count = 0;
            foreach (DataRow dtRow in table.Rows)
            {
                if (count < table.Rows.Count - 1)
                {
                    // On all tables' columns
                    foreach (DataColumn dc in table.Columns)
                    {
                        if (dc.ColumnName.Contains("LK1"))
                        {
                            dataPieYear.Add(dtRow[dc].ToString());
                        }

                        if (dc.ColumnName.Contains("LK2"))
                        {
                            dataPieLastYear.Add(dtRow[dc].ToString());
                        }

                        if (dc.ColumnName.Contains("PartnerName"))
                        {
                            dataPartner.Add(dtRow[dc].ToString());
                        }
                    }
                }

                count++;
            }
            
            countRowPercent = (dataTable.Rows.Count + 62) + 4;

            if (dataPieYear != null)
            {
                //Add Pie Chart
                Aspose.Cells.Charts.Chart leadSourcePie;
                chartIndex = sheetReport.Charts.Add(ChartType.Pie3D, countRowPercent, 0, countRowPercent + 17, 5);
                leadSourcePie = sheetReport.Charts[chartIndex];

                // Set some properties of chart plot area.
                // To set the fill color and make the border invisible.
                leadSourcePie.PlotArea.Border.IsVisible = false;
                leadSourcePie.Elevation = 45;
                // Set properties of chart title
                leadSourcePie.Title.Text = string.Concat("Lũy kế ", text, " ", year);
                leadSourcePie.Title.Font.Color = Color.Blue;
                leadSourcePie.Title.Font.IsBold = true;
                leadSourcePie.Title.Font.Size = 12;

                // Set properties of nseries
                string totalRowData = string.Concat("{", string.Join(", ", dataPieYear), "}");
                leadSourcePie.NSeries.Add(totalRowData, true);

                string categoryData = string.Concat("{", string.Join(", ", dataPartner), "}");
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
                Aspose.Cells.Charts.Chart leadSourcePie;
                chartIndex = sheetReport.Charts.Add(ChartType.Pie3D, countRowPercent, 6, countRowPercent + 17, 11);
                leadSourcePie = sheetReport.Charts[chartIndex];

                // Set some properties of chart plot area.
                // To set the fill color and make the border invisible.
                leadSourcePie.PlotArea.Border.IsVisible = false;
                leadSourcePie.Elevation = 45;
                // Set properties of chart title
                leadSourcePie.Title.Text = string.Concat("Lũy kế ", text, " ", year - 1);
                leadSourcePie.Title.Font.Color = Color.Blue;
                leadSourcePie.Title.Font.IsBold = true;
                leadSourcePie.Title.Font.Size = 12;

                // Set properties of nseries
                string totalRowData = string.Concat("{", string.Join(", ", dataPieLastYear), "}");
                leadSourcePie.NSeries.Add(totalRowData, true);

                string categoryData = string.Concat("{", string.Join(", ", dataPartner), "}");
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


            // Chạy process
            designer.Process();
            return ExportReport("ReportDetailMarketForGradation", designer);
        }


        /// <summary>
        /// Tạo mẫu cho Excel cho so sánh theo tháng
        /// </summary>
        /// <param name="gradationID"></param>
        /// <param name="year"></param>
        /// <param name="typeID"></param>
        /// <returns></returns>
        public ActionResult CreateExcelCompareForAll(int month, int year, string reportTypeID)
        {
            WorkbookDesigner designer = new WorkbookDesigner();
            string templatePath = "~/Content/Report/ReportDetailMarketForCompare.xlsx";
            // Get đường dẫn
            templatePath = System.Web.HttpContext.Current.Server.MapPath(templatePath);

            designer.Workbook = new DongA.Core.DongAExcel().OpenExcelFile(templatePath);
            designer.Workbook.CalculateFormula();

            WorksheetCollection workSheets = designer.Workbook.Worksheets;
            Worksheet sheetReport = designer.Workbook.Worksheets[0];

            // Tạo title
            string typeReport = "So sánh - Theo tháng - Tất cả thị trường";
            
            // Tạo title
            CreateTitle("A2", "K2", sheetReport, typeReport, 14);

            string stringMarket = string.Format("Tháng {0}/{1}", month, year);

            // Tạo title
            CreateTitle("A3", "K3", sheetReport, stringMarket, 14);
            
            List<ReportDetailtSTMarket> listDataCompareMonth = new ReportBL().ReportDetailtCompareMonthForAll(year, month, reportTypeID);

            // clone Object
            List<ReportDetailtSTMarket> listDataCompareMonthClone = new List<ReportDetailtSTMarket>(listDataCompareMonth);

            foreach (ReportDetailtSTMarket item in listDataCompareMonth)
            {
                item.TongDS = item.DSChiQuay + item.DSChiNha + item.DSCK;
            }

            // Danh sách mã thị trường của Tất cả
            List<string> listMarket = new List<string>() { "003", "005", "001", "002", "014", "004" };

            // Khởi tạo datatable
            DataTable table = new DataTable();
            // Tạo các cột cho datatable
            // tháng hiện tại
            table.Columns.Add("MarketName", typeof(String));
            table.Columns.Add("CQ1", typeof(double));
            table.Columns.Add("CN1", typeof(double));
            table.Columns.Add("CK1", typeof(double));
            table.Columns.Add("TDS1", typeof(double));

            // tháng trước
            table.Columns.Add("CQ2", typeof(double));
            table.Columns.Add("CN2", typeof(double));
            table.Columns.Add("CK2", typeof(double));
            table.Columns.Add("TDS2", typeof(double));

            // Cùng kì năm ngoái
            table.Columns.Add("CQ3", typeof(double));
            table.Columns.Add("CN3", typeof(double));
            table.Columns.Add("CK3", typeof(double));
            table.Columns.Add("TDS3", typeof(double));

            // Tăng giảm so với tháng trước
            table.Columns.Add("CQ4", typeof(double));
            table.Columns.Add("CN4", typeof(double));
            table.Columns.Add("CK4", typeof(double));
            table.Columns.Add("TDS4", typeof(double));

            // Tăng giảm so với cùng kì năm trước
            table.Columns.Add("CQ5", typeof(double));
            table.Columns.Add("CN5", typeof(double));
            table.Columns.Add("CK5", typeof(double));
            table.Columns.Add("TDS5", typeof(double));

            foreach (string item in listMarket)
            {
                
                ReportDetailtSTMarket dataItemLastYear = listDataCompareMonth.Find(x => x.MarketCode == item && x.Month == month.ToString() && x.Year == (year - 1).ToString());
                ReportDetailtSTMarket dataItemYear = listDataCompareMonth.Find(x => x.MarketCode == item && x.Month == month.ToString() && x.Year == year.ToString());
                ReportDetailtSTMarket dataItemLastMonth = listDataCompareMonth.Find(x => x.MarketCode == item && x.Month == (month - 1).ToString() && x.Year == year.ToString());

                if (dataItemLastYear == null)
                {
                    dataItemLastYear = new ReportDetailtSTMarket();
                }

                if (dataItemYear == null)
                {
                    dataItemYear = new ReportDetailtSTMarket();
                }

                if (dataItemLastMonth == null)
                {
                    dataItemLastMonth = new ReportDetailtSTMarket();
                }

                double sumDSChiQuayYear = dataItemYear.DSChiQuay - dataItemLastMonth.DSChiQuay;
                double sumDSChiNhaYear = dataItemYear.DSChiNha - dataItemLastMonth.DSChiNha;
                double sumDSCKYear = dataItemYear.DSCK - dataItemLastMonth.DSCK;
                double sumTongDSYear = dataItemYear.TongDS - dataItemLastMonth.TongDS;

                double sumDSChiQuayLastYear = dataItemYear.DSChiQuay - dataItemLastYear.DSChiQuay;
                double sumDSChiNhaLastYear = dataItemYear.DSChiNha - dataItemLastYear.DSChiNha;
                double sumDSCKLastYear = dataItemYear.DSCK - dataItemLastYear.DSCK;
                double sumTongDSLastYear = dataItemYear.TongDS - dataItemLastYear.TongDS;

                if (dataItemLastYear != null && dataItemYear != null && dataItemLastMonth != null)
                {
                    // add item vào table
                    table.Rows.Add(dataItemYear.MarketName
                        , dataItemYear.DSChiQuay, dataItemYear.DSChiNha, dataItemYear.DSCK, dataItemYear.TongDS
                        , dataItemLastMonth.DSChiQuay, dataItemLastMonth.DSChiNha, dataItemLastMonth.DSCK, dataItemLastMonth.TongDS
                        , dataItemLastYear.DSChiQuay, dataItemLastYear.DSChiNha, dataItemLastYear.DSCK, dataItemLastYear.TongDS
                        , sumDSChiQuayYear, sumDSChiNhaYear, sumDSCKYear, sumTongDSYear
                        , sumDSChiQuayLastYear, sumDSChiNhaLastYear, sumDSCKLastYear, sumTongDSLastYear);
                }
            }

            DataRow row = table.NewRow();
            row["MarketName"] = "Tổng";
            row["CQ1"] = table.Compute("Sum(CQ1)", "");
            row["CN1"] = table.Compute("Sum(CN1)", "");
            row["CK1"] = table.Compute("Sum(CK1)", "");
            row["TDS1"] = table.Compute("Sum(TDS1)", "");

            row["CQ2"] = table.Compute("Sum(CQ2)", "");
            row["CN2"] = table.Compute("Sum(CN2)", "");
            row["CK2"] = table.Compute("Sum(CK2)", "");
            row["TDS2"] = table.Compute("Sum(TDS2)", "");

            row["CQ3"] = table.Compute("Sum(CQ3)", "");
            row["CN3"] = table.Compute("Sum(CN3)", "");
            row["CK3"] = table.Compute("Sum(CK3)", "");
            row["TDS3"] = table.Compute("Sum(TDS3)", "");

            row["CQ4"] = table.Compute("Sum(CQ4)", "");
            row["CN4"] = table.Compute("Sum(CN4)", "");
            row["CK4"] = table.Compute("Sum(CK4)", "");
            row["TDS4"] = table.Compute("Sum(TDS4)", "");

            row["CQ5"] = table.Compute("Sum(CQ5)", "");
            row["CN5"] = table.Compute("Sum(CN5)", "");
            row["CK5"] = table.Compute("Sum(CK5)", "");
            row["TDS5"] = table.Compute("Sum(TDS5)", "");
            table.Rows.Add(row);

            // Set border
            Style style = new CellsFactory().CreateStyle();
            style.SetBorder(BorderType.BottomBorder, CellBorderType.Thin, Color.Black);
            style.SetBorder(BorderType.LeftBorder, CellBorderType.Thin, Color.Black);
            style.SetBorder(BorderType.RightBorder, CellBorderType.Thin, Color.Black);
            style.SetBorder(BorderType.TopBorder, CellBorderType.Thin, Color.Black);


            // Tạo header cho table
            stringMarket = string.Format("Tháng {0}/{1}", month, year);
            CreateTitle("C61", "F61", sheetReport, stringMarket, 14, true);
            // Tháng trước
            stringMarket = string.Format("Tháng {0}/{1}", month - 1, year);
            CreateTitle("G61", "J61", sheetReport, stringMarket, 14, true);
            // Cùng kì năm trước
            stringMarket = string.Format("Tháng {0}/{1}", month, year - 1);
            CreateTitle("K61", "N61", sheetReport, stringMarket, 14, true);

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
                    int totalCol = 1 + 21;
                    for (int b = 1; b < totalCol; b++)
                    {
                        // Giá trị của value trong table
                        string valueOfTable = table.Rows[stepRow][stepColumn].ToString();

                        // Tô màu cho các dòng có giá trị tăng giảm
                        if (b >= 14)
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

            // vẻ biểu đồ cột
            // Create Chart Line
            //Chart reference
            Aspose.Cells.Charts.Chart leadSourceColumn;
            //Add Pie Chart
            int chartIndex = sheetReport.Charts.Add(ChartType.Column3DClustered, 6, 0, 30, 13);
            leadSourceColumn = sheetReport.Charts[chartIndex];


            //Chart title
            leadSourceColumn.Title.Text = string.Format("Doanh số từng thị trường từng loại dịch vụ \n {0}", string.Format("Tháng {0}/{1}", month, year));
            leadSourceColumn.Title.Font.Color = Color.Black;

            // Set width cho column
            sheetReport.Cells.SetColumnWidthPixel(15, 220);

            // List loại chi trả
            string[] str = {
                string.Format("Chi Quầy {0}/ {1}", month, year), string.Format("Chi Nhà {0}/ {1}", month, year), string.Format("Chuyển khoản {0}/ {1}", month, year)
              , string.Format("Chi Quầy {0}/ {1}", month - 1, year), string.Format("Chi Nhà {0}/ {1}", month - 1, year), string.Format("Chuyển khoản {0}/ {1}", month - 1, year)
              , string.Format("Chi Quầy {0}/ {1}", month, year - 1), string.Format("Chi Nhà {0}/ {1}", month, year - 1), string.Format("Chuyển khoản {0}/ {1}", month, year - 1)
            };

            int count = 0;
            foreach (ReportDetailtSTMarket item in listDataCompareMonthClone)
            {
                // Loại bỏ sự lập lại của Market name
                if (count < 6)
                {
                    // Cùng kì
                    ReportDetailtSTMarket dataItemLastYear = listDataCompareMonth.Find(x => x.MarketCode == item.MarketCode && x.Month == month.ToString() && x.Year == (year - 1).ToString());
                    ReportDetailtSTMarket dataItemYear = listDataCompareMonth.Find(x => x.MarketCode == item.MarketCode && x.Month == month.ToString() && x.Year == year.ToString());
                    ReportDetailtSTMarket dataItemLastMonth = listDataCompareMonth.Find(x => x.MarketCode == item.MarketCode && x.Month == (month - 1).ToString() && x.Year == year.ToString());
                    
                    string totalRowData = string.Concat("{", string.Format("{0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}"
                        , dataItemYear.DSChiQuay, dataItemYear.DSChiNha, dataItemYear.DSCK
                        , dataItemLastMonth.DSChiQuay, dataItemLastMonth.DSChiNha, dataItemLastMonth.DSCK
                        , dataItemLastYear.DSChiQuay, dataItemLastYear.DSChiNha, dataItemLastYear.DSCK)
                        , "}");

                    leadSourceColumn.NSeries.Add(totalRowData, true);

                    string categoryData = string.Concat("{", string.Join(", ", str), "}");
                    leadSourceColumn.NSeries.CategoryData = categoryData;

                    leadSourceColumn.NSeries[count].Name = dataItemYear.MarketName;

                    switch (count)
                    {
                        case 1:
                            // Set the 1st series fill color.
                            leadSourceColumn.NSeries[count].Area.ForegroundColor = Color.Orange;
                            leadSourceColumn.NSeries[count].Area.Formatting = FormattingType.Custom;
                            break;
                        case 2:
                            // Set the 1st series fill color.
                            leadSourceColumn.NSeries[count].Area.ForegroundColor = Color.Green;
                            leadSourceColumn.NSeries[count].Area.Formatting = FormattingType.Custom;
                            break;
                        case 3:
                            // Set the 1st series fill color.
                            leadSourceColumn.NSeries[count].Area.ForegroundColor = Color.Blue;
                            leadSourceColumn.NSeries[count].Area.Formatting = FormattingType.Custom;
                            break;
                        case 4:
                            // Set the 1st series fill color.
                            leadSourceColumn.NSeries[count].Area.ForegroundColor = Color.Red;
                            leadSourceColumn.NSeries[count].Area.Formatting = FormattingType.Custom;
                            break;
                        case 5:
                            // Set the 1st series fill color.
                            leadSourceColumn.NSeries[count].Area.ForegroundColor = Color.Silver;
                            leadSourceColumn.NSeries[count].Area.Formatting = FormattingType.Custom;
                            break;
                        default:
                            // Set the 1st series fill color.
                            leadSourceColumn.NSeries[count].Area.ForegroundColor = Color.Pink;
                            leadSourceColumn.NSeries[count].Area.Formatting = FormattingType.Custom;
                            break;
                    }

                    count++;
                }
            }
            
            // Set plot area formatting as none and hide its border.
            leadSourceColumn.PlotArea.Area.FillFormat.FillType = FillType.None;
            leadSourceColumn.PlotArea.Border.IsVisible = false;

            // vẻ biểu đồ cột Stack
            // Create Chart Line
            //Chart reference
            Aspose.Cells.Charts.Chart leadSourceColumnStack;
            //Add Pie Chart
            int chartIndexStack = sheetReport.Charts.Add(ChartType.Column3D100PercentStacked, 32, 0, 56, 13);
            leadSourceColumnStack = sheetReport.Charts[chartIndexStack];


            //Chart title
            leadSourceColumnStack.Title.Text = string.Format("Doanh số từng thị trường từng loại dịch vụ \n {0}", string.Format("Tháng {0}/{1}", month, year));
            leadSourceColumnStack.Title.Font.Color = Color.Black;

            // Set width cho column
            sheetReport.Cells.SetColumnWidthPixel(15, 220);

            string[] strStack = {
                string.Format("Chi Quầy {0}/ {1}", month, year), string.Format("Chi Nhà {0}/ {1}", month, year), string.Format("Chuyển khoản {0}/ {1}", month, year)
              , string.Format("Chi Quầy {0}/ {1}", month - 1, year), string.Format("Chi Nhà {0}/ {1}", month - 1, year), string.Format("Chuyển khoản {0}/ {1}", month - 1, year)
              , string.Format("Chi Quầy {0}/ {1}", month, year - 1), string.Format("Chi Nhà {0}/ {1}", month, year - 1), string.Format("Chuyển khoản {0}/ {1}", month, year - 1)
            };

            count = 0;

            List<ReportDetailtSTMarket> resultConvert = new List<ReportDetailtSTMarket>();

            if (listDataCompareMonthClone.Count > 0)
            {
                double sumDSChiQuayYear = 0;
                double sumDSChiNhaYear = 0;
                double sumDSCKYear = 0;

                double sumDSChiQuayLastMonth = 0;
                double sumDSChiNhaLastMonth = 0;
                double sumDSCKLastMonth = 0;

                double sumDSChiQuayLastYear = 0;
                double sumDSChiNhaLastYear = 0;
                double sumDSCKLastYear = 0;

                if (listDataCompareMonthClone.Count > 0)
                {
                    // Tháng hiện tại
                    sumDSChiQuayYear = listDataCompareMonthClone.Where(x => x.Year == year.ToString() && x.Month == month.ToString()).Sum(y => y.DSChiQuay);
                    sumDSChiNhaYear = listDataCompareMonthClone.Where(x => x.Year == year.ToString() && x.Month == month.ToString()).Sum(y => y.DSChiNha);
                    sumDSCKYear = listDataCompareMonthClone.Where(x => x.Year == year.ToString() && x.Month == month.ToString()).Sum(y => y.DSCK);

                    // Tháng hiện tại
                    sumDSChiQuayLastMonth = listDataCompareMonthClone.Where(x => x.Year == year.ToString() && x.Month == (month - 1).ToString()).Sum(y => y.DSChiQuay);
                    sumDSChiNhaLastMonth = listDataCompareMonthClone.Where(x => x.Year == year.ToString() && x.Month == (month - 1).ToString()).Sum(y => y.DSChiNha);
                    sumDSCKLastMonth = listDataCompareMonthClone.Where(x => x.Year == year.ToString() && x.Month == (month - 1).ToString()).Sum(y => y.DSCK);

                    // Cùng kì năm trước
                    sumDSChiQuayLastYear = listDataCompareMonthClone.Where(x => x.Year == (year - 1).ToString() && x.Month == month.ToString()).Sum(y => y.DSChiQuay);
                    sumDSChiNhaLastYear = listDataCompareMonthClone.Where(x => x.Year == (year - 1).ToString() && x.Month == month.ToString()).Sum(y => y.DSChiNha);
                    sumDSCKLastYear = listDataCompareMonthClone.Where(x => x.Year == (year - 1).ToString() && x.Month == month.ToString()).Sum(y => y.DSCK);
                }

                foreach (ReportDetailtSTMarket item in listDataCompareMonthClone)
                {
                    item.TongDS = item.DSChiQuay + item.DSChiNha + item.DSCK;
                    ReportDetailtSTMarket itemDetailtPercent = new ReportDetailtSTMarket();
                    
                    // Tháng hiện tại
                    if (item.Year == year.ToString() && item.Month == month.ToString())
                    {
                        itemDetailtPercent = new ReportDetailtSTMarket()
                        {
                            MarketCode = item.MarketCode,
                            MarketName = item.MarketName,
                            DSChiQuay = sumDSChiQuayYear == 0 ? 0 : Math.Round((item.DSChiQuay / sumDSChiQuayYear) * 100, 2, MidpointRounding.ToEven),
                            DSChiNha = sumDSChiNhaYear == 0 ? 0 : Math.Round((item.DSChiNha / sumDSChiNhaYear) * 100, 2, MidpointRounding.ToEven),
                            DSCK = sumDSCKYear == 0 ? 0 : Math.Round((item.DSCK / sumDSCKYear) * 100, 2, MidpointRounding.ToEven),
                            Year = item.Year,
                            Month = item.Month
                        };
                    }

                    // Tháng trước
                    if (item.Year == year.ToString() && item.Month == (month - 1).ToString())
                    {
                        itemDetailtPercent = new ReportDetailtSTMarket()
                        {
                            MarketCode = item.MarketCode,
                            MarketName = item.MarketName,
                            DSChiQuay = sumDSChiQuayLastMonth == 0 ? 0 : Math.Round((item.DSChiQuay / sumDSChiQuayLastMonth) * 100, 2, MidpointRounding.ToEven),
                            DSChiNha = sumDSChiNhaLastMonth == 0 ? 0 : Math.Round((item.DSChiNha / sumDSChiNhaLastMonth) * 100, 2, MidpointRounding.ToEven),
                            DSCK = sumDSCKLastMonth == 0 ? 0 : Math.Round((item.DSCK / sumDSCKLastMonth) * 100, 2, MidpointRounding.ToEven),
                            Year = item.Year,
                            Month = item.Month
                        };
                    }

                    // Cùng kì năm trước
                    if (item.Year == (year - 1).ToString() && item.Month == month.ToString())
                    {
                        itemDetailtPercent = new ReportDetailtSTMarket()
                        {
                            MarketCode = item.MarketCode,
                            MarketName = item.MarketName,
                            DSChiQuay = sumDSChiQuayLastYear == 0 ? 0 : Math.Round((item.DSChiQuay / sumDSChiQuayLastYear) * 100, 2, MidpointRounding.ToEven),
                            DSChiNha = sumDSChiNhaLastYear == 0 ? 0 : Math.Round((item.DSChiNha / sumDSChiNhaLastYear) * 100, 2, MidpointRounding.ToEven),
                            DSCK = sumDSCKLastYear == 0 ? 0 : Math.Round((item.DSCK / sumDSCKLastYear) * 100, 2, MidpointRounding.ToEven),
                            Year = item.Year,
                            Month = item.Month
                        };
                    }

                    resultConvert.Add(itemDetailtPercent);
                }
            }

            foreach (ReportDetailtSTMarket item in resultConvert)
            {

                // Loại bỏ sự lập lại của Market name
                if (count < 6)
                {
                    //item.TongDS = item.DSChiQuay + item.DSChiNha + item.DSCK;

                    // Cùng kì
                    ReportDetailtSTMarket dataItemLastYear = resultConvert.Find(x => x.MarketCode == item.MarketCode && x.Month == month.ToString() && x.Year == (year - 1).ToString());
                    ReportDetailtSTMarket dataItemYear = resultConvert.Find(x => x.MarketCode == item.MarketCode && x.Month == month.ToString() && x.Year == year.ToString());
                    ReportDetailtSTMarket dataItemLastMonth = resultConvert.Find(x => x.MarketCode == item.MarketCode && x.Month == (month - 1).ToString() && x.Year == year.ToString());

                    string totalRowData = string.Concat("{", string.Format("{0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}"
                        , dataItemYear.DSChiQuay, dataItemYear.DSChiNha, dataItemYear.DSCK
                        , dataItemLastMonth.DSChiQuay, dataItemLastMonth.DSChiNha, dataItemLastMonth.DSCK
                        , dataItemLastYear.DSChiQuay, dataItemLastYear.DSChiNha, dataItemLastYear.DSCK)
                        , "}");

                    leadSourceColumnStack.NSeries.Add(totalRowData, true);

                    string categoryData = string.Concat("{", string.Join(", ", strStack), "}");
                    leadSourceColumnStack.NSeries.CategoryData = categoryData;

                    leadSourceColumnStack.NSeries[count].Name = dataItemYear.MarketName;

                    switch (count)
                    {
                        case 1:
                            // Set the 1st series fill color.
                            leadSourceColumnStack.NSeries[count].Area.ForegroundColor = Color.Orange;
                            leadSourceColumnStack.NSeries[count].Area.Formatting = FormattingType.Custom;
                            break;
                        case 2:
                            // Set the 1st series fill color.
                            leadSourceColumnStack.NSeries[count].Area.ForegroundColor = Color.Green;
                            leadSourceColumnStack.NSeries[count].Area.Formatting = FormattingType.Custom;
                            break;
                        case 3:
                            // Set the 1st series fill color.
                            leadSourceColumnStack.NSeries[count].Area.ForegroundColor = Color.Blue;
                            leadSourceColumnStack.NSeries[count].Area.Formatting = FormattingType.Custom;
                            break;
                        case 4:
                            // Set the 1st series fill color.
                            leadSourceColumnStack.NSeries[count].Area.ForegroundColor = Color.Red;
                            leadSourceColumnStack.NSeries[count].Area.Formatting = FormattingType.Custom;
                            break;
                        case 5:
                            // Set the 1st series fill color.
                            leadSourceColumnStack.NSeries[count].Area.ForegroundColor = Color.Silver;
                            leadSourceColumnStack.NSeries[count].Area.Formatting = FormattingType.Custom;
                            break;
                        default:
                            // Set the 1st series fill color.
                            leadSourceColumnStack.NSeries[count].Area.ForegroundColor = Color.Pink;
                            leadSourceColumnStack.NSeries[count].Area.Formatting = FormattingType.Custom;
                            break;
                    }

                    count++;
                }
            }


            // Set plot area formatting as none and hide its border.
            leadSourceColumnStack.PlotArea.Area.FillFormat.FillType = FillType.None;
            leadSourceColumnStack.PlotArea.Border.IsVisible = false;
            // Chạy process
            designer.Process();
            return ExportReport("ReportDetailMarketCompareForAll", designer);
        }

        /// <summary>
        /// Tạo mẫu cho Excel cho so sánh theo tháng cho từng thị trường
        /// </summary>
        /// <param name="gradationID"></param>
        /// <param name="year"></param>
        /// <param name="typeID"></param>
        /// <returns></returns>
        public ActionResult CreateExcelCompareMonthForOne(int month, int year, string reportTypeID, string marketID)
        {
            WorkbookDesigner designer = new WorkbookDesigner();
            string templatePath = "~/Content/Report/ReportDetailMarketCompareForOne.xlsx";
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
            
            string stringMarket = string.Empty;

            switch (marketID)
            {
                case "001":
                    stringMarket = "Thị trường: Châu Âu";
                    break;
                case "002":
                    stringMarket = "Thị trường: Mỹ";
                    break;
                case "003":
                    stringMarket = "Thị trường: Canada";
                    break;
                case "004":
                    stringMarket = "Thị trường: Úc";
                    break;
                case "005":
                    stringMarket = "Thị trường: Châu Á";
                    break;
                default:
                    stringMarket = "Thị trường: Toàn cầu";
                    break;
            }
            // Tạo title
            CreateTitle("A3", "K3", sheetReport, stringMarket, 14);

            // Tạo title detailt
            string titleDetailt = string.Format("Tháng {0}/{1}", month, year);
            CreateTitle("A4", "K4", sheetReport, titleDetailt, 12);

            List<ReportDetailtServiceType> listDataCompareMonth = new ReportBL().ReportDetailtCompareMonthForOne(year, month, reportTypeID, marketID);

            // clone Object
            List<ReportDetailtServiceType> listDataCompareMonthClone = new List<ReportDetailtServiceType>(listDataCompareMonth);

            foreach (ReportDetailtServiceType item in listDataCompareMonth)
            {
                item.TongDS = item.DSChiQuay + item.DSChiNha + item.DSCK;
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

            // tháng trước
            table.Columns.Add("CQ2", typeof(double));
            table.Columns.Add("CN2", typeof(double));
            table.Columns.Add("CK2", typeof(double));
            table.Columns.Add("TDS2", typeof(double));

            // Cùng kì năm ngoái
            table.Columns.Add("CQ3", typeof(double));
            table.Columns.Add("CN3", typeof(double));
            table.Columns.Add("CK3", typeof(double));
            table.Columns.Add("TDS3", typeof(double));

            // So sánh với tháng trước
            table.Columns.Add("CQ4", typeof(double));
            table.Columns.Add("CN4", typeof(double));
            table.Columns.Add("CK4", typeof(double));
            table.Columns.Add("TDS4", typeof(double));

            // So sánh cùng kì năm ngoái
            table.Columns.Add("CQ5", typeof(double));
            table.Columns.Add("CN5", typeof(double));
            table.Columns.Add("CK5", typeof(double));
            table.Columns.Add("TDS5", typeof(double));

            List<string> listTemp = new List<string>();

            foreach (ReportDetailtServiceType item in listDataCompareMonth)
            {
                // Cùng kì
                ReportDetailtServiceType dataItemLastYear = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == month.ToString() && x.Year == (year - 1).ToString());
                ReportDetailtServiceType dataItemYear = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == month.ToString() && x.Year == year.ToString());
                ReportDetailtServiceType dataItemLastMonth = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == (month - 1).ToString() && x.Year == year.ToString());

                if (!listTemp.Contains(item.PartnerCode))
                {
                    // Trường hợp năm trước có đối tác không có
                    if (dataItemLastYear == null)
                    {
                        dataItemLastYear = new ReportDetailtServiceType();
                        dataItemLastYear.PartnerCode = item.PartnerCode;
                        dataItemLastYear.PartnerName = item.PartnerName;
                        dataItemLastYear.DSChiQuay = 0;
                        dataItemLastYear.DSChiNha = 0;
                        dataItemLastYear.DSCK = 0;
                        dataItemLastYear.Year = item.Year;
                        dataItemLastYear.Month = item.Month;
                    }

                    // Trường hợp năm có đối tác không có
                    if (dataItemYear == null)
                    {
                        dataItemYear = new ReportDetailtServiceType();
                        dataItemYear.PartnerCode = item.PartnerCode;
                        dataItemYear.PartnerName = item.PartnerName;
                        dataItemYear.DSChiQuay = 0;
                        dataItemYear.DSChiNha = 0;
                        dataItemYear.DSCK = 0;
                        dataItemYear.Year = item.Year;
                        dataItemYear.Month = item.Month;
                    }

                    // Trường hợp năm có tháng trước có đối tác không có
                    if (dataItemLastMonth == null)
                    {
                        dataItemLastMonth = new ReportDetailtServiceType();
                        dataItemLastMonth.PartnerCode = item.PartnerCode;
                        dataItemLastMonth.PartnerName = item.PartnerName;
                        dataItemLastMonth.DSChiQuay = 0;
                        dataItemLastMonth.DSChiNha = 0;
                        dataItemLastMonth.DSCK = 0;
                        dataItemLastMonth.Year = item.Year;
                        dataItemLastMonth.Month = item.Month;
                    }

                    double sumDSChiQuayYear = dataItemYear.DSChiQuay - dataItemLastMonth.DSChiQuay;
                    double sumDSChiNhaYear = dataItemYear.DSChiNha - dataItemLastMonth.DSChiNha;
                    double sumDSCKYear = dataItemYear.DSCK - dataItemLastMonth.DSCK;
                    double sumTongDSMonthYear = dataItemYear.TongDS - dataItemLastMonth.TongDS;

                    double sumDSChiQuayLastYear = dataItemYear.DSChiQuay - dataItemLastYear.DSChiQuay;
                    double sumDSChiNhaLastYear = dataItemYear.DSChiNha - dataItemLastYear.DSChiNha;
                    double sumDSCKLastYear = dataItemYear.DSCK - dataItemLastYear.DSCK;
                    double sumTongDSLastMonthYear = dataItemYear.TongDS - dataItemLastYear.TongDS;

                    // Check tồn tại của item
                    string value = string.Format("PartnerName='{0}'", item.PartnerName);
                    DataRow[] foundRows = table.Select(value);

                    if (dataItemLastYear != null && dataItemYear != null && dataItemLastMonth != null && foundRows.Count() == 0)
                    {
                        // add item vào table
                        table.Rows.Add(dataItemYear.PartnerName
                            , dataItemYear.DSChiQuay, dataItemYear.DSChiNha, dataItemYear.DSCK, dataItemYear.TongDS
                            , dataItemLastMonth.DSChiQuay, dataItemLastMonth.DSChiNha, dataItemLastMonth.DSCK, dataItemLastMonth.TongDS
                            , dataItemLastYear.DSChiQuay, dataItemLastYear.DSChiNha, dataItemLastYear.DSCK, dataItemLastYear.TongDS
                            , sumDSChiQuayYear, sumDSChiNhaYear, sumDSCKYear, sumTongDSMonthYear
                            , sumDSChiQuayLastYear, sumDSChiNhaLastYear, sumDSCKLastYear, sumTongDSLastMonthYear
                            );
                    }

                    // Add partnerCode để kiểm tra
                    listTemp.Add(item.PartnerCode);
                }
            }

            DataRow row = table.NewRow();
            row["PartnerName"] = "Tổng";
            row["CQ1"] = table.Compute("Sum(CQ1)", "");
            row["CN1"] = table.Compute("Sum(CN1)", "");
            row["CK1"] = table.Compute("Sum(CK1)", "");
            row["TDS1"] = table.Compute("Sum(TDS1)", "");

            row["CQ2"] = table.Compute("Sum(CQ2)", "");
            row["CN2"] = table.Compute("Sum(CN2)", "");
            row["CK2"] = table.Compute("Sum(CK2)", "");
            row["TDS2"] = table.Compute("Sum(TDS2)", "");

            row["CQ3"] = table.Compute("Sum(CQ3)", "");
            row["CN3"] = table.Compute("Sum(CN3)", "");
            row["CK3"] = table.Compute("Sum(CK3)", "");
            row["TDS3"] = table.Compute("Sum(TDS3)", "");

            row["CQ4"] = table.Compute("Sum(CQ4)", "");
            row["CN4"] = table.Compute("Sum(CN4)", "");
            row["CK4"] = table.Compute("Sum(CK4)", "");
            row["TDS4"] = table.Compute("Sum(TDS4)", "");

            row["CQ5"] = table.Compute("Sum(CQ5)", "");
            row["CN5"] = table.Compute("Sum(CN5)", "");
            row["CK5"] = table.Compute("Sum(CK5)", "");
            row["TDS5"] = table.Compute("Sum(TDS5)", "");
            table.Rows.Add(row);

            // Set border
            Style style = new CellsFactory().CreateStyle();
            style.SetBorder(BorderType.BottomBorder, CellBorderType.Thin, Color.Black);
            style.SetBorder(BorderType.LeftBorder, CellBorderType.Thin, Color.Black);
            style.SetBorder(BorderType.RightBorder, CellBorderType.Thin, Color.Black);
            style.SetBorder(BorderType.TopBorder, CellBorderType.Thin, Color.Black);


            // Tạo header cho table
            stringMarket = string.Format("Tháng {0}/{1}", month, year);
            CreateTitle("C61", "F61", sheetReport, stringMarket, 14, true);
            // Tháng trước
            stringMarket = string.Format("Tháng {0}/{1}", month - 1, year);
            CreateTitle("G61", "J61", sheetReport, stringMarket, 14, true);
            // Cùng kì năm trước
            stringMarket = string.Format("Tháng {0}/{1}", month, year - 1);
            CreateTitle("K61", "N61", sheetReport, stringMarket, 14, true);

            // Get count table
            int countTable1 = table.Rows.Count;

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
                    int totalCol = 1 + 21;
                    for (int b = 1; b < totalCol; b++)
                    {
                        // Giá trị của value trong table
                        string valueOfTable = table.Rows[stepRow][stepColumn].ToString();

                        // Tô màu cho các dòng có giá trị tăng giảm
                        if (b >= 14)
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

            // vẻ biểu đồ cột
            // Create Chart Line
            //Chart reference
            Aspose.Cells.Charts.Chart leadSourceColumn;
            //Add Pie Chart
            int chartIndex = sheetReport.Charts.Add(ChartType.Column3DClustered, 6, 0, 30, 13);
            leadSourceColumn = sheetReport.Charts[chartIndex];


            //Chart title
            leadSourceColumn.Title.Text = string.Format("Doanh số từng thị trường từng loại dịch vụ \n {0}", string.Format("Tháng {0}/{1}", month, year));
            leadSourceColumn.Title.Font.Color = Color.Black;

            // Set width cho column
            sheetReport.Cells.SetColumnWidthPixel(15, 220);

            // List loại chi trả
            string[] str = {
                string.Format("Chi Quầy {0}/ {1}", month, year), string.Format("Chi Nhà {0}/ {1}", month, year), string.Format("Chuyển khoản {0}/ {1}", month, year)
              , string.Format("Chi Quầy {0}/ {1}", month - 1, year), string.Format("Chi Nhà {0}/ {1}", month - 1, year), string.Format("Chuyển khoản {0}/ {1}", month - 1, year)
              , string.Format("Chi Quầy {0}/ {1}", month, year - 1), string.Format("Chi Nhà {0}/ {1}", month, year - 1), string.Format("Chuyển khoản {0}/ {1}", month, year - 1)
            };

            List<string> Listdulicate = new List<string>();
            int count = 0;
            foreach (ReportDetailtServiceType item in listDataCompareMonthClone)
            {
                // pass qua nếu đã tồn tại trong list
                if (Listdulicate.Contains(item.PartnerCode))
                {
                    continue;
                }

                // Add item vào Listdulicate
                Listdulicate.Add(item.PartnerCode);

                // Cùng kì
                ReportDetailtServiceType dataItemLastYear = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == month.ToString() && x.Year == (year - 1).ToString());
                ReportDetailtServiceType dataItemYear = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == month.ToString() && x.Year == year.ToString());
                ReportDetailtServiceType dataItemLastMonth = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == (month - 1).ToString() && x.Year == year.ToString());

                if (dataItemLastYear == null)
                {
                    dataItemLastYear = new ReportDetailtServiceType();
                    dataItemLastYear.PartnerCode = item.PartnerCode;
                    dataItemLastYear.PartnerName = item.PartnerName;
                }

                if (dataItemYear == null)
                {
                    dataItemYear = new ReportDetailtServiceType();
                    dataItemYear.PartnerCode = item.PartnerCode;
                    dataItemYear.PartnerName = item.PartnerName;
                }

                if (dataItemLastMonth == null)
                {
                    dataItemLastMonth = new ReportDetailtServiceType();
                    dataItemLastMonth.PartnerCode = item.PartnerCode;
                    dataItemLastMonth.PartnerName = item.PartnerName;
                }

                string totalRowData = string.Concat("{", string.Format("{0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}"
                    , dataItemYear.DSChiQuay, dataItemYear.DSChiNha, dataItemYear.DSCK
                    , dataItemLastMonth.DSChiQuay, dataItemLastMonth.DSChiNha, dataItemLastMonth.DSCK
                    , dataItemLastYear.DSChiQuay, dataItemLastYear.DSChiNha, dataItemLastYear.DSCK)
                    , "}");

                leadSourceColumn.NSeries.Add(totalRowData, true);

                string categoryData = string.Concat("{", string.Join(", ", str), "}");
                leadSourceColumn.NSeries.CategoryData = categoryData;

                leadSourceColumn.NSeries[count].Name = dataItemYear.PartnerName;

                switch (count)
                {
                    case 1:
                        // Set the 1st series fill color.
                        leadSourceColumn.NSeries[count].Area.ForegroundColor = Color.Orange;
                        leadSourceColumn.NSeries[count].Area.Formatting = FormattingType.Custom;
                        break;
                    case 2:
                        // Set the 1st series fill color.
                        leadSourceColumn.NSeries[count].Area.ForegroundColor = Color.Green;
                        leadSourceColumn.NSeries[count].Area.Formatting = FormattingType.Custom;
                        break;
                    case 3:
                        // Set the 1st series fill color.
                        leadSourceColumn.NSeries[count].Area.ForegroundColor = Color.Blue;
                        leadSourceColumn.NSeries[count].Area.Formatting = FormattingType.Custom;
                        break;
                    case 4:
                        // Set the 1st series fill color.
                        leadSourceColumn.NSeries[count].Area.ForegroundColor = Color.Red;
                        leadSourceColumn.NSeries[count].Area.Formatting = FormattingType.Custom;
                        break;
                    case 5:
                        // Set the 1st series fill color.
                        leadSourceColumn.NSeries[count].Area.ForegroundColor = Color.Silver;
                        leadSourceColumn.NSeries[count].Area.Formatting = FormattingType.Custom;
                        break;
                    default:
                        // Set the 1st series fill color.
                        leadSourceColumn.NSeries[count].Area.ForegroundColor = Color.Pink;
                        leadSourceColumn.NSeries[count].Area.Formatting = FormattingType.Custom;
                        break;
                }

                count++;
            }

            // Set plot area formatting as none and hide its border.
            leadSourceColumn.PlotArea.Area.FillFormat.FillType = FillType.None;
            leadSourceColumn.PlotArea.Border.IsVisible = false;

            // vẻ biểu đồ cột Stack
            // Create Chart Line
            //Chart reference
            Aspose.Cells.Charts.Chart leadSourceColumnStack;
            //Add Pie Chart
            int chartIndexStack = sheetReport.Charts.Add(ChartType.Column3DClustered, 32, 0, 56, 13);
            leadSourceColumnStack = sheetReport.Charts[chartIndexStack];


            //Chart title
            leadSourceColumnStack.Title.Text = string.Format("Doanh số từng thị trường từng loại dịch vụ \n {0}", string.Format("Tháng {0}/{1}", month, year));
            leadSourceColumnStack.Title.Font.Color = Color.Black;

            // Set width cho column
            sheetReport.Cells.SetColumnWidthPixel(15, 220);

            // List loại chi trả
            List<string> strStack = new List<string>();

            count = 0;

            // Get list mới lưu trữ
            Listdulicate = new List<string>();
            List<string> listTongDSYear = new List<string>();
            List<string> listTongDSLastMonth = new List<string>();
            List<string> listTongDSLastYear = new List<string>();

            foreach (ReportDetailtServiceType item in listDataCompareMonthClone)
            {
                // pass qua nếu đã tồn tại trong list
                if (Listdulicate.Contains(item.PartnerCode))
                {
                    continue;
                }

                // Cùng kì
                ReportDetailtServiceType dataItemLastYear = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == month.ToString() && x.Year == (year - 1).ToString());
                ReportDetailtServiceType dataItemYear = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == month.ToString() && x.Year == year.ToString());
                ReportDetailtServiceType dataItemLastMonth = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == (month - 1).ToString() && x.Year == year.ToString());

                if (dataItemLastYear == null)
                {
                    dataItemLastYear = new ReportDetailtServiceType();
                    dataItemLastYear.PartnerCode = item.PartnerCode;
                    dataItemLastYear.PartnerName = item.PartnerName;
                }

                if (dataItemYear == null)
                {
                    dataItemYear = new ReportDetailtServiceType();
                    dataItemYear.PartnerCode = item.PartnerCode;
                    dataItemYear.PartnerName = item.PartnerName;
                }

                if (dataItemLastMonth == null)
                {
                    dataItemLastMonth = new ReportDetailtServiceType();
                    dataItemLastMonth.PartnerCode = item.PartnerCode;
                    dataItemLastMonth.PartnerName = item.PartnerName;
                }

                listTongDSYear.Add(dataItemYear.TongDS.ToString());
                listTongDSLastMonth.Add(dataItemLastMonth.TongDS.ToString());
                listTongDSLastYear.Add(dataItemLastYear.TongDS.ToString());

                // Add item vào Listdulicate
                Listdulicate.Add(item.PartnerCode);
                item.TongDS = item.DSChiQuay + item.DSChiNha + item.DSCK;

                strStack.Add(item.PartnerName);
            }


            // Reset list
            Listdulicate = new List<string>();
            // chạy 3 lần cho 3 tháng
            for (int i = 0; i < 3; i++)
            {
                if (i == 0)
                {
                    string totalRowData = string.Concat("{", string.Join(", ", listTongDSYear), "}");

                    leadSourceColumnStack.NSeries.Add(totalRowData, true);
                }
                else if(i == 1)
                {
                    string totalRowData = string.Concat("{", string.Join(", ", listTongDSLastMonth), "}");

                    leadSourceColumnStack.NSeries.Add(totalRowData, true);
                }
                else
                {
                    string totalRowData = string.Concat("{", string.Join(", ", listTongDSLastYear), "}");

                    leadSourceColumnStack.NSeries.Add(totalRowData, true);
                }
                
                string categoryDataStack = string.Concat("{", string.Join(", ", strStack), "}");
                leadSourceColumnStack.NSeries.CategoryData = categoryDataStack;
            }
            
            leadSourceColumnStack.NSeries[0].Name = string.Format("Tổng tháng {0}/{1}", month, year);
            leadSourceColumnStack.NSeries[1].Name = string.Format("Tổng tháng {0}/{1}", month - 1, year);
            leadSourceColumnStack.NSeries[2].Name = string.Format("Tổng tháng {0}/{1}", month, year - 1);
            
            // Set plot area formatting as none and hide its border.
            leadSourceColumnStack.PlotArea.Area.FillFormat.FillType = FillType.None;
            leadSourceColumnStack.PlotArea.Border.IsVisible = false;


            typeReport = "2. Theo tỷ trọng doanh số theo thị trường - loại hình chi trả";
            CreateTitle("B73", "F73", sheetReport, typeReport, 14);

            List<ReportDetailtServiceType> listDataCompareMonthConvert = new List<ReportDetailtServiceType>(listDataCompareMonth);
            // clone Object
            listDataCompareMonthClone = new List<ReportDetailtServiceType>(listDataCompareMonthConvert);

            foreach (ReportDetailtServiceType item in listDataCompareMonthClone)
            {
                // Cùng kì
                ReportDetailtServiceType dataItemLastYear = listDataCompareMonthConvert.Find(x => x.PartnerCode == item.PartnerCode && x.Month == month.ToString() && x.Year == (year - 1).ToString());
                ReportDetailtServiceType dataItemYear = listDataCompareMonthConvert.Find(x => x.PartnerCode == item.PartnerCode && x.Month == month.ToString() && x.Year == year.ToString());
                ReportDetailtServiceType dataItemLastMonth = listDataCompareMonthConvert.Find(x => x.PartnerCode == item.PartnerCode && x.Month == (month - 1).ToString() && x.Year == year.ToString());
                
                // tháng hiện tại
                if (dataItemYear == null)
                {
                    dataItemYear = new ReportDetailtServiceType();
                    dataItemYear.PartnerCode = item.PartnerCode;
                    dataItemYear.PartnerName = item.PartnerName;
                    dataItemYear.DSChiQuay = 0;
                    dataItemYear.DSChiNha = 0;
                    dataItemYear.DSCK = 0;
                    dataItemYear.DSCK = 0;
                    dataItemYear.TongDS = 0;
                    dataItemYear.Year = year.ToString();
                    dataItemYear.Month = month.ToString();
                    listDataCompareMonthConvert.Add(dataItemYear);
                }

                // Tháng trước
                if (dataItemLastMonth == null)
                {
                    dataItemLastMonth = new ReportDetailtServiceType();
                    dataItemLastMonth.PartnerCode = item.PartnerCode;
                    dataItemLastMonth.PartnerName = item.PartnerName;
                    dataItemLastMonth.DSChiQuay = 0;
                    dataItemLastMonth.DSChiNha = 0;
                    dataItemLastMonth.DSCK = 0;
                    dataItemLastMonth.DSCK = 0;
                    dataItemLastMonth.TongDS = 0;
                    dataItemLastMonth.Year = year.ToString();
                    dataItemLastMonth.Month = month.ToString();
                    listDataCompareMonthConvert.Add(dataItemLastMonth);
                }

                // Cùng kì năm trước
                if (dataItemLastYear == null)
                {
                    dataItemLastYear = new ReportDetailtServiceType();
                    dataItemLastYear.PartnerCode = item.PartnerCode;
                    dataItemLastYear.PartnerName = item.PartnerName;
                    dataItemLastYear.DSChiQuay = 0;
                    dataItemLastYear.DSChiNha = 0;
                    dataItemLastYear.DSCK = 0;
                    dataItemLastYear.DSCK = 0;
                    dataItemLastYear.TongDS = 0;
                    dataItemLastYear.Year = year.ToString();
                    dataItemLastYear.Month = month.ToString();
                    listDataCompareMonthConvert.Add(dataItemLastYear);
                }
            }

            foreach (ReportDetailtServiceType item in listDataCompareMonthConvert)
            {
                item.TongDS = item.DSChiQuay + item.DSChiNha + item.DSCK;
            }

            double sumTongDSYear = listDataCompareMonthConvert.Where(x => x.Year == year.ToString() && x.Month == month.ToString()).Sum(x => x.TongDS);
            double sumTongDSLastMonth = listDataCompareMonthConvert.Where(x => x.Year == year.ToString() && x.Month == (month - 1).ToString()).Sum(x => x.TongDS);
            double sumTongDSLastYear = listDataCompareMonthConvert.Where(x => x.Year == (year - 1).ToString() && x.Month == month.ToString()).Sum(x => x.TongDS);

            List<ReportDetailtServiceType> listDataConvert = new List<ReportDetailtServiceType>();

            // Khởi tạo datatable
            table = new DataTable();
            // Tạo các cột cho datatable
            table.Columns.Add("STT", typeof(String));
            table.Columns.Add("PartnerName", typeof(String));
            table.Columns.Add("LK1", typeof(double));
            table.Columns.Add("LK2", typeof(double));
            table.Columns.Add("LK3", typeof(double));

            count = 1;
            foreach (ReportDetailtServiceType item in listDataCompareMonthConvert)
            {
                // Cùng kì
                ReportDetailtServiceType dataItemLastYear = listDataCompareMonthConvert.Find(x => x.PartnerCode == item.PartnerCode && x.Month == month.ToString() && x.Year == (year - 1).ToString());
                ReportDetailtServiceType dataItemYear = listDataCompareMonthConvert.Find(x => x.PartnerCode == item.PartnerCode && x.Month == month.ToString() && x.Year == year.ToString());
                ReportDetailtServiceType dataItemLastMonth = listDataCompareMonthConvert.Find(x => x.PartnerCode == item.PartnerCode && x.Month == (month - 1).ToString() && x.Year == year.ToString());

                if (dataItemYear == null)
                {
                    dataItemYear = new ReportDetailtServiceType();
                    dataItemYear.PartnerCode = item.PartnerCode;
                    dataItemYear.PartnerName = item.PartnerName;
                }

                if (dataItemLastMonth == null)
                {
                    dataItemLastMonth = new ReportDetailtServiceType();
                    dataItemLastMonth.PartnerCode = item.PartnerCode;
                    dataItemLastMonth.PartnerName = item.PartnerName;
                }

                if (dataItemLastYear == null)
                {
                    dataItemLastYear = new ReportDetailtServiceType();
                    dataItemLastYear.PartnerCode = item.PartnerCode;
                    dataItemLastYear.PartnerName = item.PartnerName;
                }

                // Check tồn tại của item
                string value = string.Format("PartnerName='{0}'", item.PartnerName);
                DataRow[] foundRows = table.Select(value);

                if (dataItemYear != null && dataItemLastMonth != null && dataItemLastYear != null && foundRows.Count() == 0)
                {
                    double valueYear = Math.Round((dataItemYear.TongDS / sumTongDSYear) * 100, 2, MidpointRounding.ToEven);
                    double valueLastMonth = Math.Round((dataItemLastMonth.TongDS / sumTongDSLastMonth) * 100, 2, MidpointRounding.ToEven);
                    double valueLastYear = Math.Round((dataItemLastYear.TongDS / sumTongDSLastYear) * 100, 2, MidpointRounding.ToEven);
                    table.Rows.Add(count, item.PartnerName, valueYear, valueLastMonth, valueLastYear);
                }

                count++;
            }

            row = table.NewRow();
            row["STT"] = "";
            row["PartnerName"] = "Tổng";
            row["LK1"] = 100;
            row["LK2"] = 100;
            row["LK3"] = 100;
            table.Rows.Add(row);

            // Tạo bảng dữ liệu
            int countRowPercent = (countTable1 + 62) + 25;

            if (table.Rows.Count > 0)
            {
                int stepRow = 0;
                // total row = row start + số row hiện có
                int totalRow = table.Rows.Count + countRowPercent;

                // Add header cho bảng dữ liệu phần trăm
                sheetReport.Cells[countRowPercent - 1, 0].PutValue("STT", true, true);
                style.Font.IsBold = true;
                sheetReport.Cells[countRowPercent - 1, 0].SetStyle(style);

                sheetReport.Cells[countRowPercent - 1, 1].PutValue("Đối tác", true, true);
                style.Font.IsBold = true;
                sheetReport.Cells[countRowPercent - 1, 1].SetStyle(style);

                sheetReport.Cells[countRowPercent - 1, 2].PutValue(string.Format("Tháng {0}/{1}", month, year), true, true);
                style.Font.IsBold = true;
                sheetReport.Cells[countRowPercent - 1, 2].SetStyle(style);

                sheetReport.Cells[countRowPercent - 1, 3].PutValue(string.Format("Tháng {0}/{1}", month - 1, year), true, true);
                style.Font.IsBold = true;
                sheetReport.Cells[countRowPercent - 1, 3].SetStyle(style);

                sheetReport.Cells[countRowPercent - 1, 4].PutValue(string.Format("Tháng {0}/{1}", month, year - 1), true, true);
                style.Font.IsBold = true;
                sheetReport.Cells[countRowPercent - 1, 4].SetStyle(style);

                // Số dòng của row
                for (int a = countRowPercent; a < totalRow; a++)
                {
                    int stepColumn = 0;
                    // Số cột trong báo cáo cần hiển thị
                    // Tổng số cột hiển thị = Số cột hiển thị bắt đầu + tổng số cột cần hiển thị
                    int totalCol = 0 + 5;
                    for (int b = 0; b < totalCol; b++)
                    {
                        // Giá trị của value trong table
                        string valueOfTable = table.Rows[stepRow][stepColumn].ToString();

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

            List<string> dataPieYear = new List<string>();
            List<string> dataPieLastMonth = new List<string>();
            List<string> dataPieLastYear = new List<string>();
            List<string> dataPartner = new List<string>();

            count = 0;
            foreach (DataRow dtRow in table.Rows)
            {
                if (count < table.Rows.Count - 1)
                {
                    // On all tables' columns
                    foreach (DataColumn dc in table.Columns)
                    {
                        if (dc.ColumnName.Contains("LK1"))
                        {
                            dataPieYear.Add(dtRow[dc].ToString());
                        }

                        if (dc.ColumnName.Contains("LK2"))
                        {
                            dataPieLastMonth.Add(dtRow[dc].ToString());
                        }

                        if (dc.ColumnName.Contains("LK3"))
                        {
                            dataPieLastYear.Add(dtRow[dc].ToString());
                        }

                        if (dc.ColumnName.Contains("PartnerName"))
                        {
                            dataPartner.Add(dtRow[dc].ToString());
                        }
                    }
                }

                count++;
            }

            countRowPercent = (countTable1 + 62) + 4;

            if (dataPieYear != null)
            {
                //Add Pie Chart
                Aspose.Cells.Charts.Chart leadSourcePie;
                chartIndex = sheetReport.Charts.Add(ChartType.Pie3D, countRowPercent, 0, countRowPercent + 17, 5);
                leadSourcePie = sheetReport.Charts[chartIndex];

                // Set some properties of chart plot area.
                // To set the fill color and make the border invisible.
                leadSourcePie.PlotArea.Border.IsVisible = false;
                leadSourcePie.Elevation = 45;
                // Set properties of chart title
                leadSourcePie.Title.Text = string.Format("Tháng {0}/{1}", month, year);
                leadSourcePie.Title.Font.Color = Color.Blue;
                leadSourcePie.Title.Font.IsBold = true;
                leadSourcePie.Title.Font.Size = 12;

                // Set properties of nseries
                string totalRowData = string.Concat("{", string.Join(", ", dataPieYear), "}");
                leadSourcePie.NSeries.Add(totalRowData, true);

                string categoryData = string.Concat("{", string.Join(", ", dataPartner), "}");
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

            if (dataPieLastMonth != null)
            {
                //Add Pie Chart
                Aspose.Cells.Charts.Chart leadSourcePie;
                chartIndex = sheetReport.Charts.Add(ChartType.Pie3D, countRowPercent, 6, countRowPercent + 17, 11);
                leadSourcePie = sheetReport.Charts[chartIndex];

                // Set some properties of chart plot area.
                // To set the fill color and make the border invisible.
                leadSourcePie.PlotArea.Border.IsVisible = false;
                leadSourcePie.Elevation = 45;
                // Set properties of chart title
                leadSourcePie.Title.Text = string.Format("Tháng {0}/{1}", month - 1, year);
                leadSourcePie.Title.Font.Color = Color.Blue;
                leadSourcePie.Title.Font.IsBold = true;
                leadSourcePie.Title.Font.Size = 12;

                // Set properties of nseries
                string totalRowData = string.Concat("{", string.Join(", ", dataPieLastMonth), "}");
                leadSourcePie.NSeries.Add(totalRowData, true);

                string categoryData = string.Concat("{", string.Join(", ", dataPartner), "}");
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
                Aspose.Cells.Charts.Chart leadSourcePie;
                chartIndex = sheetReport.Charts.Add(ChartType.Pie3D, countRowPercent, 12, countRowPercent + 17, 17);
                leadSourcePie = sheetReport.Charts[chartIndex];

                // Set some properties of chart plot area.
                // To set the fill color and make the border invisible.
                leadSourcePie.PlotArea.Border.IsVisible = false;
                leadSourcePie.Elevation = 45;
                // Set properties of chart title
                leadSourcePie.Title.Text = string.Format("Tháng {0}/{1}", month, year - 1);
                leadSourcePie.Title.Font.Color = Color.Blue;
                leadSourcePie.Title.Font.IsBold = true;
                leadSourcePie.Title.Font.Size = 12;

                // Set properties of nseries
                string totalRowData = string.Concat("{", string.Join(", ", dataPieLastYear), "}");
                leadSourcePie.NSeries.Add(totalRowData, true);

                string categoryData = string.Concat("{", string.Join(", ", dataPartner), "}");
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


            // Chạy process
            designer.Process();
            return ExportReport("ReportDetailMarketCompareForOne", designer);
        }

        /// <summary>
        /// Hàm tạo tile cho xuất Excell
        /// </summary>
        /// <param name="upperLeft"></param>
        /// <param name="upperRight"></param>
        /// <param name="sheetReport"></param>
        /// <param name="Title"></param>
        private void CreateTitle(string upperLeft, string upperRight, Worksheet sheetReport, string Title, int size, bool? setBolder = null)
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

            // Trường hợp muốn set border cột
            if (setBolder == true)
            {
                styleTitle.SetBorder(BorderType.BottomBorder, CellBorderType.Thin, Color.Black);
                styleTitle.SetBorder(BorderType.LeftBorder, CellBorderType.Thin, Color.Black);
                styleTitle.SetBorder(BorderType.RightBorder, CellBorderType.Thin, Color.Black);
                styleTitle.SetBorder(BorderType.TopBorder, CellBorderType.Thin, Color.Black);
            }
            //styleTitle.Font.Name = "Times New Roman";
            styleTitle.Font.Name = "Calibri";
            styleTitle.HorizontalAlignment = TextAlignmentType.Center;
            rangeTitle.SetStyle(styleTitle);
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
            db.Columns.Add("STT", typeof(string));
            db.Columns.Add("MarketName", typeof(string));

            if (typeID.Equals("2"))
            {
                db.Columns.Add("PartnerName", typeof(string));
            }

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
                row["STT"] = (string)dr["STT"];
                row["MarketName"] = (string)dr["MarketName"];

                if (typeID.Equals("2"))
                {
                    row["PartnerName"] = (string)dr["PartnerName"];
                }

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
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
    public class ReportExcelTCKTForMarketController : Controller
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

        // GET: Admin/ReportExcelTCKTForMarket
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
        public ActionResult CreateExcelForDayMonthYear(DateTime fromDate, DateTime toDate, int typeID, string reportTypeID, string marketID)
        {
            WorkbookDesigner designer = new WorkbookDesigner();
            string templatePath = "~/Content/Report/ReportTCKT/ReportTCKTForMakets.xlsx";
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

            CreateTitle("A2", "U2", sheetReport, typeReport, 14);
            
            // Get data report ngày
            List<ReportDetailtSTMarket> listData = new List<ReportDetailtSTMarket>();

            switch (typeID)
            {
                // Theo ngày
                case 1:
                    listData = new ReportBL().SearchMarketTCKTForDay(fromDate, toDate, reportTypeID);
                    // Set from day and to day
                    sheetReport.Cells["L4"].PutValue(fromDate.ToString("dd/MM/yyyy"));
                    sheetReport.Cells["P4"].PutValue(toDate.ToString("dd/MM/yyyy"));
                    break;
                // Theo tháng
                case 2:
                    listData = new ReportBL().SearchMarketTCKTForMonth(fromDate, toDate, reportTypeID);
                    // Set from day and to day
                    sheetReport.Cells["L4"].PutValue(string.Format("{0}/{1}", fromDate.Month, fromDate.Year));
                    sheetReport.Cells["P4"].PutValue(string.Format("{0}/{1}", toDate.Month, toDate.Year));
                    break;
                // Theo năm
                default:
                    //listReportData = new ReportBL().SearchReportMaketForYear(fromDate, toDate, reportTypeID);
                    //// Set from day and to day
                    //sheetReport.Cells["L4"].PutValue(fromDate.Year.ToString());
                    //sheetReport.Cells["P4"].PutValue(toDate.Year.ToString());
                    break;
            }

            // List danh sách các
            List<string> listMarket = new List<string>();

            // Khởi tạo datatable
            DataTable table = new DataTable();

            if (!string.IsNullOrEmpty(marketID))
            {
                if (marketID == "0")
                {
                    foreach (ReportDetailtSTMarket item in listData)
                    {
                        // CHâu Á
                        if (item.ParentCode == "005")
                        {
                            item.MarketName = "Châu Á";
                        }
                        if (!listMarket.Contains(item.MarketName))
                        {
                            listMarket.Add(item.MarketName);
                        }
                        item.TongDS = item.DSChiNha + item.DSChiQuay + item.DSCK;
                    }
                    
                    table.Columns.Add("PartnerName", typeof(String));

                    // Add column
                    for (DateTime x = fromDate; x <= toDate; x = x.AddDays(1))
                    {
                        string nameDate = string.Format("COL{0}", x.Day.ToString("00"));

                        table.Columns.Add(nameDate, typeof(double));

                    }

                    table.Columns.Add("TongDS", typeof(double));

                    table.Columns.Add("MarketName", typeof(String));

                    // List thị trường
                    foreach (string item in listMarket)
                    {
                        // List Partner thuộc 1 thị trường
                        List<ReportDetailtSTMarket> listDataItem = listData.Where(x => x.MarketName == item).ToList();

                        List<string> listPartner = new List<string>();

                        foreach (ReportDetailtSTMarket item1 in listDataItem)
                        {
                            if (listPartner.Contains(item1.PartnerName))
                            {
                                continue;
                            }
                            DataRow rows = table.NewRow();

                            double tongDS = 0;
                            // Từ ngày đến ngày
                            for (DateTime x = fromDate; x <= toDate; x = x.AddDays(1))
                            {
                                string nameDate = string.Format("COL{0}", x.Day.ToString("00"));

                                ReportDetailtSTMarket dataItem = listDataItem.Find(o => o.CreatedDate == x && o.PartnerName == item1.PartnerName);
                                // Trường hợp không có dữ liệu
                                if (dataItem == null)
                                {
                                    dataItem = new ReportDetailtSTMarket()
                                    {
                                        PartnerName = item1.PartnerName,
                                        MarketName = item1.MarketName
                                    };
                                }
                                tongDS = tongDS + dataItem.TongDS;
                                rows[nameDate] = dataItem.TongDS;
                            }

                            // add dòng tổng
                            rows["TongDS"] = tongDS;

                            rows["MarketName"] = item1.MarketName;
                            rows["PartnerName"] = item1.PartnerName;

                            // Add row vào table
                            table.Rows.Add(rows);
                            // Add giá trị của partner
                            listPartner.Add(item1.PartnerName);
                        }
                    }
                }
                else
                {
                    List<ReportDetailtSTMarket> listDataConvert = listData.Where(x => x.ParentCode == "005").ToList();

                    foreach (ReportDetailtSTMarket item in listDataConvert)
                    {
                        if (!listMarket.Contains(item.MarketName))
                        {
                            listMarket.Add(item.MarketName);
                        }
                        item.TongDS = item.DSChiNha + item.DSChiQuay + item.DSCK;
                    }

                    table.Columns.Add("PartnerName", typeof(String));

                    // Add column
                    for (DateTime x = fromDate; x <= toDate; x = x.AddDays(1))
                    {
                        string nameDate = string.Format("COL{0}", x.Day.ToString("00"));

                        table.Columns.Add(nameDate, typeof(double));

                    }

                    table.Columns.Add("TongDS", typeof(double));

                    table.Columns.Add("MarketName", typeof(String));

                    // List thị trường
                    foreach (string item in listMarket)
                    {
                        // List Partner thuộc 1 thị trường
                        List<ReportDetailtSTMarket> listDataItem = listDataConvert.Where(x => x.MarketName == item).ToList();

                        List<string> listPartner = new List<string>();

                        foreach (ReportDetailtSTMarket item1 in listDataItem)
                        {
                            if (listPartner.Contains(item1.PartnerName))
                            {
                                continue;
                            }
                            DataRow rows = table.NewRow();

                            double tongDS = 0;
                            // Từ ngày đến ngày
                            for (DateTime x = fromDate; x <= toDate; x = x.AddDays(1))
                            {
                                string nameDate = string.Format("COL{0}", x.Day.ToString("00"));

                                ReportDetailtSTMarket dataItem = listDataItem.Find(o => o.CreatedDate == x && o.PartnerName == item1.PartnerName);
                                // Trường hợp không có dữ liệu
                                if (dataItem == null)
                                {
                                    dataItem = new ReportDetailtSTMarket()
                                    {
                                        PartnerName = item1.PartnerName,
                                        MarketName = item1.MarketName
                                    };
                                }
                                tongDS = tongDS + dataItem.TongDS;
                                rows[nameDate] = dataItem.TongDS;
                            }

                            // add dòng tổng
                            rows["TongDS"] = tongDS;

                            rows["MarketName"] = item1.MarketName;
                            rows["PartnerName"] = item1.PartnerName;
                            // Add row vào table
                            table.Rows.Add(rows);
                            // Add giá trị của partner
                            listPartner.Add(item1.PartnerName);
                        }
                    }
                }
            }

            // Set border
            Style style = new CellsFactory().CreateStyle();
            style.SetBorder(BorderType.BottomBorder, CellBorderType.Thin, Color.Black);
            style.SetBorder(BorderType.LeftBorder, CellBorderType.Thin, Color.Black);
            style.SetBorder(BorderType.RightBorder, CellBorderType.Thin, Color.Black);
            style.SetBorder(BorderType.TopBorder, CellBorderType.Thin, Color.Black);

            // Tạo cột header cho excel
            string[] listCol = { "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF", "AG" };
            // Add column
            int i = 0;
            for (DateTime x = fromDate; x <= toDate; x = x.AddDays(1))
            {
                string nameDate = string.Format("COL{0}", x.Day.ToString("00"));

                string nameCol = string.Format("Ngày {0}", x.Day.ToString("00"));

                sheetReport.Cells[string.Format("{0}7", listCol[i])].PutValue(nameCol);
                sheetReport.Cells[string.Format("{0}7", listCol[i])].SetStyle(style);

                i++;
            }
            // Dòng tổng
            sheetReport.Cells[string.Format("{0}7", listCol[i])].PutValue("Tổng");
            sheetReport.Cells[string.Format("{0}7", listCol[i])].SetStyle(style);

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
                    int totalCol = 1 + table.Columns.Count;
                    for (int b = 1; b < totalCol; b++)
                    {
                        // Giá trị của value trong table
                        string valueOfTable = table.Rows[stepRow][stepColumn].ToString();

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
            db.Columns.Add("American", typeof(double));
            db.Columns.Add("Asia", typeof(double));
            db.Columns.Add("Global", typeof(double));
            db.Columns.Add("Europe", typeof(double));
            db.Columns.Add("Canada", typeof(double));
            db.Columns.Add("Australia", typeof(double));
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
                row["American"] = (double)dr["American"];
                row["Asia"] = (double)dr["Asia"];
                row["Global"] = (double)dr["Global"];
                row["Europe"] = (double)dr["Europe"];
                row["Canada"] = (double)dr["Canada"];
                row["Australia"] = (double)dr["Australia"];
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
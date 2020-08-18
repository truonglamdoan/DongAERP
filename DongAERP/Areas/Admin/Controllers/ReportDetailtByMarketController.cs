using DongA.Bussiness;
using DongA.Entities;
using DongAERP.Areas.Admin.Models;
using Kendo.Mvc.Extensions;
using Kendo.Mvc.UI;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace DongAERP.Areas.Admin.Controllers
{
    public class ReportDetailtByMarketController : Controller
    {
        // GET: Admin/ReportDetailtByMarket
        public ActionResult Index()
        {
            return View();
        }

        /// <summary>
        /// Màn hình báo cáo cho ngày
        /// </summary>
        /// <returns></returns>
        public ActionResult MarketForTotal()
        {
            string nameUrl = "Báo cáo chi tiết/Theo doanh số chi trả/Theo thị trường/Loại hình dịch vụ/Tất cả";
            ViewBag.NameURL = nameUrl;
            return View();
        }

        /// <summary>
        /// Màn hình báo cáo chi tiết theo ngày
        /// </summary>
        /// <returns></returns>
        public ActionResult MarketForOne()
        {
            string nameUrl = "Báo cáo chi tiết/Theo doanh số chi trả/Theo thị trường/Loại hình dịch vụ/Từng thị trường";
            ViewBag.NameURL = nameUrl;
            return View();
        }

        public ActionResult ReportDetailtGradationCompare()
        {
            string nameUrl = "Báo cáo chi tiết/Theo doanh số chi trả/Theo thị trường/Loại hình dịch vụ/Từng thị trường";
            ViewBag.NameURL = nameUrl;
            // Khởi tạo datatable
            DataTable table = new DataTable();
            // Tạo các cột cho datatable
            table.Columns.Add("MarketName", typeof(String));
            table.Columns.Add("CQ1", typeof(double));
            table.Columns.Add("CN1", typeof(double));
            table.Columns.Add("CK1", typeof(double));
            table.Columns.Add("TDS1", typeof(double));

            table.Columns.Add("CQ2", typeof(double));
            table.Columns.Add("CN2", typeof(double));
            table.Columns.Add("CK2", typeof(double));
            table.Columns.Add("TDS2", typeof(double));
            return View(table);
        }

        public ActionResult ReportDetailtCompareMonthForAll()
        {
            string nameUrl = "Báo cáo chi tiết/Theo Thị trường/So sánh tháng/Tất cả";
            ViewBag.NameURL = nameUrl;
            // Khởi tạo datatable
            DataTable table = new DataTable();
            // Tạo các cột cho datatable
            table.Columns.Add("MarketName", typeof(String));
            table.Columns.Add("CQ1", typeof(double));
            table.Columns.Add("CN1", typeof(double));
            table.Columns.Add("CK1", typeof(double));
            table.Columns.Add("TDS1", typeof(double));

            table.Columns.Add("CQ2", typeof(double));
            table.Columns.Add("CN2", typeof(double));
            table.Columns.Add("CK2", typeof(double));
            table.Columns.Add("TDS2", typeof(double));

            table.Columns.Add("CQ3", typeof(double));
            table.Columns.Add("CN3", typeof(double));
            table.Columns.Add("CK3", typeof(double));
            table.Columns.Add("TDS3", typeof(double));
            return View(table);
        }

        public ActionResult ReportDetailtCompareMonthForOne()
        {
            string nameUrl = "Báo cáo chi tiết/Theo Thị trường/So sánh tháng/Từng thị trường";
            ViewBag.NameURL = nameUrl;
            // Khởi tạo datatable
            DataTable table = new DataTable();
            // Tạo các cột cho datatable
            table.Columns.Add("PartnerName", typeof(String));
            table.Columns.Add("CQ1", typeof(double));
            table.Columns.Add("CN1", typeof(double));
            table.Columns.Add("CK1", typeof(double));
            table.Columns.Add("TDS1", typeof(double));

            table.Columns.Add("CQ2", typeof(double));
            table.Columns.Add("CN2", typeof(double));
            table.Columns.Add("CK2", typeof(double));
            table.Columns.Add("TDS2", typeof(double));

            table.Columns.Add("CQ3", typeof(double));
            table.Columns.Add("CN3", typeof(double));
            table.Columns.Add("CK3", typeof(double));
            table.Columns.Add("TDS3", typeof(double));
            return View(table);
        }

        public ActionResult ReportDetailtGradationCompareForOne()
        {
            string nameUrl = "Báo cáo chi tiết/Theo doanh số chi trả/Theo thị trường/Loại hình dịch vụ/Từng thị trường";
            ViewBag.NameURL = nameUrl;
            // Khởi tạo datatable
            DataTable table = new DataTable();
            // Tạo các cột cho datatable
            table.Columns.Add("PartnerName", typeof(String));
            table.Columns.Add("CQ1", typeof(double));
            table.Columns.Add("CN1", typeof(double));
            table.Columns.Add("CK1", typeof(double));
            table.Columns.Add("TDS1", typeof(double));

            table.Columns.Add("CQ2", typeof(double));
            table.Columns.Add("CN2", typeof(double));
            table.Columns.Add("CK2", typeof(double));
            table.Columns.Add("TDS2", typeof(double));
            return View(table);
        }

        /// <summary>
        /// Bảng hiển thị thông tin các giao dịch qua các ngày
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult MarketForTotal([DataSourceRequest]DataSourceRequest request, string reportTypeID)
        {
            List<ReportDetailtServiceType> listData = new ReportBL().GetListReportDetailtForDay(reportTypeID);
            int count = 1;
            foreach (ReportDetailtServiceType item in listData)
            {
                item.STT = (count++).ToString();
                item.TongDS = item.DSChiNha + item.DSChiQuay + item.DSCK;
            }

            ReportDetailtServiceType dataItem = new ReportDetailtServiceType()
            {
                STT = "Tổng",
                DSChiQuay = listData.Sum(x => x.DSChiQuay),
                DSChiNha = listData.Sum(x => x.DSChiNha),
                DSCK = listData.Sum(x => x.DSCK),
                TongDS = listData.Sum(x => x.TongDS)
            };

            listData.Add(dataItem);

            return Json(listData.ToDataSourceResult(request), JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// Search report day theo ngày nhập vào
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public ActionResult SearchMarketForTotal([DataSourceRequest]DataSourceRequest request, DateTime fromDay, DateTime toDay, string reportTypeID, string marketID)
        {
            List<ReportDetailtSTMarket> listData = new ReportBL().SearchReportDetailtForDay(fromDay, toDay, reportTypeID, marketID);

            foreach (ReportDetailtSTMarket item in listData)
            {
                item.TongDS = item.DSChiNha + item.DSChiQuay + item.DSCK;
            }

            ReportDetailtSTMarket dataItem = new ReportDetailtSTMarket()
            {
                MarketName = "Tổng",
                DSChiQuay = listData.Sum(x => x.DSChiQuay),
                DSChiNha = listData.Sum(x => x.DSChiNha),
                DSCK = listData.Sum(x => x.DSCK),
                TongDS = listData.Sum(x => x.TongDS)
            };

            listData.Add(dataItem);

            return Json(listData.ToDataSourceResult(request), JsonRequestBehavior.AllowGet);
        }

        ///// <summary>
        ///// Bảng hiển thị thông tin các giao dịch qua các ngày
        ///// </summary>
        ///// <returns></returns>
        ///// <history>
        /////     [Truong Lam]   Created [10/06/2020]
        ///// </history>
        //[HttpPost]
        //public ActionResult MarketForPartner([DataSourceRequest]DataSourceRequest request, string reportTypeID)
        //{
        //    List<ReportDetailtServiceType> listData = new ReportBL().MarketForPartner(reportTypeID);
        //    foreach (ReportDetailtServiceType item in listData)
        //    {
        //        item.STT = string.Concat("Tháng ", item.Month, "/", item.Year);
        //        item.TongDS = item.DSChiNha + item.DSChiQuay + item.DSCK;
        //    }

        //    ReportDetailtServiceType dataItem = new ReportDetailtServiceType()
        //    {
        //        ReportID = "Tổng",
        //        DSChiQuay = listData.Sum(x => x.DSChiQuay),
        //        DSChiNha = listData.Sum(x => x.DSChiNha),
        //        DSCK = listData.Sum(x => x.DSCK),
        //        TongDS = listData.Sum(x => x.TongDS)
        //    };

        //    listData.Add(dataItem);

        //    return Json(listData.ToDataSourceResult(request), JsonRequestBehavior.AllowGet);
        //}

        /// <summary>
        /// Danh sách các thị trường
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Trường Lãm]   Created [10/08/2020]
        /// </history>
        public JsonResult ListMarket()
        {
            List<Market> list = new List<Market>()
            {
                new Market()
                {
                    MarketCode = "0",
                    MarketName = "Tất cả"
                },
                new Market()
                {
                    MarketCode = "1",
                    MarketName = "Thị trường Châu Á"
                }
            };

            //list = new ReportBL().ListMarket();
            //list.Add(marketItem);

            return Json(list, JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// Danh sách các thị trường theo đối tác
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Trường Lãm]   Created [10/08/2020]
        /// </history>
        public JsonResult ListMarketForPartner()
        {
            List<Market> list = new ReportBL().ListMarket();
            List<Market> listMarket = new List<Market>();
            // Get danh sách các thị trường chính
            if (list.Count > 0)
            {
                listMarket = list.Where(x => x.ParentCode == null).ToList();
            }
            return Json(listMarket, JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// Bảng hiển thị thông tin các giao dịch qua các ngày
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult MarketForOne([DataSourceRequest]DataSourceRequest request, string reportTypeID)
        {
            List<ReportDetailtServiceType> listData = new ReportBL().GetListReportDetailtForOneMarket(reportTypeID);
            int count = 1;
            foreach (ReportDetailtServiceType item in listData)
            {
                item.STT = (count++).ToString();
                item.TongDS = item.DSChiNha + item.DSChiQuay + item.DSCK;
            }

            ReportDetailtServiceType dataItem = new ReportDetailtServiceType()
            {
                STT = "Tổng",
                DSChiQuay = listData.Sum(x => x.DSChiQuay),
                DSChiNha = listData.Sum(x => x.DSChiNha),
                DSCK = listData.Sum(x => x.DSCK),
                TongDS = listData.Sum(x => x.TongDS)
            };

            listData.Add(dataItem);

            return Json(listData.ToDataSourceResult(request), JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// Search report day theo ngày nhập vào
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public ActionResult SearchMarketForOne([DataSourceRequest]DataSourceRequest request, DateTime fromDay, DateTime toDay, string reportTypeID, string marketID)
        {
            List<ReportDetailtServiceType> listData = new ReportBL().SearchMarketForOne(fromDay, toDay, reportTypeID, marketID);
            int count = 1;

            foreach (ReportDetailtServiceType item in listData)
            {
                if (string.IsNullOrEmpty(item.MarketName))
                {
                    item.MarketName = "Tất cả Thị trường";
                }
                item.STT = (count++).ToString();
                item.TongDS = item.DSChiNha + item.DSChiQuay + item.DSCK;
            }

            return Json(listData.ToDataSourceResult(request), JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// Get data cho báo cáo so sánh giai đoạn
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult ReportDetailtGradationCompareForAll([DataSourceRequest]DataSourceRequest request, string reportTypeID)
        {
            // Danh sach của data gradation gồm key và value

            string[] ArrayData = { "1", "3 tháng đầu năm" };

            int toYear = DateTime.Now.Year;
            List<ReportDetailtServiceType> listDataGradation = new ReportBL().ReportDetailtGradationCompareForAll(toYear, int.Parse(ArrayData[0]), reportTypeID);

            foreach (ReportDetailtServiceType item in listDataGradation)
            {
                item.TongDS = item.DSChiQuay + item.DSChiNha + item.DSCK;
            }

            // Danh sách mã thị trường của Tất cả
            List<string> listMarket = new List<string>() { "003", "005", "001", "002", "014", "004" };

            // Khởi tạo datatable
            DataTable table = new DataTable();
            // Tạo các cột cho datatable
            table.Columns.Add("MarketName", typeof(String));
            table.Columns.Add("CQ1", typeof(double));
            table.Columns.Add("CN1", typeof(double));
            table.Columns.Add("CK1", typeof(double));
            table.Columns.Add("TDS1", typeof(double));

            table.Columns.Add("CQ2", typeof(double));
            table.Columns.Add("CN2", typeof(double));
            table.Columns.Add("CK2", typeof(double));
            table.Columns.Add("TDS2", typeof(double));

            foreach (string item in listMarket)
            {
                // Cùng kì
                ReportDetailtServiceType dataItemLastYear = listDataGradation.Find(x => x.MarketCode == item && x.Year == (toYear - 1).ToString());
                ReportDetailtServiceType dataItemYear = listDataGradation.Find(x => x.MarketCode == item && x.Year == toYear.ToString());

                // add item vào table
                table.Rows.Add(dataItemLastYear.MarketName, dataItemLastYear.DSChiQuay, dataItemLastYear.DSChiNha, dataItemLastYear.DSCK, dataItemLastYear.TongDS
                    , dataItemYear.DSChiQuay, dataItemYear.DSChiNha, dataItemYear.DSCK, dataItemYear.TongDS);
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
            table.Rows.Add(row);

            //ViewBag.ListDataGradation = listDataGradation;
            return Json(table.ToDataSourceResult(request), JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// Get data cho việc vẽ biểu đồ cột chồng cho so sánh giai đoạn
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult ColumnChartGradationCompareStack(string reportTypeID)
        {
            DateTime now = DateTime.Now;
            // Với mặt định chọn giai đoạn 3 tháng đầu năm
            int gradationID = 1;
            List<ReportDetailtServiceType> listDataGradation = new ReportBL().ReportDetailtGradationCompareForAllPercent(now.Year, gradationID, reportTypeID);

            // Số mảng cần tạo
            int arrayCount = 3;
            GradationCharColumn[] arrayGradation = new GradationCharColumn[listDataGradation.Count * arrayCount];
            int count = 0;
            // group theo nhóm
            int tooltip = 1;
            foreach (ReportDetailtServiceType item in listDataGradation)
            {
                // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                arrayGradation[count] = new GradationCharColumn()
                {
                    Serie = item.MarketName,
                    Segmento = string.Concat("Chi Quầy ", item.Year),
                    Valor1 = item.DSChiQuay,
                    Tooltip = tooltip
                };

                count++;
                arrayGradation[count] = new GradationCharColumn()
                {
                    Serie = item.MarketName,
                    Segmento = string.Concat("Chi Nhà ", item.Year),
                    Valor1 = item.DSChiNha,
                    Tooltip = tooltip
                };

                count++;
                arrayGradation[count] = new GradationCharColumn()
                {
                    Serie = item.MarketName,
                    Segmento = string.Concat("DSCK ", item.Year),
                    Valor1 = item.DSCK,
                    Tooltip = tooltip
                };
                count++;
                tooltip++;
            }

            if (listDataGradation.Count == 0)
            {
                arrayGradation = new GradationCharColumn[1];
                arrayGradation[0] = new GradationCharColumn()
                {
                    Serie = "1",
                    Segmento = "1",
                    Valor1 = 0,
                    Tooltip = 1

                };
            }

            return Json(arrayGradation);
        }

        /// <summary>
        /// Get data cho việc vẽ biểu đồ cột cho so sánh giai đoạn
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult ColumnsChartGradationCompare(string reportTypeID)
        {
            int gradationID = 1;
            int year = DateTime.Today.Year;
            List<ReportDetailtServiceType> listDataGradation = new ReportBL().ReportDetailtGradationCompareForAll(year, gradationID, reportTypeID);

            // Số record của mảng
            int countArray = 3;
            GradationCompare[] arrayGradation = new GradationCompare[countArray * listDataGradation.Count];
            int count = 0;
            foreach (ReportDetailtServiceType item in listDataGradation)
            {
                // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                arrayGradation[count] = new GradationCompare()
                {
                    NameGradationCompare = item.MarketName,
                    amount = item.DSChiQuay,
                    NameType = string.Format("Chi Quầy {0}", item.Year)
                };

                count++;
                arrayGradation[count] = new GradationCompare()
                {
                    NameGradationCompare = item.MarketName,
                    amount = item.DSChiNha,
                    NameType = string.Format("Chi Nhà {0}", item.Year)
                };

                count++;
                arrayGradation[count] = new GradationCompare()
                {
                    NameGradationCompare = item.MarketName,
                    amount = item.DSCK,
                    NameType = string.Format("DSCK {0}", item.Year)
                };

                // Tăng count lên 1 đơn vị
                count++;
                year = year - 1;
            }

            if (arrayGradation == null)
            {
                arrayGradation = new GradationCompare[1];
                arrayGradation[0] = new GradationCompare()
                {
                    NameGradationCompare = "1",
                    NameType = ""

                };
            }

            return Json(arrayGradation);
        }

        /// <summary>
        /// Get data cho việc vẽ biểu đồ cột cho so sánh theo giai đoạn
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult SearchColumnChartReportForGradation([DataSourceRequest]DataSourceRequest request, int gradation, int year, string reportTypeID)
        {
            List<ReportDetailtServiceType> listDataGradation = new ReportBL().ReportDetailtGradationCompareForAll(year, gradation, reportTypeID);

            // Số record của mảng
            int countArray = 3;
            GradationCompare[] arrayGradation = new GradationCompare[countArray * listDataGradation.Count];
            int count = 0;
            foreach (ReportDetailtServiceType item in listDataGradation)
            {
                // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                arrayGradation[count] = new GradationCompare()
                {
                    NameGradationCompare = item.MarketName,
                    amount = item.DSChiQuay,
                    NameType = string.Format("Chi Quầy {0}", item.Year)
                };

                count++;
                arrayGradation[count] = new GradationCompare()
                {
                    NameGradationCompare = item.MarketName,
                    amount = item.DSChiNha,
                    NameType = string.Format("Chi Nhà {0}", item.Year)
                };

                count++;
                arrayGradation[count] = new GradationCompare()
                {
                    NameGradationCompare = item.MarketName,
                    amount = item.DSCK,
                    NameType = string.Format("DSCK {0}", item.Year)
                };

                // Tăng count lên 1 đơn vị
                count++;
                year = year - 1;
            }

            if (arrayGradation == null)
            {
                arrayGradation = new GradationCompare[1];
                arrayGradation[0] = new GradationCompare()
                {
                    NameGradationCompare = "1",
                    NameType = ""

                };
            }

            return Json(arrayGradation);
        }

        /// <summary>
        /// Get data cho việc vẽ biểu đồ cột cho so sánh theo giai đoạn
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult SearchColumnChartReportForGradationPercent([DataSourceRequest]DataSourceRequest request, int gradation, int year, string reportTypeID)
        {
            List<ReportDetailtServiceType> listDataGradation = new ReportBL().ReportDetailtGradationCompareForAllPercent(year, gradation, reportTypeID);

            // Số mảng cần tạo
            int arrayCount = 3;
            GradationCharColumn[] arrayGradation = new GradationCharColumn[listDataGradation.Count * arrayCount];
            int count = 0;
            // group theo nhóm
            int tooltip = 1;
            foreach (ReportDetailtServiceType item in listDataGradation)
            {
                // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                arrayGradation[count] = new GradationCharColumn()
                {
                    Serie = item.MarketName,
                    Segmento = string.Concat("Chi Quầy ", item.Year),
                    Valor1 = item.DSChiQuay,
                    Tooltip = tooltip
                };

                count++;
                arrayGradation[count] = new GradationCharColumn()
                {
                    Serie = item.MarketName,
                    Segmento = string.Concat("Chi Nhà ", item.Year),
                    Valor1 = item.DSChiNha,
                    Tooltip = tooltip
                };

                count++;
                arrayGradation[count] = new GradationCharColumn()
                {
                    Serie = item.MarketName,
                    Segmento = string.Concat("DSCK ", item.Year),
                    Valor1 = item.DSCK,
                    Tooltip = tooltip
                };
                count++;
                tooltip++;
            }

            if (listDataGradation.Count == 0)
            {
                arrayGradation = new GradationCharColumn[1];
                arrayGradation[0] = new GradationCharColumn()
                {
                    Serie = "1",
                    Segmento = "1",
                    Valor1 = 0,
                    Tooltip = 1

                };
            }

            return Json(arrayGradation);
        }

        /// <summary>
        /// Get data cho việc vẽ biểu đồ cột cho so sánh theo giai đoạn
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult SearchGridReportForGradation([DataSourceRequest]DataSourceRequest request, int gradation, int year, string reportTypeID)
        {
            List<ReportDetailtServiceType> listDataGradation = new ReportBL().ReportDetailtGradationCompareForAll(year, gradation, reportTypeID);

            foreach (ReportDetailtServiceType item in listDataGradation)
            {
                item.TongDS = item.DSChiQuay + item.DSChiNha + item.DSCK;
            }

            // Danh sách mã thị trường của Tất cả
            List<string> listMarket = new List<string>() { "003", "005", "001", "002", "014", "004" };

            // Khởi tạo datatable
            DataTable table = new DataTable();
            // Tạo các cột cho datatable
            table.Columns.Add("MarketName", typeof(String));
            table.Columns.Add("CQ1", typeof(double));
            table.Columns.Add("CN1", typeof(double));
            table.Columns.Add("CK1", typeof(double));
            table.Columns.Add("TDS1", typeof(double));

            table.Columns.Add("CQ2", typeof(double));
            table.Columns.Add("CN2", typeof(double));
            table.Columns.Add("CK2", typeof(double));
            table.Columns.Add("TDS2", typeof(double));

            if (listDataGradation.Count() > 0)
            {
                foreach (string item in listMarket)
                {
                    // Cùng kì
                    ReportDetailtServiceType dataItemLastYear = listDataGradation.Find(x => x.MarketCode == item && x.Year == (year - 1).ToString());
                    ReportDetailtServiceType dataItemYear = listDataGradation.Find(x => x.MarketCode == item && x.Year == year.ToString());

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

                    // add item vào table
                    table.Rows.Add(dataItemLastYear.MarketName, dataItemLastYear.DSChiQuay, dataItemLastYear.DSChiNha, dataItemLastYear.DSCK, dataItemLastYear.TongDS
                        , dataItemYear.DSChiQuay, dataItemYear.DSChiNha, dataItemYear.DSCK, dataItemYear.TongDS);
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
                table.Rows.Add(row);
            }

            return Json(table.ToDataSourceResult(request), JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// Get data cho báo cáo so sánh giai đoạn
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult ReportDetailtGradationCompareForOne([DataSourceRequest]DataSourceRequest request, string reportTypeID, string marketID)
        {
            // Danh sach của data gradation gồm key và value

            string[] ArrayData = { "1", "3 tháng đầu năm" };

            int year = DateTime.Now.Year;
            // Giá trị ban đầu
            marketID = "001";
            List<ReportDetailtServiceType> listDataGradation = new ReportBL().ReportDetailtGradationCompareForOne(year, int.Parse(ArrayData[0]), reportTypeID, marketID);

            foreach (ReportDetailtServiceType item in listDataGradation)
            {
                item.TongDS = item.DSChiQuay + item.DSChiNha + item.DSCK;
            }

            // Khởi tạo datatable
            DataTable table = new DataTable();
            // Tạo các cột cho datatable
            table.Columns.Add("PartnerName", typeof(String));
            table.Columns.Add("CQ1", typeof(double));
            table.Columns.Add("CN1", typeof(double));
            table.Columns.Add("CK1", typeof(double));
            table.Columns.Add("TDS1", typeof(double));

            table.Columns.Add("CQ2", typeof(double));
            table.Columns.Add("CN2", typeof(double));
            table.Columns.Add("CK2", typeof(double));
            table.Columns.Add("TDS2", typeof(double));

            foreach (ReportDetailtServiceType item in listDataGradation)
            {
                // Cùng kì
                ReportDetailtServiceType dataItemLastYear = listDataGradation.Find(x => x.PartnerCode == item.PartnerCode && x.Year == (year - 1).ToString());
                ReportDetailtServiceType dataItemYear = listDataGradation.Find(x => x.PartnerCode == item.PartnerCode && x.Year == year.ToString());

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

                // Check tồn tại của item
                string value = string.Format("PartnerName='{0}'", item.PartnerName);
                DataRow[] foundRows = table.Select(value);
                if (dataItemLastYear != null && dataItemYear != null && foundRows.Count() == 0)
                {
                    // add item vào table
                    table.Rows.Add(dataItemLastYear.PartnerName, dataItemLastYear.DSChiQuay, dataItemLastYear.DSChiNha, dataItemLastYear.DSCK, dataItemLastYear.TongDS
                        , dataItemYear.DSChiQuay, dataItemYear.DSChiNha, dataItemYear.DSCK, dataItemYear.TongDS);
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
            table.Rows.Add(row);
            return Json(table.ToDataSourceResult(request), JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// Get data cho việc vẽ biểu đồ cột chồng cho so sánh giai đoạn
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult ColumnChartGradationCompareStackForOne(string reportTypeID)
        {
            int year = DateTime.Now.Year;
            // Với mặt định chọn giai đoạn 3 tháng đầu năm
            int gradationID = 1;
            // Giá trị ban đầu
            string marketID = "001";
            List<ReportDetailtServiceType> listDataGradation = new ReportBL().ReportDetailtGradationCompareForOne(year, gradationID, reportTypeID, marketID);
            List<ReportDetailtServiceType> listDataGradationClone = new List<ReportDetailtServiceType>(listDataGradation);

            foreach (ReportDetailtServiceType item in listDataGradationClone)
            {
                ReportDetailtServiceType dataItemYear = listDataGradation.Find(x => x.PartnerCode == item.PartnerCode && x.Year == year.ToString());
                ReportDetailtServiceType dataItemLastYear = listDataGradation.Find(x => x.PartnerCode == item.PartnerCode && x.Year == (year - 1).ToString());

                // Trường hợp năm trước có đối tác và năm nay không có
                if (dataItemLastYear != null && dataItemYear == null)
                {
                    dataItemYear = new ReportDetailtServiceType();
                    dataItemYear.PartnerCode = dataItemLastYear.PartnerCode;
                    dataItemYear.PartnerName = dataItemLastYear.PartnerName;
                    dataItemYear.DSChiQuay = 0;
                    dataItemYear.DSChiNha = 0;
                    dataItemYear.DSCK = 0;
                    dataItemYear.Year = year.ToString();

                    // Add item thiếu vào
                    listDataGradation.Add(dataItemYear);
                }

                // Trường hợp năm trước không có đối tác và năm nay có đối tác
                if (dataItemYear != null && dataItemLastYear == null)
                {
                    dataItemLastYear = new ReportDetailtServiceType();
                    dataItemLastYear.PartnerCode = dataItemYear.PartnerCode;
                    dataItemLastYear.PartnerName = dataItemYear.PartnerName;
                    dataItemLastYear.DSChiQuay = 0;
                    dataItemLastYear.DSChiNha = 0;
                    dataItemLastYear.DSCK = 0;
                    dataItemLastYear.Year = (year - 1).ToString();

                    // Add item thiếu vào
                    listDataGradation.Add(dataItemLastYear);
                }
            }


            // Số mảng cần tạo
            int arrayCount = 1;
            GradationCharColumn[] arrayGradation = new GradationCharColumn[listDataGradation.Count * arrayCount];
            int count = 0;
            // group theo nhóm
            int tooltip = 1;
            foreach (ReportDetailtServiceType item in listDataGradation)
            {
                item.TongDS = item.DSChiQuay + item.DSChiNha + item.DSCK;

                // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                arrayGradation[count] = new GradationCharColumn()
                {
                    Serie = string.Format("Tổng 3T {0}", item.Year),
                    Segmento = item.PartnerName,
                    Valor1 = item.TongDS,
                    //Tooltip = tooltip
                };
                count++;
                //tooltip++;
            }

            if (listDataGradation.Count == 0)
            {
                arrayGradation = new GradationCharColumn[1];
                arrayGradation[0] = new GradationCharColumn()
                {
                    Serie = "1",
                    Segmento = "1",
                    Valor1 = 0,
                    Tooltip = 1

                };
            }

            return Json(arrayGradation);
        }

        /// <summary>
        /// Get data cho việc vẽ biểu đồ cột cho so sánh giai đoạn
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult ColumnsChartGradationCompareForOne(string reportTypeID)
        {
            int gradationID = 1;
            int year = DateTime.Today.Year;
            string marketID = "001";
            List<ReportDetailtServiceType> listDataGradation = new ReportBL().ReportDetailtGradationCompareForOne(year, gradationID, reportTypeID, marketID);

            // Số record của mảng
            int countArray = 3;
            GradationCompare[] arrayGradation = new GradationCompare[countArray * listDataGradation.Count];
            int count = 0;
            foreach (ReportDetailtServiceType item in listDataGradation)
            {
                // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                arrayGradation[count] = new GradationCompare()
                {
                    NameGradationCompare = item.PartnerName,
                    amount = item.DSChiQuay,
                    NameType = string.Format("Chi Quầy {0}", item.Year)
                };

                count++;
                arrayGradation[count] = new GradationCompare()
                {
                    NameGradationCompare = item.PartnerName,
                    amount = item.DSChiNha,
                    NameType = string.Format("Chi Nhà {0}", item.Year)
                };

                count++;
                arrayGradation[count] = new GradationCompare()
                {
                    NameGradationCompare = item.PartnerName,
                    amount = item.DSCK,
                    NameType = string.Format("DSCK {0}", item.Year)
                };

                // Tăng count lên 1 đơn vị
                count++;
                year = year - 1;
            }

            if (arrayGradation == null)
            {
                arrayGradation = new GradationCompare[1];
                arrayGradation[0] = new GradationCompare()
                {
                    NameGradationCompare = "1",
                    NameType = ""

                };
            }

            return Json(arrayGradation);
        }

        /// <summary>
        /// Get data cho việc vẽ biểu đồ cột cho so sánh theo giai đoạn
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult SearchGridReportForGradationForOne([DataSourceRequest]DataSourceRequest request, int gradation, int year, string reportTypeID, string marketID)
        {
            List<ReportDetailtServiceType> listDataGradation = new ReportBL().ReportDetailtGradationCompareForOne(year, gradation, reportTypeID, marketID);

            foreach (ReportDetailtServiceType item in listDataGradation)
            {
                item.TongDS = item.DSChiQuay + item.DSChiNha + item.DSCK;
            }

            // Khởi tạo datatable
            DataTable table = new DataTable();
            // Tạo các cột cho datatable
            table.Columns.Add("PartnerName", typeof(String));
            table.Columns.Add("CQ1", typeof(double));
            table.Columns.Add("CN1", typeof(double));
            table.Columns.Add("CK1", typeof(double));
            table.Columns.Add("TDS1", typeof(double));

            table.Columns.Add("CQ2", typeof(double));
            table.Columns.Add("CN2", typeof(double));
            table.Columns.Add("CK2", typeof(double));
            table.Columns.Add("TDS2", typeof(double));

            foreach (ReportDetailtServiceType item in listDataGradation)
            {
                // Cùng kì
                ReportDetailtServiceType dataItemLastYear = listDataGradation.Find(x => x.PartnerCode == item.PartnerCode && x.Year == (year - 1).ToString());
                ReportDetailtServiceType dataItemYear = listDataGradation.Find(x => x.PartnerCode == item.PartnerCode && x.Year == year.ToString());

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

                // Check tồn tại của item
                string value = string.Format("PartnerName='{0}'", item.PartnerName);
                DataRow[] foundRows = table.Select(value);
                if (dataItemLastYear != null && dataItemYear != null && foundRows.Count() == 0)
                {
                    // add item vào table
                    table.Rows.Add(dataItemLastYear.PartnerName, dataItemLastYear.DSChiQuay, dataItemLastYear.DSChiNha, dataItemLastYear.DSCK, dataItemLastYear.TongDS
                        , dataItemYear.DSChiQuay, dataItemYear.DSChiNha, dataItemYear.DSCK, dataItemYear.TongDS);
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
            table.Rows.Add(row);

            return Json(table.ToDataSourceResult(request), JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// Get data cho việc vẽ biểu đồ cột cho so sánh theo giai đoạn
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult SearchColumnsChartGradationCompareForOne([DataSourceRequest]DataSourceRequest request, int gradation, int year, string reportTypeID, string marketID)
        {
            List<ReportDetailtServiceType> listDataGradation = new ReportBL().ReportDetailtGradationCompareForOne(year, gradation, reportTypeID, marketID);

            // Số record của mảng
            int countArray = 3;
            GradationCompare[] arrayGradation = new GradationCompare[countArray * listDataGradation.Count];
            int count = 0;
            foreach (ReportDetailtServiceType item in listDataGradation)
            {
                // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                arrayGradation[count] = new GradationCompare()
                {
                    NameGradationCompare = item.PartnerName,
                    amount = item.DSChiQuay,
                    NameType = string.Format("Chi Quầy {0}", item.Year)
                };

                count++;
                arrayGradation[count] = new GradationCompare()
                {
                    NameGradationCompare = item.PartnerName,
                    amount = item.DSChiNha,
                    NameType = string.Format("Chi Nhà {0}", item.Year)
                };

                count++;
                arrayGradation[count] = new GradationCompare()
                {
                    NameGradationCompare = item.PartnerName,
                    amount = item.DSCK,
                    NameType = string.Format("DSCK {0}", item.Year)
                };

                // Tăng count lên 1 đơn vị
                count++;
                year = year - 1;
            }

            if (arrayGradation == null)
            {
                arrayGradation = new GradationCompare[1];
                arrayGradation[0] = new GradationCompare()
                {
                    NameGradationCompare = "1",
                    NameType = ""

                };
            }

            return Json(arrayGradation);
        }

        /// <summary>
        /// Get data cho việc vẽ biểu đồ cột cho so sánh theo giai đoạn
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult SearchColumnChartGradationCompareStackForOne([DataSourceRequest]DataSourceRequest request, int gradation, int year, string reportTypeID, string marketID)
        {
            List<ReportDetailtServiceType> listDataGradation = new ReportBL().ReportDetailtGradationCompareForOne(year, gradation, reportTypeID, marketID);
            List<ReportDetailtServiceType> listDataGradationClone = new List<ReportDetailtServiceType>(listDataGradation);

            foreach (ReportDetailtServiceType item in listDataGradationClone)
            {
                ReportDetailtServiceType dataItemYear = listDataGradation.Find(x => x.PartnerCode == item.PartnerCode && x.Year == year.ToString());
                ReportDetailtServiceType dataItemLastYear = listDataGradation.Find(x => x.PartnerCode == item.PartnerCode && x.Year == (year - 1).ToString());

                // Trường hợp năm trước có đối tác và năm nay không có
                if (dataItemLastYear != null && dataItemYear == null)
                {
                    dataItemYear = new ReportDetailtServiceType();
                    dataItemYear.PartnerCode = dataItemLastYear.PartnerCode;
                    dataItemYear.PartnerName = dataItemLastYear.PartnerName;
                    dataItemYear.DSChiQuay = 0;
                    dataItemYear.DSChiNha = 0;
                    dataItemYear.DSCK = 0;
                    dataItemYear.Year = year.ToString();

                    // Add item thiếu vào
                    listDataGradation.Add(dataItemYear);
                }

                // Trường hợp năm trước không có đối tác và năm nay có đối tác
                if (dataItemYear != null && dataItemLastYear == null)
                {
                    dataItemLastYear = new ReportDetailtServiceType();
                    dataItemLastYear.PartnerCode = dataItemYear.PartnerCode;
                    dataItemLastYear.PartnerName = dataItemYear.PartnerName;
                    dataItemLastYear.DSChiQuay = 0;
                    dataItemLastYear.DSChiNha = 0;
                    dataItemLastYear.DSCK = 0;
                    dataItemLastYear.Year = (year - 1).ToString();

                    // Add item thiếu vào
                    listDataGradation.Add(dataItemLastYear);
                }
            }


            // Số mảng cần tạo
            int arrayCount = 1;
            GradationCharColumn[] arrayGradation = new GradationCharColumn[listDataGradation.Count * arrayCount];
            int count = 0;
            // group theo nhóm
            int tooltip = 1;
            foreach (ReportDetailtServiceType item in listDataGradation)
            {
                item.TongDS = item.DSChiQuay + item.DSChiNha + item.DSCK;

                // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                arrayGradation[count] = new GradationCharColumn()
                {
                    Serie = string.Format("Tổng 3T {0}", item.Year),
                    Segmento = item.PartnerName,
                    Valor1 = item.TongDS,
                    //Tooltip = tooltip
                };
                count++;
                //tooltip++;
            }

            if (listDataGradation.Count == 0)
            {
                arrayGradation = new GradationCharColumn[1];
                arrayGradation[0] = new GradationCharColumn()
                {
                    Serie = "1",
                    Segmento = "1",
                    Valor1 = 0,
                    Tooltip = 1

                };
            }

            return Json(arrayGradation);
        }

        /// <summary>
        /// Get data cho báo cáo so sánh giai đoạn
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult ReportDetailtGradationCompareForOneCompare([DataSourceRequest]DataSourceRequest request, string reportTypeID, string marketID)
        {
            // Danh sach của data gradation gồm key và value

            string[] ArrayData = { "1", "3 tháng đầu năm" };

            int year = DateTime.Now.Year;
            // Giá trị ban đầu
            marketID = "001";
            List<ReportDetailtServiceType> listDataGradation = new ReportBL().ReportDetailtGradationCompareForOne(year, int.Parse(ArrayData[0]), reportTypeID, marketID);

            foreach (ReportDetailtServiceType item in listDataGradation)
            {
                item.TongDS = item.DSChiQuay + item.DSChiNha + item.DSCK;
            }

            // Khởi tạo datatable
            DataTable table = new DataTable();
            // Tạo các cột cho datatable
            table.Columns.Add("PartnerName", typeof(String));
            table.Columns.Add("CQ1", typeof(double));
            table.Columns.Add("CN1", typeof(double));
            table.Columns.Add("CK1", typeof(double));
            table.Columns.Add("TDS1", typeof(double));

            foreach (ReportDetailtServiceType item in listDataGradation)
            {
                // Cùng kì
                ReportDetailtServiceType dataItemLastYear = listDataGradation.Find(x => x.PartnerCode == item.PartnerCode && x.Year == (year - 1).ToString());
                ReportDetailtServiceType dataItemYear = listDataGradation.Find(x => x.PartnerCode == item.PartnerCode && x.Year == year.ToString());

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

                // Check tồn tại của item
                string value = string.Format("PartnerName='{0}'", item.PartnerName);
                DataRow[] foundRows = table.Select(value);

                if (dataItemLastYear != null && dataItemYear != null && foundRows.Count() == 0)
                {
                    // add item vào table
                    table.Rows.Add(dataItemLastYear.PartnerName, dataItemYear.DSChiQuay - dataItemLastYear.DSChiQuay
                        , dataItemYear.DSChiNha - dataItemLastYear.DSChiNha, dataItemYear.DSCK - dataItemLastYear.DSCK, dataItemYear.TongDS - dataItemLastYear.TongDS);
                }
            }

            DataRow row = table.NewRow();
            row["PartnerName"] = "Tổng";
            row["CQ1"] = table.Compute("Sum(CQ1)", "");
            row["CN1"] = table.Compute("Sum(CN1)", "");
            row["CK1"] = table.Compute("Sum(CK1)", "");
            row["TDS1"] = table.Compute("Sum(TDS1)", "");
            table.Rows.Add(row);

            //ViewBag.ListDataGradation = listDataGradation;
            return Json(table.ToDataSourceResult(request), JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// Get data cho việc vẽ biểu đồ cột cho so sánh theo giai đoạn
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult SearchGridReportForGradationForOneCompare([DataSourceRequest]DataSourceRequest request, int gradation, int year, string reportTypeID, string marketID)
        {
            List<ReportDetailtServiceType> listDataGradation = new ReportBL().ReportDetailtGradationCompareForOne(year, gradation, reportTypeID, marketID);

            foreach (ReportDetailtServiceType item in listDataGradation)
            {
                item.TongDS = item.DSChiQuay + item.DSChiNha + item.DSCK;
            }

            // Khởi tạo datatable
            DataTable table = new DataTable();
            // Tạo các cột cho datatable
            table.Columns.Add("PartnerName", typeof(String));
            table.Columns.Add("CQ1", typeof(double));
            table.Columns.Add("CN1", typeof(double));
            table.Columns.Add("CK1", typeof(double));
            table.Columns.Add("TDS1", typeof(double));

            foreach (ReportDetailtServiceType item in listDataGradation)
            {
                // Cùng kì
                ReportDetailtServiceType dataItemLastYear = listDataGradation.Find(x => x.PartnerCode == item.PartnerCode && x.Year == (year - 1).ToString());
                ReportDetailtServiceType dataItemYear = listDataGradation.Find(x => x.PartnerCode == item.PartnerCode && x.Year == year.ToString());

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

                // Check tồn tại của item
                string value = string.Format("PartnerName='{0}'", item.PartnerName);
                DataRow[] foundRows = table.Select(value);

                if (dataItemLastYear != null && dataItemYear != null && foundRows.Count() == 0)
                {
                    // add item vào table
                    table.Rows.Add(dataItemLastYear.PartnerName, dataItemYear.DSChiQuay - dataItemLastYear.DSChiQuay
                        , dataItemYear.DSChiNha - dataItemLastYear.DSChiNha, dataItemYear.DSCK - dataItemLastYear.DSCK, dataItemYear.TongDS - dataItemLastYear.TongDS);
                }
            }

            DataRow row = table.NewRow();
            row["PartnerName"] = "Tổng";
            row["CQ1"] = table.Compute("Sum(CQ1)", "");
            row["CN1"] = table.Compute("Sum(CN1)", "");
            row["CK1"] = table.Compute("Sum(CK1)", "");
            row["TDS1"] = table.Compute("Sum(TDS1)", "");
            table.Rows.Add(row);

            //ViewBag.ListDataGradation = listDataGradation;
            return Json(table.ToDataSourceResult(request), JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// Get data cho việc vẽ biểu đồ tròn cho so sánh giai đoạn theo năm hiện tại
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult DataGradationComparePieYear(string reportTypeID)
        {
            int year = DateTime.Today.Year;
            // Xử lý gradation sau
            // Set giá trị gradation cho báo cáo theo giai đoạn
            int gradation = 1;
            string marketID = "001";
            List<ReportDetailtServiceType> listDataGradation = new ReportBL().ReportDetailtGradationCompareForOne(year, gradation, reportTypeID, marketID);
            // Get dữ liệu của năm hiện tại
            List<ReportDetailtServiceType> listData = listDataGradation.Where(x => x.Year == year.ToString()).ToList();

            foreach (ReportDetailtServiceType item in listData)
            {
                item.TongDS = item.DSChiQuay + item.DSChiNha + item.DSCK;
            }

            double sumTongDS = listData.Sum(x => x.TongDS);

            GradationChartPie[] arrayGradation = new GradationChartPie[listData.Count()];
            int count = 0;
            List<string> listColor = new List<string>() { "#FFBF00", "#40FF00", "#2ECCFA", "#9A2EFE", "#FE2EF7", "#0000FF", "#08088A", "#B40431", "#6E6E6E" };

            foreach (ReportDetailtServiceType item in listData)
            {
                // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                arrayGradation[count] = new GradationChartPie()
                {
                    category = item.PartnerName,
                    value = sumTongDS == 0 ? 0 : Math.Round((item.TongDS / sumTongDS * 100), 2, MidpointRounding.ToEven),
                    color = listColor[count]
                };
                count++;
            }

            return Json(arrayGradation);
        }

        /// <summary>
        /// Get data cho việc vẽ biểu đồ tròn cho so sánh giai đoạn theo năm hiện tại
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult SearchDataGradationComparePieYear([DataSourceRequest]DataSourceRequest request, int gradation, int year, string reportTypeID, string marketID)
        {
            List<ReportDetailtServiceType> listDataGradation = new ReportBL().ReportDetailtGradationCompareForOne(year, gradation, reportTypeID, marketID);
            // Get dữ liệu của năm hiện tại
            List<ReportDetailtServiceType> listData = listDataGradation.Where(x => x.Year == year.ToString()).ToList();

            foreach (ReportDetailtServiceType item in listData)
            {
                item.TongDS = item.DSChiQuay + item.DSChiNha + item.DSCK;
            }

            double sumTongDS = listData.Sum(x => x.TongDS);

            GradationChartPie[] arrayGradation = new GradationChartPie[listData.Count()];
            int count = 0;
            List<string> listColor = new List<string>() { "#FFBF00", "#40FF00", "#2ECCFA", "#9A2EFE", "#FE2EF7", "#0000FF", "#08088A", "#B40431", "#6E6E6E" };

            foreach (ReportDetailtServiceType item in listData)
            {
                // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                arrayGradation[count] = new GradationChartPie()
                {
                    category = item.PartnerName,
                    value = sumTongDS == 0 ? 0 : Math.Round((item.TongDS / sumTongDS) * 100, 2, MidpointRounding.ToEven),
                    color = listColor[count]
                };
                count++;
            }

            return Json(arrayGradation);
        }

        /// <summary>
        /// Get data cho việc vẽ biểu đồ tròn cho so sánh giai đoạn theo năm hiện tại
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult DataGradationComparePieLastYear(string reportTypeID)
        {
            int year = DateTime.Today.Year;
            // Xử lý gradation sau
            // Set giá trị gradation cho báo cáo theo giai đoạn
            int gradation = 1;
            string marketID = "001";
            List<ReportDetailtServiceType> listDataGradation = new ReportBL().ReportDetailtGradationCompareForOne(year, gradation, reportTypeID, marketID);
            // Get dữ liệu của năm hiện tại
            List<ReportDetailtServiceType> listData = listDataGradation.Where(x => x.Year == (year - 1).ToString()).ToList();

            foreach (ReportDetailtServiceType item in listData)
            {
                item.TongDS = item.DSChiQuay + item.DSChiNha + item.DSCK;
            }

            double sumTongDS = listData.Sum(x => x.TongDS);

            GradationChartPie[] arrayGradation = new GradationChartPie[listData.Count()];
            int count = 0;
            List<string> listColor = new List<string>() { "#FFBF00", "#40FF00", "#2ECCFA", "#9A2EFE", "#FE2EF7", "#0000FF", "#08088A", "#B40431", "#6E6E6E", "#610B0B", "#8A4B08", "#088A4B", "#8A0829", "#2ECCFA", "#4B088A", "#B404AE" };

            foreach (ReportDetailtServiceType item in listData)
            {
                // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                arrayGradation[count] = new GradationChartPie()
                {
                    category = item.PartnerName,
                    value = sumTongDS == 0 ? 0 : Math.Round((item.TongDS / sumTongDS) * 100, 2, MidpointRounding.ToEven),
                    color = listColor[count]
                };
                count++;
            }

            return Json(arrayGradation);
        }

        /// <summary>
        /// Get data cho việc vẽ biểu đồ tròn cho so sánh giai đoạn theo năm hiện tại
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult SearchDataGradationComparePieLastYear([DataSourceRequest]DataSourceRequest request, int gradation, int year, string reportTypeID, string marketID)
        {
            List<ReportDetailtServiceType> listDataGradation = new ReportBL().ReportDetailtGradationCompareForOne(year, gradation, reportTypeID, marketID);
            // clone Object
            List<ReportDetailtServiceType> listDataGradationClone = new List<ReportDetailtServiceType>(listDataGradation);

            // Xử lý dữ liệu trường hợp các đối tác lấy lên giữa 2 năm không đồng nhất
            foreach(ReportDetailtServiceType item in listDataGradationClone)
            {
                // Cùng kì
                ReportDetailtServiceType dataItemLastYear = listDataGradationClone.Find(x => x.PartnerCode == item.PartnerCode && x.Year == (year - 1).ToString());
                ReportDetailtServiceType dataItemYear = listDataGradationClone.Find(x => x.PartnerCode == item.PartnerCode && x.Year == year.ToString());

                // Trường hợp năm không có đối tác
                if (dataItemLastYear == null && dataItemYear != null)
                {
                    dataItemLastYear = new ReportDetailtServiceType();
                    dataItemLastYear.MarketName = dataItemYear.MarketName;
                    dataItemLastYear.DSChiQuay = 0;
                    dataItemLastYear.DSChiNha = 0;
                    dataItemLastYear.DSCK = 0;
                    dataItemLastYear.Year = (year - 1).ToString();
                    listDataGradation.Add(dataItemLastYear);
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
                    listDataGradation.Add(dataItemYear);
                }
            }
            // Get dữ liệu của năm hiện tại
            List<ReportDetailtServiceType> listData = listDataGradation.Where(x => x.Year == (year - 1).ToString()).ToList();

            foreach (ReportDetailtServiceType item in listData)
            {
                item.TongDS = item.DSChiQuay + item.DSChiNha + item.DSCK;
            }

            double sumTongDS = listData.Sum(x => x.TongDS);

            GradationChartPie[] arrayGradation = new GradationChartPie[listData.Count()];
            int count = 0;
            List<string> listColor = new List<string>() { "#FFBF00", "#40FF00", "#2ECCFA", "#9A2EFE", "#FE2EF7", "#0000FF", "#08088A", "#B40431", "#6E6E6E", "#610B0B", "#8A4B08", "#088A4B", "#8A0829", "#2ECCFA", "#4B088A", "#B404AE" };

            foreach (ReportDetailtServiceType item in listData)
            {
                // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                arrayGradation[count] = new GradationChartPie()
                {
                    category = item.PartnerName,
                    value = sumTongDS == 0 ? 0 : Math.Round((item.TongDS / sumTongDS) * 100, 2, MidpointRounding.ToEven),
                    color = listColor[count]
                };
                count++;
            }

            return Json(arrayGradation);
        }

        /// <summary>
        /// Get data cho việc vẽ biểu đồ tròn cho so sánh giai đoạn theo năm hiện tại
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult DataDetailtGridGradationComparePercent([DataSourceRequest]DataSourceRequest request, string reportTypeID)
        {
            int year = DateTime.Today.Year;
            // Xử lý gradation sau
            // Set giá trị gradation cho báo cáo theo giai đoạn
            int gradation = 1;
            string marketID = "001";
            List<ReportDetailtServiceType> listDataGradation = new ReportBL().ReportDetailtGradationCompareForOne(year, gradation, reportTypeID, marketID);

            // clone Object
            List<ReportDetailtServiceType> listDataGradationClone = new List<ReportDetailtServiceType>(listDataGradation);
            foreach (ReportDetailtServiceType item in listDataGradationClone)
            {
                // Get dữ liệu của năm hiện tại
                ReportDetailtServiceType listDataYear = listDataGradation.Find(x => x.Year == year.ToString() && x.PartnerCode == item.PartnerCode);
                // Get dữ liệu của năm trước
                ReportDetailtServiceType listDataLastYear = listDataGradation.Find(x => x.Year == (year - 1).ToString() && x.PartnerCode == item.PartnerCode);

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
                    listDataGradation.Add(listDataLastYear);
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
                    listDataGradation.Add(listDataYear);
                }
            }

            foreach (ReportDetailtServiceType item in listDataGradation)
            {
                item.TongDS = item.DSChiQuay + item.DSChiNha + item.DSCK;
            }

            double sumTongDSYear = listDataGradation.Where(x => x.Year == year.ToString()).Sum(x => x.TongDS);
            double sumTongDSLastYear = listDataGradation.Where(x => x.Year == (year - 1).ToString()).Sum(x => x.TongDS);

            List<ReportDetailtServiceType> listDataConvert = new List<ReportDetailtServiceType>();

            // Khởi tạo datatable
            DataTable table = new DataTable();
            // Tạo các cột cho datatable
            table.Columns.Add("STT", typeof(String));
            table.Columns.Add("PartnerName", typeof(String));
            table.Columns.Add("LK1", typeof(double));
            table.Columns.Add("LK2", typeof(double));

            int count = 1;
            foreach (ReportDetailtServiceType item in listDataGradation)
            {
                // Get dữ liệu của năm hiện tại
                ReportDetailtServiceType listDataYear = listDataGradation.Find(x => x.Year == year.ToString() && x.PartnerCode == item.PartnerCode);
                // Get dữ liệu của năm trước
                ReportDetailtServiceType listDataLastYear = listDataGradation.Find(x => x.Year == (year - 1).ToString() && x.PartnerCode == item.PartnerCode);

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

            DataRow row = table.NewRow();
            row["STT"] = "";
            row["PartnerName"] = "Tổng";
            row["LK1"] = 100;
            row["LK2"] = 100;
            table.Rows.Add(row);

            return Json(table.ToDataSourceResult(request), JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// Get data cho việc vẽ biểu đồ tròn cho so sánh giai đoạn theo năm hiện tại
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult SearchDataDetailtGridGradationComparePercent([DataSourceRequest]DataSourceRequest request, int gradation, int year, string reportTypeID, string marketID)
        {
            List<ReportDetailtServiceType> listDataGradation = new ReportBL().ReportDetailtGradationCompareForOne(year, gradation, reportTypeID, marketID);

            // clone Object
            List<ReportDetailtServiceType> listDataGradationClone = new List<ReportDetailtServiceType>(listDataGradation);
            foreach (ReportDetailtServiceType item in listDataGradationClone)
            {
                // Get dữ liệu của năm hiện tại
                ReportDetailtServiceType listDataYear = listDataGradation.Find(x => x.Year == year.ToString() && x.PartnerCode == item.PartnerCode);
                // Get dữ liệu của năm trước
                ReportDetailtServiceType listDataLastYear = listDataGradation.Find(x => x.Year == (year - 1).ToString() && x.PartnerCode == item.PartnerCode);

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
                    listDataGradation.Add(listDataLastYear);
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
                    listDataGradation.Add(listDataYear);
                }
            }

            foreach (ReportDetailtServiceType item in listDataGradation)
            {
                item.TongDS = item.DSChiQuay + item.DSChiNha + item.DSCK;
            }

            double sumTongDSYear = listDataGradation.Where(x => x.Year == year.ToString()).Sum(x => x.TongDS);
            double sumTongDSLastYear = listDataGradation.Where(x => x.Year == (year - 1).ToString()).Sum(x => x.TongDS);

            List<ReportDetailtServiceType> listDataConvert = new List<ReportDetailtServiceType>();

            // Khởi tạo datatable
            DataTable table = new DataTable();
            // Tạo các cột cho datatable
            table.Columns.Add("STT", typeof(String));
            table.Columns.Add("PartnerName", typeof(String));
            table.Columns.Add("LK1", typeof(double));
            table.Columns.Add("LK2", typeof(double));

            int count = 1;
            foreach (ReportDetailtServiceType item in listDataGradation)
            {
                // Get dữ liệu của năm hiện tại
                ReportDetailtServiceType listDataYear = listDataGradation.Find(x => x.Year == year.ToString() && x.PartnerCode == item.PartnerCode);
                // Get dữ liệu của năm trước
                ReportDetailtServiceType listDataLastYear = listDataGradation.Find(x => x.Year == (year - 1).ToString() && x.PartnerCode == item.PartnerCode);

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

            DataRow row = table.NewRow();
            row["STT"] = "";
            row["PartnerName"] = "Tổng";
            row["LK1"] = 100;
            row["LK2"] = 100;
            table.Rows.Add(row);

            return Json(table.ToDataSourceResult(request), JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// Get data cho báo cáo so sánh giai đoạn
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult ReportDetailtCompareMonthForAll([DataSourceRequest]DataSourceRequest request, string reportTypeID)
        {
            // Danh sach của data gradation gồm key và value
            int toYear = DateTime.Now.Year;
            int toMonth = DateTime.Now.Month;
            toMonth = 6;

            List<ReportDetailtSTMarket> listDataCompareMonth = new ReportBL().ReportDetailtCompareMonthForAll(toYear, toMonth, reportTypeID);

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

            foreach (string item in listMarket)
            {
                // Cùng kì
                ReportDetailtSTMarket dataItemLastYear = listDataCompareMonth.Find(x => x.MarketCode == item && x.Month == toMonth.ToString() && x.Year == (toYear - 1).ToString());
                ReportDetailtSTMarket dataItemYear = listDataCompareMonth.Find(x => x.MarketCode == item && x.Month == toMonth.ToString() && x.Year == toYear.ToString());
                ReportDetailtSTMarket dataItemLastMonth = listDataCompareMonth.Find(x => x.MarketCode == item && x.Month == (toMonth - 1).ToString() && x.Year == toYear.ToString());

                if (dataItemLastYear != null && dataItemYear != null && dataItemLastMonth != null)
                {
                    // add item vào table
                    table.Rows.Add(dataItemYear.MarketName, dataItemYear.DSChiQuay, dataItemYear.DSChiNha, dataItemYear.DSCK, dataItemYear.TongDS
                        , dataItemLastMonth.DSChiQuay, dataItemLastMonth.DSChiNha, dataItemLastMonth.DSCK, dataItemLastMonth.TongDS
                        , dataItemLastYear.DSChiQuay, dataItemLastYear.DSChiNha, dataItemLastYear.DSCK, dataItemLastYear.TongDS);
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
            table.Rows.Add(row);

            //ViewBag.ListDataGradation = listDataGradation;
            return Json(table.ToDataSourceResult(request), JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// Get data cho báo cáo chi tiết so sánh theo tháng
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult SearchReportDetailtCompareMonthForAll([DataSourceRequest]DataSourceRequest request, int month, int year, string reportTypeID)
        {
            List<ReportDetailtSTMarket> listDataCompareMonth = new ReportBL().ReportDetailtCompareMonthForAll(year, month, reportTypeID);

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

            foreach (string item in listMarket)
            {
                // Cùng kì
                ReportDetailtSTMarket dataItemLastYear = listDataCompareMonth.Find(x => x.MarketCode == item && x.Month == month.ToString() && x.Year == (year - 1).ToString());
                ReportDetailtSTMarket dataItemYear = listDataCompareMonth.Find(x => x.MarketCode == item && x.Month == month.ToString() && x.Year == year.ToString());
                ReportDetailtSTMarket dataItemLastMonth = listDataCompareMonth.Find(x => x.MarketCode == item && x.Month == (month - 1).ToString() && x.Year == year.ToString());

                if (dataItemLastYear != null && dataItemYear != null && dataItemLastMonth != null)
                {
                    // add item vào table
                    table.Rows.Add(dataItemYear.MarketName, dataItemYear.DSChiQuay, dataItemYear.DSChiNha, dataItemYear.DSCK, dataItemYear.TongDS
                        , dataItemLastMonth.DSChiQuay, dataItemLastMonth.DSChiNha, dataItemLastMonth.DSCK, dataItemLastMonth.TongDS
                        , dataItemLastYear.DSChiQuay, dataItemLastYear.DSChiNha, dataItemLastYear.DSCK, dataItemLastYear.TongDS);
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
            table.Rows.Add(row);

            //ViewBag.ListDataGradation = listDataGradation;
            return Json(table.ToDataSourceResult(request), JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// Get data cho báo cáo chi tiết theo tháng
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult ReportDetailtCompareMonthForAllCompare([DataSourceRequest]DataSourceRequest request, string reportTypeID)
        {
            int toYear = DateTime.Now.Year;
            int toMonth = DateTime.Now.Month;
            toMonth = 6;

            List<ReportDetailtSTMarket> listDataCompareMonth = new ReportBL().ReportDetailtCompareMonthForAll(toYear, toMonth, reportTypeID);

            foreach (ReportDetailtSTMarket item in listDataCompareMonth)
            {
                item.TongDS = item.DSChiQuay + item.DSChiNha + item.DSCK;
            }

            // Khởi tạo datatable
            DataTable table = new DataTable();
            // Tạo các cột cho datatable
            table.Columns.Add("MarketName", typeof(String));

            // So sánh với tháng trước
            table.Columns.Add("CQ1", typeof(double));
            table.Columns.Add("CN1", typeof(double));
            table.Columns.Add("CK1", typeof(double));
            table.Columns.Add("TDS1", typeof(double));

            // So sánh với cùng kì năm trước
            table.Columns.Add("CQ2", typeof(double));
            table.Columns.Add("CN2", typeof(double));
            table.Columns.Add("CK2", typeof(double));
            table.Columns.Add("TDS2", typeof(double));

            foreach (ReportDetailtSTMarket item in listDataCompareMonth)
            {
                // Cùng kì
                ReportDetailtSTMarket dataItemLastYear = listDataCompareMonth.Find(x => x.MarketCode == item.MarketCode && x.Month == item.Month && x.Year == (toYear - 1).ToString());
                ReportDetailtSTMarket dataItemMonth = listDataCompareMonth.Find(x => x.MarketCode == item.MarketCode && x.Month == item.Month && x.Year == toYear.ToString());
                ReportDetailtSTMarket dataItemLastMonth = listDataCompareMonth.Find(x => x.MarketCode == item.MarketCode && x.Month == (Int32.Parse(item.Month) - 1).ToString() && x.Year == toYear.ToString());

                // Check tồn tại của item
                string value = string.Format("MarketName='{0}'", item.MarketName);
                DataRow[] foundRows = table.Select(value);

                if (dataItemLastYear != null && dataItemMonth != null && dataItemLastMonth != null && foundRows.Count() == 0)
                {
                    // add item vào table
                    table.Rows.Add(dataItemLastYear.MarketName
                        , dataItemMonth.DSChiQuay - dataItemLastMonth.DSChiQuay, dataItemMonth.DSChiNha - dataItemLastMonth.DSChiNha, dataItemMonth.DSCK - dataItemLastMonth.DSCK, dataItemMonth.TongDS - dataItemLastMonth.TongDS
                        , dataItemMonth.DSChiQuay - dataItemLastYear.DSChiQuay, dataItemMonth.DSChiNha - dataItemLastYear.DSChiNha, dataItemMonth.DSCK - dataItemLastYear.DSCK, dataItemMonth.TongDS - dataItemLastYear.TongDS);
                }
            }

            DataRow row = table.NewRow();
            row["MarketName"] = "Tổng";
            // So sánh với tháng trước
            row["CQ1"] = table.Compute("Sum(CQ1)", "");
            row["CN1"] = table.Compute("Sum(CN1)", "");
            row["CK1"] = table.Compute("Sum(CK1)", "");
            row["TDS1"] = table.Compute("Sum(TDS1)", "");

            // So sánh với cùng kì năm trươc
            row["CQ2"] = table.Compute("Sum(CQ2)", "");
            row["CN2"] = table.Compute("Sum(CN2)", "");
            row["CK2"] = table.Compute("Sum(CK2)", "");
            row["TDS2"] = table.Compute("Sum(TDS2)", "");
            table.Rows.Add(row);

            //ViewBag.ListDataGradation = listDataGradation;
            return Json(table.ToDataSourceResult(request), JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// Get data cho báo cáo chi tiết theo tháng
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult SearchReportDetailtCompareMonthForAllCompare([DataSourceRequest]DataSourceRequest request, int month, int year, string reportTypeID)
        {
            List<ReportDetailtSTMarket> listDataCompareMonth = new ReportBL().ReportDetailtCompareMonthForAll(year, month, reportTypeID);

            foreach (ReportDetailtSTMarket item in listDataCompareMonth)
            {
                item.TongDS = item.DSChiQuay + item.DSChiNha + item.DSCK;
            }

            // Khởi tạo datatable
            DataTable table = new DataTable();
            // Tạo các cột cho datatable
            table.Columns.Add("MarketName", typeof(String));

            // So sánh với tháng trước
            table.Columns.Add("CQ1", typeof(double));
            table.Columns.Add("CN1", typeof(double));
            table.Columns.Add("CK1", typeof(double));
            table.Columns.Add("TDS1", typeof(double));

            // So sánh với cùng kì năm trước
            table.Columns.Add("CQ2", typeof(double));
            table.Columns.Add("CN2", typeof(double));
            table.Columns.Add("CK2", typeof(double));
            table.Columns.Add("TDS2", typeof(double));

            foreach (ReportDetailtSTMarket item in listDataCompareMonth)
            {
                // Cùng kì
                ReportDetailtSTMarket dataItemLastYear = listDataCompareMonth.Find(x => x.MarketCode == item.MarketCode && x.Month == item.Month && x.Year == (year - 1).ToString());
                ReportDetailtSTMarket dataItemMonth = listDataCompareMonth.Find(x => x.MarketCode == item.MarketCode && x.Month == item.Month && x.Year == year.ToString());
                ReportDetailtSTMarket dataItemLastMonth = listDataCompareMonth.Find(x => x.MarketCode == item.MarketCode && x.Month == (Int32.Parse(item.Month) - 1).ToString() && x.Year == year.ToString());

                // Check tồn tại của item
                string value = string.Format("MarketName='{0}'", item.MarketName);
                DataRow[] foundRows = table.Select(value);

                if (dataItemLastYear != null && dataItemMonth != null && dataItemLastMonth != null && foundRows.Count() == 0)
                {
                    // add item vào table
                    table.Rows.Add(dataItemLastYear.MarketName
                        , dataItemMonth.DSChiQuay - dataItemLastMonth.DSChiQuay, dataItemMonth.DSChiNha - dataItemLastMonth.DSChiNha, dataItemMonth.DSCK - dataItemLastMonth.DSCK, dataItemMonth.TongDS - dataItemLastMonth.TongDS
                        , dataItemMonth.DSChiQuay - dataItemLastYear.DSChiQuay, dataItemMonth.DSChiNha - dataItemLastYear.DSChiNha, dataItemMonth.DSCK - dataItemLastYear.DSCK, dataItemMonth.TongDS - dataItemLastYear.TongDS);
                }
            }

            DataRow row = table.NewRow();
            row["MarketName"] = "Tổng";
            // So sánh với tháng trước
            row["CQ1"] = table.Compute("Sum(CQ1)", "");
            row["CN1"] = table.Compute("Sum(CN1)", "");
            row["CK1"] = table.Compute("Sum(CK1)", "");
            row["TDS1"] = table.Compute("Sum(TDS1)", "");

            // So sánh với cùng kì năm trươc
            row["CQ2"] = table.Compute("Sum(CQ2)", "");
            row["CN2"] = table.Compute("Sum(CN2)", "");
            row["CK2"] = table.Compute("Sum(CK2)", "");
            row["TDS2"] = table.Compute("Sum(TDS2)", "");
            table.Rows.Add(row);

            //ViewBag.ListDataGradation = listDataGradation;
            return Json(table.ToDataSourceResult(request), JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// Get data cho việc vẽ biểu đồ cột chồng cho so sánh giai đoạn
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult ColumnChartStackCompareMonthForAll(string reportTypeID)
        {
            int year = DateTime.Now.Year;
            int toMonth = DateTime.Now.Month;
            toMonth = 6;

            List<ReportDetailtSTMarket> listDataCompareMonth = new ReportBL().ColumnChartStackCompareMonthForAllPercent(year, toMonth, reportTypeID);

            GradationCharColumn[] arrayGradation = null;

            if (listDataCompareMonth.Count > 0)
            {
                List<ReportDetailtSTMarket> listDataYear = listDataCompareMonth.Where(x => x.Month == toMonth.ToString() && x.Year == year.ToString()).ToList();
                List<ReportDetailtSTMarket> listDataLastMonth = listDataCompareMonth.Where(x => x.Month == (toMonth - 1).ToString() && x.Year == year.ToString()).ToList();
                List<ReportDetailtSTMarket> listDataLastYear = listDataCompareMonth.Where(x => x.Month == toMonth.ToString() && x.Year == (year - 1).ToString()).ToList();

                // Trường hợp đủ dữ liệu tháng hiện tại, tháng trước và cùng kì năm trước
                if (listDataYear.Count > 0 && listDataLastMonth.Count > 0 && listDataLastYear.Count > 0)
                {
                    // Số record của mảng
                    int countArray = 3;
                    arrayGradation = new GradationCharColumn[countArray * listDataCompareMonth.Count];
                    int count = 0;
                    int tooltip = 1;

                    foreach (ReportDetailtSTMarket item in listDataCompareMonth)
                    {
                        // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                        arrayGradation[count] = new GradationCharColumn()
                        {
                            Serie = item.MarketName,
                            Segmento = string.Format("Chi Quầy {0}/{1}", item.Month, item.Year),
                            Valor1 = item.DSChiQuay,
                            Tooltip = tooltip
                        };

                        count++;
                        arrayGradation[count] = new GradationCharColumn()
                        {
                            Serie = item.MarketName,
                            Segmento = string.Format("Chi Nhà {0}/{1}", item.Month, item.Year),
                            Valor1 = item.DSChiNha,
                            Tooltip = tooltip
                        };

                        count++;
                        arrayGradation[count] = new GradationCharColumn()
                        {
                            Serie = item.MarketName,
                            Segmento = string.Format("DSCK {0}/{1}", item.Month, item.Year),
                            Valor1 = item.DSCK,
                            Tooltip = tooltip
                        };
                        count++;
                        tooltip++;
                    }
                }
            }

            if (arrayGradation == null)
            {
                arrayGradation = new GradationCharColumn[1];
                arrayGradation[0] = new GradationCharColumn()
                {
                    Serie = "1",
                    Segmento = "1",
                    Valor1 = 0,
                    Tooltip = 1

                };
            }

            return Json(arrayGradation);
        }

        /// <summary>
        /// Get data cho việc vẽ biểu đồ cột chồng cho so sánh giai đoạn
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult SearchColumnChartStackCompareMonthForAll([DataSourceRequest]DataSourceRequest request, int month, int year, string reportTypeID)
        {
            List<ReportDetailtSTMarket> listDataCompareMonth = new ReportBL().ColumnChartStackCompareMonthForAllPercent(year, month, reportTypeID);

            GradationCharColumn[] arrayGradation = null;

            if (listDataCompareMonth.Count > 0)
            {
                List<ReportDetailtSTMarket> listDataYear = listDataCompareMonth.Where(x => x.Month == month.ToString() && x.Year == year.ToString()).ToList();
                List<ReportDetailtSTMarket> listDataLastMonth = listDataCompareMonth.Where(x => x.Month == (month - 1).ToString() && x.Year == year.ToString()).ToList();
                List<ReportDetailtSTMarket> listDataLastYear = listDataCompareMonth.Where(x => x.Month == month.ToString() && x.Year == (year - 1).ToString()).ToList();

                // Trường hợp đủ dữ liệu tháng hiện tại, tháng trước và cùng kì năm trước
                if (listDataYear.Count > 0 && listDataLastMonth.Count > 0 && listDataLastYear.Count > 0)
                {
                    // Số record của mảng
                    int countArray = 3;
                    arrayGradation = new GradationCharColumn[countArray * listDataCompareMonth.Count];
                    int count = 0;
                    int tooltip = 1;

                    foreach (ReportDetailtSTMarket item in listDataCompareMonth)
                    {
                        // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                        arrayGradation[count] = new GradationCharColumn()
                        {
                            Serie = item.MarketName,
                            Segmento = string.Format("Chi Quầy {0}/{1}", item.Month, item.Year),
                            Valor1 = item.DSChiQuay,
                            Tooltip = tooltip
                        };

                        count++;
                        arrayGradation[count] = new GradationCharColumn()
                        {
                            Serie = item.MarketName,
                            Segmento = string.Format("Chi Nhà {0}/{1}", item.Month, item.Year),
                            Valor1 = item.DSChiNha,
                            Tooltip = tooltip
                        };

                        count++;
                        arrayGradation[count] = new GradationCharColumn()
                        {
                            Serie = item.MarketName,
                            Segmento = string.Format("DSCK {0}/{1}", item.Month, item.Year),
                            Valor1 = item.DSCK,
                            Tooltip = tooltip
                        };
                        count++;
                        tooltip++;
                    }
                }
            }

            if (arrayGradation == null)
            {
                arrayGradation = new GradationCharColumn[1];
                arrayGradation[0] = new GradationCharColumn()
                {
                    Serie = "1",
                    Segmento = "1",
                    Valor1 = 0,
                    Tooltip = 1

                };
            }

            return Json(arrayGradation);
        }

        /// <summary>
        /// Get data cho việc vẽ biểu đồ cột cho so sánh giai đoạn
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult ColumnsChartCompareMonthForAll(string reportTypeID)
        {
            int year = DateTime.Now.Year;
            int toMonth = DateTime.Now.Month;
            toMonth = 6;

            List<ReportDetailtSTMarket> listDataCompareMonth = new ReportBL().ReportDetailtCompareMonthForAll(year, toMonth, reportTypeID);

            GradationCompare[] arrayGradation = null;
            if (listDataCompareMonth.Count > 0)
            {
                List<ReportDetailtSTMarket> listDataYear = listDataCompareMonth.Where(x => x.Month == toMonth.ToString() && x.Year == year.ToString()).ToList();
                List<ReportDetailtSTMarket> listDataLastMonth = listDataCompareMonth.Where(x => x.Month == (toMonth - 1).ToString() && x.Year == year.ToString()).ToList();
                List<ReportDetailtSTMarket> listDataLastYear = listDataCompareMonth.Where(x => x.Month == toMonth.ToString() && x.Year == (year - 1).ToString()).ToList();

                // Trường hợp đủ dữ liệu tháng hiện tại, tháng trước và cùng kì năm trước
                if (listDataYear.Count > 0 && listDataLastMonth.Count > 0 && listDataLastYear.Count > 0)
                {
                    // Số record của mảng
                    int countArray = 3;
                    arrayGradation = new GradationCompare[countArray * listDataCompareMonth.Count];
                    int count = 0;

                    foreach (ReportDetailtSTMarket item in listDataCompareMonth)
                    {
                        // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                        arrayGradation[count] = new GradationCompare()
                        {
                            NameGradationCompare = item.MarketName,
                            amount = item.DSChiQuay,
                            NameType = string.Format("Chi Quầy {0}/{1}", item.Month, item.Year)
                        };

                        count++;
                        arrayGradation[count] = new GradationCompare()
                        {
                            NameGradationCompare = item.MarketName,
                            amount = item.DSChiNha,
                            NameType = string.Format("Chi Nhà {0}/{1}", item.Month, item.Year)
                        };

                        count++;
                        arrayGradation[count] = new GradationCompare()
                        {
                            NameGradationCompare = item.MarketName,
                            amount = item.DSCK,
                            NameType = string.Format("DSCK {0}/{1}", item.Month, item.Year)
                        };

                        // Tăng count lên 1 đơn vị
                        count++;
                        year = year - 1;
                    }
                }
            }

            if (arrayGradation == null)
            {
                arrayGradation = new GradationCompare[1];
                arrayGradation[0] = new GradationCompare()
                {
                    NameGradationCompare = "1",
                    NameType = ""

                };
            }

            return Json(arrayGradation);
        }

        /// <summary>
        /// Get data cho việc vẽ biểu đồ cột cho so sánh giai đoạn
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult SearchColumnsChartCompareMonthForAll([DataSourceRequest]DataSourceRequest request, int month, int year, string reportTypeID)
        {
            List<ReportDetailtSTMarket> listDataCompareMonth = new ReportBL().ReportDetailtCompareMonthForAll(year, month, reportTypeID);

            GradationCompare[] arrayGradation = null;
            if (listDataCompareMonth.Count > 0)
            {
                List<ReportDetailtSTMarket> listDataYear = listDataCompareMonth.Where(x => x.Month == month.ToString() && x.Year == year.ToString()).ToList();
                List<ReportDetailtSTMarket> listDataLastMonth = listDataCompareMonth.Where(x => x.Month == (month - 1).ToString() && x.Year == year.ToString()).ToList();
                List<ReportDetailtSTMarket> listDataLastYear = listDataCompareMonth.Where(x => x.Month == month.ToString() && x.Year == (year - 1).ToString()).ToList();

                // Trường hợp đủ dữ liệu tháng hiện tại, tháng trước và cùng kì năm trước
                if (listDataYear.Count > 0 && listDataLastMonth.Count > 0 && listDataLastYear.Count > 0)
                {
                    // Số record của mảng
                    int countArray = 3;
                    arrayGradation = new GradationCompare[countArray * listDataCompareMonth.Count];
                    int count = 0;

                    foreach (ReportDetailtSTMarket item in listDataCompareMonth)
                    {
                        // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                        arrayGradation[count] = new GradationCompare()
                        {
                            NameGradationCompare = item.MarketName,
                            amount = item.DSChiQuay,
                            NameType = string.Format("Chi Quầy {0}/{1}", item.Month, item.Year)
                        };

                        count++;
                        arrayGradation[count] = new GradationCompare()
                        {
                            NameGradationCompare = item.MarketName,
                            amount = item.DSChiNha,
                            NameType = string.Format("Chi Nhà {0}/{1}", item.Month, item.Year)
                        };

                        count++;
                        arrayGradation[count] = new GradationCompare()
                        {
                            NameGradationCompare = item.MarketName,
                            amount = item.DSCK,
                            NameType = string.Format("DSCK {0}/{1}", item.Month, item.Year)
                        };

                        // Tăng count lên 1 đơn vị
                        count++;
                        year = year - 1;
                    }
                }
            }

            if (arrayGradation == null)
            {
                arrayGradation = new GradationCompare[1];
                arrayGradation[0] = new GradationCompare()
                {
                    NameGradationCompare = "1",
                    NameType = ""

                };
            }

            return Json(arrayGradation);
        }

        /// <summary>
        /// Get data cho báo cáo so sánh giai đoạn
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult ReportDetailtCompareMonthForOne([DataSourceRequest]DataSourceRequest request, string reportTypeID)
        {
            // Danh sach của data gradation gồm key và value
            int toYear = DateTime.Now.Year;
            int toMonth = DateTime.Now.Month;
            string marketID = "001";

            // Test
            toMonth = 6;

            List<ReportDetailtServiceType> listDataCompareMonth = new ReportBL().ReportDetailtCompareMonthForOne(toYear, toMonth, reportTypeID, marketID);

            foreach (ReportDetailtServiceType item in listDataCompareMonth)
            {
                item.TongDS = item.DSChiQuay + item.DSChiNha + item.DSCK;
            }

            // Danh sách mã thị trường của Tất cả
            List<string> listMarket = new List<string>() { "003", "005", "001", "002", "014", "004" };

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

            List<string> listTemp = new List<string>();

            foreach (ReportDetailtServiceType item in listDataCompareMonth)
            {
                // Cùng kì
                ReportDetailtServiceType dataItemLastYear = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == toMonth.ToString() && x.Year == (toYear - 1).ToString());
                ReportDetailtServiceType dataItemYear = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == toMonth.ToString() && x.Year == toYear.ToString());
                ReportDetailtServiceType dataItemLastMonth = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == (toMonth - 1).ToString() && x.Year == toYear.ToString());

                if (!listTemp.Contains(item.PartnerCode))
                {
                    // Trường hợp năm trước có đối tác không có
                    if (dataItemLastYear == null)
                    {
                        dataItemLastYear = new ReportDetailtServiceType();
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
                        dataItemLastMonth.PartnerName = item.PartnerName;
                        dataItemLastMonth.DSChiQuay = 0;
                        dataItemLastMonth.DSChiNha = 0;
                        dataItemLastMonth.DSCK = 0;
                        dataItemLastMonth.Year = item.Year;
                        dataItemLastMonth.Month = item.Month;
                    }

                    // Check tồn tại của item
                    string value = string.Format("PartnerName='{0}'", item.PartnerName);
                    DataRow[] foundRows = table.Select(value);

                    if (dataItemLastYear != null && dataItemYear != null && dataItemLastMonth != null && foundRows.Count() == 0)
                    {
                        // add item vào table
                        table.Rows.Add(dataItemYear.PartnerName, dataItemYear.DSChiQuay, dataItemYear.DSChiNha, dataItemYear.DSCK, dataItemYear.TongDS
                            , dataItemLastMonth.DSChiQuay, dataItemLastMonth.DSChiNha, dataItemLastMonth.DSCK, dataItemLastMonth.TongDS
                            , dataItemLastYear.DSChiQuay, dataItemLastYear.DSChiNha, dataItemLastYear.DSCK, dataItemLastYear.TongDS);
                    }
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
            table.Rows.Add(row);

            //ViewBag.ListDataGradation = listDataGradation;
            return Json(table.ToDataSourceResult(request), JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// Get data cho báo cáo so sánh giai đoạn
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult SearchReportDetailtCompareMonthForOne([DataSourceRequest]DataSourceRequest request, int month, int year, string reportTypeID, string marketID)
        {
            List<ReportDetailtServiceType> listDataCompareMonth = new ReportBL().ReportDetailtCompareMonthForOne(year, month, reportTypeID, marketID);

            foreach (ReportDetailtServiceType item in listDataCompareMonth)
            {
                item.TongDS = item.DSChiQuay + item.DSChiNha + item.DSCK;
            }

            // Danh sách mã thị trường của Tất cả
            List<string> listMarket = new List<string>() { "003", "005", "001", "002", "014", "004" };

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
                        dataItemLastMonth.PartnerName = item.PartnerName;
                        dataItemLastMonth.DSChiQuay = 0;
                        dataItemLastMonth.DSChiNha = 0;
                        dataItemLastMonth.DSCK = 0;
                        dataItemLastMonth.Year = item.Year;
                        dataItemLastMonth.Month = item.Month;
                    }

                    // Check tồn tại của item
                    string value = string.Format("PartnerName='{0}'", item.PartnerName);
                    DataRow[] foundRows = table.Select(value);

                    if (dataItemLastYear != null && dataItemYear != null && dataItemLastMonth != null && foundRows.Count() == 0)
                    {
                        // add item vào table
                        table.Rows.Add(dataItemYear.PartnerName, dataItemYear.DSChiQuay, dataItemYear.DSChiNha, dataItemYear.DSCK, dataItemYear.TongDS
                            , dataItemLastMonth.DSChiQuay, dataItemLastMonth.DSChiNha, dataItemLastMonth.DSCK, dataItemLastMonth.TongDS
                            , dataItemLastYear.DSChiQuay, dataItemLastYear.DSChiNha, dataItemLastYear.DSCK, dataItemLastYear.TongDS);
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
            table.Rows.Add(row);

            //ViewBag.ListDataGradation = listDataGradation;
            return Json(table.ToDataSourceResult(request), JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// Get data cho báo cáo so sánh giai đoạn
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult ReportDetailtCompareMonthForOneCompare([DataSourceRequest]DataSourceRequest request, string reportTypeID)
        {
            // Danh sach của data gradation gồm key và value
            int toYear = DateTime.Now.Year;
            int toMonth = DateTime.Now.Month;
            string marketID = "001";
            // Test
            toMonth = 6;

            List<ReportDetailtServiceType> listDataCompareMonth = new ReportBL().ReportDetailtCompareMonthForOne(toYear, toMonth, reportTypeID, marketID);

            foreach (ReportDetailtServiceType item in listDataCompareMonth)
            {
                item.TongDS = item.DSChiQuay + item.DSChiNha + item.DSCK;
            }

            // Khởi tạo datatable
            DataTable table = new DataTable();
            // Tạo các cột cho datatable
            // So sánh với tháng trước
            table.Columns.Add("PartnerName", typeof(String));
            table.Columns.Add("CQ1", typeof(double));
            table.Columns.Add("CN1", typeof(double));
            table.Columns.Add("CK1", typeof(double));
            table.Columns.Add("TDS1", typeof(double));

            // So sánh với cùng kì năm trước
            table.Columns.Add("CQ2", typeof(double));
            table.Columns.Add("CN2", typeof(double));
            table.Columns.Add("CK2", typeof(double));
            table.Columns.Add("TDS2", typeof(double));

            List<string> listTemp = new List<string>();

            foreach (ReportDetailtServiceType item in listDataCompareMonth)
            {
                // Cùng kì
                ReportDetailtServiceType dataItemLastYear = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == toMonth.ToString() && x.Year == (toYear - 1).ToString());
                ReportDetailtServiceType dataItemYear = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == toMonth.ToString() && x.Year == toYear.ToString());
                ReportDetailtServiceType dataItemLastMonth = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == (toMonth - 1).ToString() && x.Year == toYear.ToString());

                if (!listTemp.Contains(item.PartnerCode))
                {
                    // Trường hợp năm không có đối tác
                    if (dataItemLastYear == null)
                    {
                        dataItemLastYear = new ReportDetailtServiceType();
                        dataItemLastYear.PartnerName = item.PartnerName;
                        dataItemLastYear.DSChiQuay = 0;
                        dataItemLastYear.DSChiNha = 0;
                        dataItemLastYear.DSCK = 0;
                        dataItemLastYear.Year = item.Year;
                        dataItemLastYear.Month = item.Month;
                    }

                    // Trường hợp năm hiện tại không có đối tác
                    if (dataItemYear == null)
                    {
                        dataItemYear = new ReportDetailtServiceType();
                        dataItemYear.PartnerName = item.PartnerName;
                        dataItemYear.DSChiQuay = 0;
                        dataItemYear.DSChiNha = 0;
                        dataItemYear.DSCK = 0;
                        dataItemYear.Year = item.Year;
                        dataItemYear.Month = item.Month;
                    }

                    // Trường hợp tháng trước không có
                    if (dataItemLastMonth == null)
                    {
                        dataItemLastMonth = new ReportDetailtServiceType();
                        dataItemLastMonth.PartnerName = item.PartnerName;
                        dataItemLastMonth.DSChiQuay = 0;
                        dataItemLastMonth.DSChiNha = 0;
                        dataItemLastMonth.DSCK = 0;
                        dataItemLastMonth.Year = item.Year;
                        dataItemLastMonth.Month = item.Month;
                    }

                    // Check tồn tại của item
                    string value = string.Format("PartnerName='{0}'", item.PartnerName);
                    DataRow[] foundRows = table.Select(value);

                    if (dataItemLastYear != null && dataItemYear != null && dataItemLastMonth != null && foundRows.Count() == 0)
                    {
                        // add item vào table
                        table.Rows.Add(dataItemYear.PartnerName
                            , dataItemYear.DSChiQuay - dataItemLastMonth.DSChiQuay, dataItemYear.DSChiNha - dataItemLastMonth.DSChiNha, dataItemYear.DSCK - dataItemLastMonth.DSCK, dataItemYear.TongDS - dataItemLastMonth.TongDS
                            , dataItemYear.DSChiQuay - dataItemLastYear.DSChiQuay, dataItemYear.DSChiNha - dataItemLastYear.DSChiNha, dataItemYear.DSCK - dataItemLastYear.DSCK, dataItemYear.TongDS - dataItemLastYear.TongDS);
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
            table.Rows.Add(row);

            //ViewBag.ListDataGradation = listDataGradation;
            return Json(table.ToDataSourceResult(request), JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// Get data cho báo cáo so sánh giai đoạn
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult SearchReportDetailtCompareMonthForOneCompare([DataSourceRequest]DataSourceRequest request, int month, int year, string reportTypeID, string marketID)
        {
            List<ReportDetailtServiceType> listDataCompareMonth = new ReportBL().ReportDetailtCompareMonthForOne(year, month, reportTypeID, marketID);

            foreach (ReportDetailtServiceType item in listDataCompareMonth)
            {
                item.TongDS = item.DSChiQuay + item.DSChiNha + item.DSCK;
            }

            // Khởi tạo datatable
            DataTable table = new DataTable();
            // Tạo các cột cho datatable
            // So sánh với tháng trước
            table.Columns.Add("PartnerName", typeof(String));
            table.Columns.Add("CQ1", typeof(double));
            table.Columns.Add("CN1", typeof(double));
            table.Columns.Add("CK1", typeof(double));
            table.Columns.Add("TDS1", typeof(double));

            // So sánh với cùng kì năm trước
            table.Columns.Add("CQ2", typeof(double));
            table.Columns.Add("CN2", typeof(double));
            table.Columns.Add("CK2", typeof(double));
            table.Columns.Add("TDS2", typeof(double));

            List<string> listTemp = new List<string>();

            foreach (ReportDetailtServiceType item in listDataCompareMonth)
            {
                // Cùng kì
                ReportDetailtServiceType dataItemLastYear = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == month.ToString() && x.Year == (year - 1).ToString());
                ReportDetailtServiceType dataItemYear = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == month.ToString() && x.Year == year.ToString());
                ReportDetailtServiceType dataItemLastMonth = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == (month - 1).ToString() && x.Year == year.ToString());

                if (!listTemp.Contains(item.PartnerCode))
                {
                    // Trường hợp năm không có đối tác
                    if (dataItemLastYear == null)
                    {
                        dataItemLastYear = new ReportDetailtServiceType();
                        dataItemLastYear.PartnerName = item.PartnerName;
                        dataItemLastYear.DSChiQuay = 0;
                        dataItemLastYear.DSChiNha = 0;
                        dataItemLastYear.DSCK = 0;
                        dataItemLastYear.Year = item.Year;
                        dataItemLastYear.Month = item.Month;
                    }

                    // Trường hợp năm hiện tại không có đối tác
                    if (dataItemYear == null)
                    {
                        dataItemYear = new ReportDetailtServiceType();
                        dataItemYear.PartnerName = item.PartnerName;
                        dataItemYear.DSChiQuay = 0;
                        dataItemYear.DSChiNha = 0;
                        dataItemYear.DSCK = 0;
                        dataItemYear.Year = item.Year;
                        dataItemYear.Month = item.Month;
                    }

                    // Trường hợp tháng trước không có
                    if (dataItemLastMonth == null)
                    {
                        dataItemLastMonth = new ReportDetailtServiceType();
                        dataItemLastMonth.PartnerName = item.PartnerName;
                        dataItemLastMonth.DSChiQuay = 0;
                        dataItemLastMonth.DSChiNha = 0;
                        dataItemLastMonth.DSCK = 0;
                        dataItemLastMonth.Year = item.Year;
                        dataItemLastMonth.Month = item.Month;
                    }

                    // Check tồn tại của item
                    string value = string.Format("PartnerName='{0}'", item.PartnerName);
                    DataRow[] foundRows = table.Select(value);

                    if (dataItemLastYear != null && dataItemYear != null && dataItemLastMonth != null && foundRows.Count() == 0)
                    {
                        // add item vào table
                        table.Rows.Add(dataItemYear.PartnerName
                            , dataItemYear.DSChiQuay - dataItemLastMonth.DSChiQuay, dataItemYear.DSChiNha - dataItemLastMonth.DSChiNha, dataItemYear.DSCK - dataItemLastMonth.DSCK, dataItemYear.TongDS - dataItemLastMonth.TongDS
                            , dataItemYear.DSChiQuay - dataItemLastYear.DSChiQuay, dataItemYear.DSChiNha - dataItemLastYear.DSChiNha, dataItemYear.DSCK - dataItemLastYear.DSCK, dataItemYear.TongDS - dataItemLastYear.TongDS);
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
            table.Rows.Add(row);

            //ViewBag.ListDataGradation = listDataGradation;
            return Json(table.ToDataSourceResult(request), JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// Get data cho việc vẽ biểu đồ cột chồng cho so sánh giai đoạn
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult ColumnChartCompareMonthStackForOne(string reportTypeID)
        {
            int toYear = DateTime.Now.Year;
            int toMonth = DateTime.Now.Month;
            string marketID = "001";
            // Test
            toMonth = 6;

            List<ReportDetailtServiceType> listDataCompareMonth = new ReportBL().ReportDetailtCompareMonthForOne(toYear, toMonth, reportTypeID, marketID);

            List<ReportDetailtServiceType> listDataCommpareMonthClone = new List<ReportDetailtServiceType>(listDataCompareMonth);

            List<string> listTemp = new List<string>();

            foreach (ReportDetailtServiceType item in listDataCommpareMonthClone)
            {
                // Cùng kì
                ReportDetailtServiceType dataItemLastYear = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == toMonth.ToString() && x.Year == (toYear - 1).ToString());
                ReportDetailtServiceType dataItemYear = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == toMonth.ToString() && x.Year == toYear.ToString());
                ReportDetailtServiceType dataItemLastMonth = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == (toMonth - 1).ToString() && x.Year == toYear.ToString());

                if (!listTemp.Contains(item.PartnerCode))
                {
                    // Trường hợp năm không có đối tác
                    if (dataItemLastYear == null)
                    {
                        dataItemLastYear = new ReportDetailtServiceType();
                        dataItemLastYear.PartnerName = item.PartnerName;
                        dataItemLastYear.DSChiQuay = 0;
                        dataItemLastYear.DSChiNha = 0;
                        dataItemLastYear.DSCK = 0;
                        dataItemLastYear.Year = (toYear - 1).ToString();
                        dataItemLastYear.Month = toMonth.ToString();
                        listDataCompareMonth.Add(dataItemLastYear);
                    }

                    // Trường hợp năm hiện tại không có đối tác
                    if (dataItemYear == null)
                    {
                        dataItemYear = new ReportDetailtServiceType();
                        dataItemYear.PartnerName = item.PartnerName;
                        dataItemYear.DSChiQuay = 0;
                        dataItemYear.DSChiNha = 0;
                        dataItemYear.DSCK = 0;
                        dataItemYear.Year = toYear.ToString();
                        dataItemYear.Month = toMonth.ToString();
                        listDataCompareMonth.Add(dataItemYear);
                    }

                    // Trường hợp tháng trước không có
                    if (dataItemLastMonth == null)
                    {
                        dataItemLastMonth = new ReportDetailtServiceType();
                        dataItemLastMonth.PartnerName = item.PartnerName;
                        dataItemLastMonth.DSChiQuay = 0;
                        dataItemLastMonth.DSChiNha = 0;
                        dataItemLastMonth.DSCK = 0;
                        dataItemLastMonth.Year = toYear.ToString();
                        dataItemLastMonth.Month = (toMonth - 1).ToString();
                        listDataCompareMonth.Add(dataItemLastMonth);
                    }

                    // Add partnerCode để kiểm tra
                    listTemp.Add(item.PartnerCode);
                }
            }


            // Số mảng cần tạo
            int arrayCount = 1;
            GradationCharColumn[] arrayGradation = new GradationCharColumn[listDataCompareMonth.Count * arrayCount];
            int count = 0;

            foreach (ReportDetailtServiceType item in listDataCompareMonth)
            {
                item.TongDS = item.DSChiQuay + item.DSChiNha + item.DSCK;

                // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                arrayGradation[count] = new GradationCharColumn()
                {
                    Serie = string.Format("Tháng {0}/{1}", item.Month, item.Year),
                    Segmento = item.PartnerName,
                    Valor1 = item.TongDS
                };
                count++;
            }

            if (listDataCompareMonth.Count == 0)
            {
                arrayGradation = new GradationCharColumn[1];
                arrayGradation[0] = new GradationCharColumn()
                {
                    Serie = "1",
                    Segmento = "1",
                    Valor1 = 0,
                    Tooltip = 1

                };
            }

            return Json(arrayGradation);
        }

        /// <summary>
        /// Get data cho việc vẽ biểu đồ cột chồng cho so sánh giai đoạn
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult SearchColumnChartCompareMonthStackForOne([DataSourceRequest]DataSourceRequest request, int month, int year, string reportTypeID, string marketID)
        {

            List<ReportDetailtServiceType> listDataCompareMonth = new ReportBL().ReportDetailtCompareMonthForOne(year, month, reportTypeID, marketID);

            List<ReportDetailtServiceType> listDataCommpareMonthClone = new List<ReportDetailtServiceType>(listDataCompareMonth);
            List<string> listTemp = new List<string>();

            foreach (ReportDetailtServiceType item in listDataCommpareMonthClone)
            {
                if (!listTemp.Contains(item.PartnerCode))
                {
                    // Cùng kì
                    ReportDetailtServiceType dataItemLastYear = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == month.ToString() && x.Year == (year - 1).ToString());
                    ReportDetailtServiceType dataItemYear = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == month.ToString() && x.Year == year.ToString());
                    ReportDetailtServiceType dataItemLastMonth = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == (month - 1).ToString() && x.Year == year.ToString());

                    // Trường hợp năm không có đối tác
                    if (dataItemLastYear == null)
                    {
                        dataItemLastYear = new ReportDetailtServiceType();
                        dataItemLastYear.PartnerName = item.PartnerName;
                        dataItemLastYear.DSChiQuay = 0;
                        dataItemLastYear.DSChiNha = 0;
                        dataItemLastYear.DSCK = 0;
                        dataItemLastYear.Year = (year - 1).ToString();
                        dataItemLastYear.Month = month.ToString();
                        listDataCompareMonth.Add(dataItemLastYear);
                    }

                    // Trường hợp năm hiện tại không có đối tác
                    if (dataItemYear == null)
                    {
                        dataItemYear = new ReportDetailtServiceType();
                        dataItemYear.PartnerName = item.PartnerName;
                        dataItemYear.DSChiQuay = 0;
                        dataItemYear.DSChiNha = 0;
                        dataItemYear.DSCK = 0;
                        dataItemYear.Year = year.ToString();
                        dataItemYear.Month = month.ToString();
                        listDataCompareMonth.Add(dataItemYear);
                    }

                    // Trường hợp tháng trước không có
                    if (dataItemLastMonth == null)
                    {
                        dataItemLastMonth = new ReportDetailtServiceType();
                        dataItemLastMonth.PartnerName = item.PartnerName;
                        dataItemLastMonth.DSChiQuay = 0;
                        dataItemLastMonth.DSChiNha = 0;
                        dataItemLastMonth.DSCK = 0;
                        dataItemLastMonth.Year = year.ToString();
                        dataItemLastMonth.Month = (month - 1).ToString();
                        listDataCompareMonth.Add(dataItemLastMonth);
                    }

                    // Add partnerCode để kiểm tra
                    listTemp.Add(item.PartnerCode);
                }
            }


            // Số mảng cần tạo
            int arrayCount = 1;
            GradationCharColumn[] arrayGradation = new GradationCharColumn[listDataCompareMonth.Count * arrayCount];
            int count = 0;

            foreach (ReportDetailtServiceType item in listDataCompareMonth)
            {
                item.TongDS = item.DSChiQuay + item.DSChiNha + item.DSCK;

                // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                arrayGradation[count] = new GradationCharColumn()
                {
                    Serie = string.Format("Tháng {0}/{1}", item.Month, item.Year),
                    Segmento = item.PartnerName,
                    Valor1 = item.TongDS
                };
                count++;
            }

            if (listDataCompareMonth.Count == 0)
            {
                arrayGradation = new GradationCharColumn[1];
                arrayGradation[0] = new GradationCharColumn()
                {
                    Serie = "1",
                    Segmento = "1",
                    Valor1 = 0,
                    Tooltip = 1

                };
            }

            return Json(arrayGradation);
        }

        /// <summary>
        /// Get data cho việc vẽ biểu đồ cột cho so sánh giai đoạn
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult ColumnsChartCompareMonthForOne(string reportTypeID)
        {
            int year = DateTime.Now.Year;
            int toMonth = DateTime.Now.Month;
            string marketID = "001";
            // Test
            toMonth = 6;

            List<ReportDetailtServiceType> listDataCompareMonth = new ReportBL().ReportDetailtCompareMonthForOne(year, toMonth, reportTypeID, marketID);

            // Số record của mảng
            int countArray = 3;
            GradationCompare[] arrayGradation = new GradationCompare[countArray * listDataCompareMonth.Count];
            int count = 0;
            foreach (ReportDetailtServiceType item in listDataCompareMonth)
            {
                // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                arrayGradation[count] = new GradationCompare()
                {
                    NameGradationCompare = item.PartnerName,
                    amount = item.DSChiQuay,
                    NameType = string.Format("Chi Quầy {0}/{1}", item.Month, item.Year)
                };

                count++;
                arrayGradation[count] = new GradationCompare()
                {
                    NameGradationCompare = item.PartnerName,
                    amount = item.DSChiNha,
                    NameType = string.Format("Chi Nhà {0}/{1}", item.Month, item.Year)
                };

                count++;
                arrayGradation[count] = new GradationCompare()
                {
                    NameGradationCompare = item.PartnerName,
                    amount = item.DSCK,
                    NameType = string.Format("DSCK {0}/{1}", item.Month, item.Year)
                };

                // Tăng count lên 1 đơn vị
                count++;
                year = year - 1;
            }

            if (arrayGradation == null)
            {
                arrayGradation = new GradationCompare[1];
                arrayGradation[0] = new GradationCompare()
                {
                    NameGradationCompare = "1",
                    NameType = ""

                };
            }

            return Json(arrayGradation);
        }


        /// <summary>
        /// Get data cho việc vẽ biểu đồ cột cho so sánh giai đoạn
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult SearchColumnsChartCompareMonthForOne([DataSourceRequest]DataSourceRequest request, int month, int year, string reportTypeID, string marketID)
        {

            List<ReportDetailtServiceType> listDataCompareMonth = new ReportBL().ReportDetailtCompareMonthForOne(year, month, reportTypeID, marketID);

            // Số record của mảng
            int countArray = 3;
            GradationCompare[] arrayGradation = new GradationCompare[countArray * listDataCompareMonth.Count];
            int count = 0;
            foreach (ReportDetailtServiceType item in listDataCompareMonth)
            {
                // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                arrayGradation[count] = new GradationCompare()
                {
                    NameGradationCompare = item.PartnerName,
                    amount = item.DSChiQuay,
                    NameType = string.Format("Chi Quầy {0}/{1}", item.Month, item.Year)
                };

                count++;
                arrayGradation[count] = new GradationCompare()
                {
                    NameGradationCompare = item.PartnerName,
                    amount = item.DSChiNha,
                    NameType = string.Format("Chi Nhà {0}/{1}", item.Month, item.Year)
                };

                count++;
                arrayGradation[count] = new GradationCompare()
                {
                    NameGradationCompare = item.PartnerName,
                    amount = item.DSCK,
                    NameType = string.Format("DSCK {0}/{1}", item.Month, item.Year)
                };

                // Tăng count lên 1 đơn vị
                count++;
                year = year - 1;
            }

            if (arrayGradation == null)
            {
                arrayGradation = new GradationCompare[1];
                arrayGradation[0] = new GradationCompare()
                {
                    NameGradationCompare = "1",
                    NameType = ""

                };
            }

            return Json(arrayGradation);
        }

        /// <summary>
        /// Get data cho việc vẽ biểu đồ tròn cho so sánh theo tháng của tháng hiện tại
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult DataCompareMonthPieYear(string reportTypeID)
        {
            int year = DateTime.Now.Year;
            int toMonth = DateTime.Now.Month;
            string marketID = "001";
            // Test
            toMonth = 6;

            List<ReportDetailtServiceType> listDataCompareMonth = new ReportBL().ReportDetailtCompareMonthForOne(year, toMonth, reportTypeID, marketID);

            // Get dữ liệu của năm hiện tại
            List<ReportDetailtServiceType> listData = listDataCompareMonth.Where(x => x.Year == year.ToString() && x.Month == toMonth.ToString()).ToList();

            foreach (ReportDetailtServiceType item in listData)
            {
                item.TongDS = item.DSChiQuay + item.DSChiNha + item.DSCK;
            }

            double sumTongDS = listData.Sum(x => x.TongDS);

            GradationChartPie[] arrayGradation = new GradationChartPie[listData.Count()];
            int count = 0;
            List<string> listColor = new List<string>() { "#FFBF00", "#40FF00", "#2ECCFA", "#9A2EFE", "#FE2EF7", "#0000FF", "#08088A", "#B40431", "#6E6E6E" };

            foreach (ReportDetailtServiceType item in listData)
            {
                // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                arrayGradation[count] = new GradationChartPie()
                {
                    category = item.PartnerName,
                    value = sumTongDS == 0 ? 0 : Math.Round((item.TongDS / sumTongDS) * 100, 2, MidpointRounding.ToEven),
                    color = listColor[count]
                };
                count++;
            }

            if (arrayGradation.Count() == 0)
            {
                arrayGradation[count] = new GradationChartPie()
                {
                    category = "1",
                    value = 1,
                    color = "1"
                };
            }

            return Json(arrayGradation);
        }

        /// <summary>
        /// Get data cho việc vẽ biểu đồ tròn cho so sánh giai đoạn theo năm hiện tại
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult SearchDataCompareMonthPieYear([DataSourceRequest]DataSourceRequest request, int month, int year, string reportTypeID, string marketID)
        {

            List<ReportDetailtServiceType> listDataCompareMonth = new ReportBL().ReportDetailtCompareMonthForOne(year, month, reportTypeID, marketID);

            // Get dữ liệu của năm hiện tại
            List<ReportDetailtServiceType> listData = listDataCompareMonth.Where(x => x.Year == year.ToString() && x.Month == month.ToString()).ToList();

            foreach (ReportDetailtServiceType item in listData)
            {
                item.TongDS = item.DSChiQuay + item.DSChiNha + item.DSCK;
            }

            double sumTongDS = listData.Sum(x => x.TongDS);

            GradationChartPie[] arrayGradation = new GradationChartPie[listData.Count()];
            int count = 0;
            List<string> listColor = new List<string>() { "#FFBF00", "#40FF00", "#2ECCFA", "#9A2EFE", "#FE2EF7", "#0000FF", "#08088A", "#B40431", "#6E6E6E" };

            foreach (ReportDetailtServiceType item in listData)
            {
                // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                arrayGradation[count] = new GradationChartPie()
                {
                    category = item.PartnerName,
                    value = sumTongDS == 0 ? 0 : Math.Round((item.TongDS / sumTongDS) * 100, 2, MidpointRounding.ToEven),
                    color = listColor[count]
                };
                count++;
            }

            if (arrayGradation.Count() == 0)
            {
                arrayGradation[count] = new GradationChartPie()
                {
                    category = "1",
                    value = 1,
                    color = "1"
                };
            }

            return Json(arrayGradation);
        }

        /// <summary>
        /// Get data cho việc vẽ biểu đồ tròn cho so sánh theo tháng của tháng hiện tại
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult DataCompareMonthPieLastMonth(string reportTypeID)
        {
            int year = DateTime.Now.Year;
            int toMonth = DateTime.Now.Month;
            string marketID = "001";
            // Test
            toMonth = 6;

            List<ReportDetailtServiceType> listDataCompareMonth = new ReportBL().ReportDetailtCompareMonthForOne(year, toMonth, reportTypeID, marketID);

            // Get dữ liệu của năm hiện tại
            List<ReportDetailtServiceType> listData = listDataCompareMonth.Where(x => x.Year == year.ToString() && x.Month == (toMonth - 1).ToString()).ToList();

            foreach (ReportDetailtServiceType item in listData)
            {
                item.TongDS = item.DSChiQuay + item.DSChiNha + item.DSCK;
            }

            double sumTongDS = listData.Sum(x => x.TongDS);

            GradationChartPie[] arrayGradation = new GradationChartPie[listData.Count()];
            int count = 0;
            List<string> listColor = new List<string>() { "#FFBF00", "#40FF00", "#2ECCFA", "#9A2EFE", "#FE2EF7", "#0000FF", "#08088A", "#B40431", "#6E6E6E" };

            foreach (ReportDetailtServiceType item in listData)
            {
                // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                arrayGradation[count] = new GradationChartPie()
                {
                    category = item.PartnerName,
                    value = sumTongDS == 0 ? 0 : Math.Round((item.TongDS / sumTongDS) * 100, 2, MidpointRounding.ToEven),
                    color = listColor[count]
                };
                count++;
            }

            if (arrayGradation.Count() == 0)
            {
                arrayGradation[count] = new GradationChartPie()
                {
                    category = "1",
                    value = 1,
                    color = "1"
                };
            }

            return Json(arrayGradation);
        }

        /// <summary>
        /// Get data cho việc vẽ biểu đồ tròn cho so sánh giai đoạn theo năm hiện tại
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult SearchDataCompareMonthPieLastMonth([DataSourceRequest]DataSourceRequest request, int month, int year, string reportTypeID, string marketID)
        {

            List<ReportDetailtServiceType> listDataCompareMonth = new ReportBL().ReportDetailtCompareMonthForOne(year, month, reportTypeID, marketID);

            // Get dữ liệu của năm hiện tại
            List<ReportDetailtServiceType> listData = listDataCompareMonth.Where(x => x.Year == year.ToString() && x.Month == (month - 1).ToString()).ToList();

            foreach (ReportDetailtServiceType item in listData)
            {
                item.TongDS = item.DSChiQuay + item.DSChiNha + item.DSCK;
            }

            double sumTongDS = listData.Sum(x => x.TongDS);

            GradationChartPie[] arrayGradation = new GradationChartPie[listData.Count()];
            int count = 0;
            List<string> listColor = new List<string>() { "#FFBF00", "#40FF00", "#2ECCFA", "#9A2EFE", "#FE2EF7", "#0000FF", "#08088A", "#B40431", "#6E6E6E" };

            foreach (ReportDetailtServiceType item in listData)
            {
                // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                arrayGradation[count] = new GradationChartPie()
                {
                    category = item.PartnerName,
                    value = sumTongDS == 0 ? 0 : Math.Round((item.TongDS / sumTongDS) * 100, 2, MidpointRounding.ToEven),
                    color = listColor[count]
                };
                count++;
            }

            if (arrayGradation.Count() == 0)
            {
                arrayGradation[count] = new GradationChartPie()
                {
                    category = "1",
                    value = 1,
                    color = "1"
                };
            }

            return Json(arrayGradation);
        }


        /// <summary>
        /// Get data cho việc vẽ biểu đồ tròn cho so sánh theo tháng của tháng hiện tại
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult DataCompareMonthPieLastYear(string reportTypeID)
        {
            int year = DateTime.Now.Year;
            int toMonth = DateTime.Now.Month;
            string marketID = "001";
            // Test
            toMonth = 6;

            List<ReportDetailtServiceType> listDataCompareMonth = new ReportBL().ReportDetailtCompareMonthForOne(year, toMonth, reportTypeID, marketID);

            // Get dữ liệu của năm hiện tại
            List<ReportDetailtServiceType> listData = listDataCompareMonth.Where(x => x.Year == (year - 1).ToString() && x.Month == toMonth.ToString()).ToList();

            foreach (ReportDetailtServiceType item in listData)
            {
                item.TongDS = item.DSChiQuay + item.DSChiNha + item.DSCK;
            }

            double sumTongDS = listData.Sum(x => x.TongDS);

            GradationChartPie[] arrayGradation = new GradationChartPie[listData.Count()];
            int count = 0;
            List<string> listColor = new List<string>() { "#FFBF00", "#40FF00", "#2ECCFA", "#9A2EFE", "#FE2EF7", "#0000FF", "#08088A", "#B40431", "#6E6E6E" };

            foreach (ReportDetailtServiceType item in listData)
            {
                // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                arrayGradation[count] = new GradationChartPie()
                {
                    category = item.PartnerName,
                    value = sumTongDS == 0 ? 0 : Math.Round((item.TongDS / sumTongDS) * 100, 2, MidpointRounding.ToEven),
                    color = listColor[count]
                };
                count++;
            }

            if (arrayGradation.Count() == 0)
            {
                arrayGradation[count] = new GradationChartPie()
                {
                    category = "1",
                    value = 1,
                    color = "1"
                };
            }

            return Json(arrayGradation);
        }

        /// <summary>
        /// Get data cho việc vẽ biểu đồ tròn cho so sánh giai đoạn theo năm hiện tại
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult SearchDataCompareMonthPieLastYear([DataSourceRequest]DataSourceRequest request, int month, int year, string reportTypeID, string marketID)
        {

            List<ReportDetailtServiceType> listDataCompareMonth = new ReportBL().ReportDetailtCompareMonthForOne(year, month, reportTypeID, marketID);

            // Get dữ liệu của năm hiện tại
            List<ReportDetailtServiceType> listData = listDataCompareMonth.Where(x => x.Year == (year - 1).ToString() && x.Month == month.ToString()).ToList();

            foreach (ReportDetailtServiceType item in listData)
            {
                item.TongDS = item.DSChiQuay + item.DSChiNha + item.DSCK;
            }

            double sumTongDS = listData.Sum(x => x.TongDS);

            GradationChartPie[] arrayGradation = new GradationChartPie[listData.Count()];
            int count = 0;
            List<string> listColor = new List<string>() { "#FFBF00", "#40FF00", "#2ECCFA", "#9A2EFE", "#FE2EF7", "#0000FF", "#08088A", "#B40431", "#6E6E6E" };

            foreach (ReportDetailtServiceType item in listData)
            {
                // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                arrayGradation[count] = new GradationChartPie()
                {
                    category = item.PartnerName,
                    value = sumTongDS == 0 ? 0 : Math.Round((item.TongDS / sumTongDS) * 100, 2, MidpointRounding.ToEven),
                    color = listColor[count]
                };
                count++;
            }

            if (arrayGradation.Count() == 0)
            {
                arrayGradation[count] = new GradationChartPie()
                {
                    category = "1",
                    value = 1,
                    color = "1"
                };
            }

            return Json(arrayGradation);
        }

        /// <summary>
        /// Get data cho việc vẽ biểu đồ tròn cho so sánh theo tháng với tháng hiện tại, tháng trước và cùng kì năm trước
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult DataDetailtGridCompareMonthPercent([DataSourceRequest]DataSourceRequest request, string reportTypeID)
        {
            int year = DateTime.Now.Year;
            int toMonth = DateTime.Now.Month;
            string marketID = "001";
            // Test
            toMonth = 6;

            List<ReportDetailtServiceType> listDataCompareMonth = new ReportBL().ReportDetailtCompareMonthForOne(year, toMonth, reportTypeID, marketID);

            List<ReportDetailtServiceType> listDataCommpareMonthClone = new List<ReportDetailtServiceType>(listDataCompareMonth);

            foreach (ReportDetailtServiceType item in listDataCompareMonth)
            {
                item.TongDS = item.DSChiQuay + item.DSChiNha + item.DSCK;
            }

            double sumTongDSYear = listDataCompareMonth.Where(x => x.Year == year.ToString() && x.Month == toMonth.ToString()).Sum(x => x.TongDS);
            double sumTongDSLastYear = listDataCompareMonth.Where(x => x.Year == (year - 1).ToString() && x.Month == toMonth.ToString()).Sum(x => x.TongDS);
            double sumTongDSLastMonth = listDataCompareMonth.Where(x => x.Year == year.ToString() && x.Month == (toMonth - 1).ToString()).Sum(x => x.TongDS);

            List<ReportDetailtServiceType> listDataConvert = new List<ReportDetailtServiceType>();

            // Khởi tạo datatable
            DataTable table = new DataTable();
            // Tạo các cột cho datatable
            table.Columns.Add("STT", typeof(String));
            table.Columns.Add("PartnerName", typeof(String));
            table.Columns.Add("LK1", typeof(double));
            table.Columns.Add("LK2", typeof(double));
            table.Columns.Add("LK3", typeof(double));

            List<string> listTemp = new List<string>();

            int count = 1;
            foreach (ReportDetailtServiceType item in listDataCommpareMonthClone)
            {
                if (!listTemp.Contains(item.PartnerCode))
                {
                    // Cùng kì
                    ReportDetailtServiceType dataItemLastYear = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == toMonth.ToString() && x.Year == (year - 1).ToString());
                    ReportDetailtServiceType dataItemYear = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == toMonth.ToString() && x.Year == year.ToString());
                    ReportDetailtServiceType dataItemLastMonth = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == (toMonth - 1).ToString() && x.Year == year.ToString());

                    // Trường hợp năm không có đối tác
                    if (dataItemLastYear == null)
                    {
                        dataItemLastYear = new ReportDetailtServiceType();
                        dataItemLastYear.PartnerName = item.PartnerName;
                        dataItemLastYear.DSChiQuay = 0;
                        dataItemLastYear.DSChiNha = 0;
                        dataItemLastYear.DSCK = 0;
                        dataItemLastYear.TongDS = 0;
                        dataItemLastYear.Year = (year - 1).ToString();
                        dataItemLastYear.Month = toMonth.ToString();
                        listDataCompareMonth.Add(dataItemLastYear);
                    }

                    // Trường hợp năm hiện tại không có đối tác
                    if (dataItemYear == null)
                    {
                        dataItemYear = new ReportDetailtServiceType();
                        dataItemYear.PartnerName = item.PartnerName;
                        dataItemYear.DSChiQuay = 0;
                        dataItemYear.DSChiNha = 0;
                        dataItemYear.DSCK = 0;
                        dataItemYear.TongDS = 0;
                        dataItemYear.Year = year.ToString();
                        dataItemYear.Month = toMonth.ToString();
                        listDataCompareMonth.Add(dataItemYear);
                    }

                    // Trường hợp tháng trước không có
                    if (dataItemLastMonth == null)
                    {
                        dataItemLastMonth = new ReportDetailtServiceType();
                        dataItemLastMonth.PartnerName = item.PartnerName;
                        dataItemLastMonth.DSChiQuay = 0;
                        dataItemLastMonth.DSChiNha = 0;
                        dataItemLastMonth.DSCK = 0;
                        dataItemLastMonth.TongDS = 0;
                        dataItemLastMonth.Year = year.ToString();
                        dataItemLastMonth.Month = (toMonth - 1).ToString();
                        listDataCompareMonth.Add(dataItemLastMonth);
                    }


                    // Check tồn tại của item
                    string value = string.Format("PartnerName='{0}'", item.PartnerName);
                    DataRow[] foundRows = table.Select(value);

                    if (dataItemLastYear != null && dataItemYear != null && dataItemLastMonth != null && foundRows.Count() == 0)
                    {
                        // Tháng hiện tại
                        double valueYear = sumTongDSYear == 0 ? 0 : Math.Round((dataItemYear.TongDS / sumTongDSYear) * 100, 2, MidpointRounding.ToEven);
                        double valueLastYear = sumTongDSLastYear == 0 ? 0 : Math.Round((dataItemLastYear.TongDS / sumTongDSLastYear) * 100, 2, MidpointRounding.ToEven);
                        double valueLastMonth = sumTongDSLastMonth == 0 ? 0 : Math.Round((dataItemLastMonth.TongDS / sumTongDSLastMonth) * 100, 2, MidpointRounding.ToEven);

                        table.Rows.Add(count, dataItemYear.PartnerName, valueYear, valueLastMonth, valueLastYear);
                    }

                    count++;

                    // Add partnerCode để kiểm tra
                    listTemp.Add(item.PartnerCode);
                }
            }

            DataRow row = table.NewRow();
            row["STT"] = "";
            row["PartnerName"] = "Tổng";
            row["LK1"] = 100;
            row["LK2"] = 100;
            row["LK3"] = 100;
            table.Rows.Add(row);

            return Json(table.ToDataSourceResult(request), JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// Get data cho việc vẽ biểu đồ tròn cho so sánh theo tháng với tháng hiện tại, tháng trước và cùng kì năm trước
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult SearchDataDetailtGridCompareMonthPercent([DataSourceRequest]DataSourceRequest request, int month, int year, string reportTypeID, string marketID)
        {
            List<ReportDetailtServiceType> listDataCompareMonth = new ReportBL().ReportDetailtCompareMonthForOne(year, month, reportTypeID, marketID);

            List<ReportDetailtServiceType> listDataCommpareMonthClone = new List<ReportDetailtServiceType>(listDataCompareMonth);

            foreach (ReportDetailtServiceType item in listDataCompareMonth)
            {
                item.TongDS = item.DSChiQuay + item.DSChiNha + item.DSCK;
            }

            double sumTongDSYear = listDataCompareMonth.Where(x => x.Year == year.ToString() && x.Month == month.ToString()).Sum(x => x.TongDS);
            double sumTongDSLastYear = listDataCompareMonth.Where(x => x.Year == (year - 1).ToString() && x.Month == month.ToString()).Sum(x => x.TongDS);
            double sumTongDSLastMonth = listDataCompareMonth.Where(x => x.Year == year.ToString() && x.Month == (month - 1).ToString()).Sum(x => x.TongDS);

            List<ReportDetailtServiceType> listDataConvert = new List<ReportDetailtServiceType>();

            // Khởi tạo datatable
            DataTable table = new DataTable();
            // Tạo các cột cho datatable
            table.Columns.Add("STT", typeof(String));
            table.Columns.Add("PartnerName", typeof(String));
            table.Columns.Add("LK1", typeof(double));
            table.Columns.Add("LK2", typeof(double));
            table.Columns.Add("LK3", typeof(double));

            List<string> listTemp = new List<string>();

            int count = 1;
            foreach (ReportDetailtServiceType item in listDataCommpareMonthClone)
            {
                if (!listTemp.Contains(item.PartnerCode))
                {
                    // Cùng kì
                    ReportDetailtServiceType dataItemLastYear = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == month.ToString() && x.Year == (year - 1).ToString());
                    ReportDetailtServiceType dataItemYear = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == month.ToString() && x.Year == year.ToString());
                    ReportDetailtServiceType dataItemLastMonth = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == (month - 1).ToString() && x.Year == year.ToString());

                    // Trường hợp năm không có đối tác
                    if (dataItemLastYear == null)
                    {
                        dataItemLastYear = new ReportDetailtServiceType();
                        dataItemLastYear.PartnerName = item.PartnerName;
                        dataItemLastYear.DSChiQuay = 0;
                        dataItemLastYear.DSChiNha = 0;
                        dataItemLastYear.DSCK = 0;
                        dataItemLastYear.TongDS = 0;
                        dataItemLastYear.Year = (year - 1).ToString();
                        dataItemLastYear.Month = month.ToString();
                        listDataCompareMonth.Add(dataItemLastYear);
                    }

                    // Trường hợp năm hiện tại không có đối tác
                    if (dataItemYear == null)
                    {
                        dataItemYear = new ReportDetailtServiceType();
                        dataItemYear.PartnerName = item.PartnerName;
                        dataItemYear.DSChiQuay = 0;
                        dataItemYear.DSChiNha = 0;
                        dataItemYear.DSCK = 0;
                        dataItemYear.TongDS = 0;
                        dataItemYear.Year = year.ToString();
                        dataItemYear.Month = month.ToString();
                        listDataCompareMonth.Add(dataItemYear);
                    }

                    // Trường hợp tháng trước không có
                    if (dataItemLastMonth == null)
                    {
                        dataItemLastMonth = new ReportDetailtServiceType();
                        dataItemLastMonth.PartnerName = item.PartnerName;
                        dataItemLastMonth.DSChiQuay = 0;
                        dataItemLastMonth.DSChiNha = 0;
                        dataItemLastMonth.DSCK = 0;
                        dataItemLastMonth.TongDS = 0;
                        dataItemLastMonth.Year = year.ToString();
                        dataItemLastMonth.Month = (month - 1).ToString();
                        listDataCompareMonth.Add(dataItemLastMonth);
                    }


                    // Check tồn tại của item
                    string value = string.Format("PartnerName='{0}'", item.PartnerName);
                    DataRow[] foundRows = table.Select(value);

                    if (dataItemLastYear != null && dataItemYear != null && dataItemLastMonth != null && foundRows.Count() == 0)
                    {
                        // Tháng hiện tại
                        double valueYear = sumTongDSYear == 0 ? 0 : Math.Round((dataItemYear.TongDS / sumTongDSYear) * 100, 2, MidpointRounding.ToEven);
                        double valueLastYear = sumTongDSLastYear == 0 ? 0 : Math.Round((dataItemLastYear.TongDS / sumTongDSLastYear) * 100, 2, MidpointRounding.ToEven);
                        double valueLastMonth = sumTongDSLastMonth == 0 ? 0 : Math.Round((dataItemLastMonth.TongDS / sumTongDSLastMonth) * 100, 2, MidpointRounding.ToEven);

                        table.Rows.Add(count, dataItemYear.PartnerName, valueYear, valueLastMonth, valueLastYear);
                    }

                    count++;

                    // Add partnerCode để kiểm tra
                    listTemp.Add(item.PartnerCode);
                }
            }

            DataRow row = table.NewRow();
            row["STT"] = "";
            row["PartnerName"] = "Tổng";
            row["LK1"] = 100;
            row["LK2"] = 100;
            row["LK3"] = 100;
            table.Rows.Add(row);

            return Json(table.ToDataSourceResult(request), JsonRequestBehavior.AllowGet);
        }
    }
}
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
        public ActionResult MarketForTotal(DateTime? fromDay, DateTime? toDay, string reportTypeID, string marketID)
        {
            string nameUrl = "Báo cáo chi tiết/Theo doanh số chi trả/Theo thị trường/Loại hình dịch vụ/Tất cả";
            if (marketID.Equals("1"))
            {
                nameUrl = "Báo cáo chi tiết/Theo doanh số chi trả/Theo thị trường/Loại hình dịch vụ/Thị trường Châu Á";
            }
            ViewBag.NameURL = nameUrl;

            if (fromDay != null && toDay != null && reportTypeID != null && marketID != null)
            {
                List<string> listData = new List<string>()
                {
                    fromDay.Value.ToString("MM/dd/yyyy"),
                    toDay.Value.ToString("MM/dd/yyyy"),
                    reportTypeID,
                    marketID
                };

                ViewData["listData"] = listData;
            }
            return View();
        }

        /// <summary>
        /// Màn hình báo cáo cho tháng
        /// </summary>
        /// <returns></returns>
        public ActionResult MarketForTotalForMonth(DateTime? fromDate, DateTime? toDate, string reportTypeID, string marketID)
        {
            string nameUrl = "Báo cáo chi tiết/Theo doanh số chi trả/Theo thị trường/Loại hình dịch vụ/Tất cả";
            if (marketID.Equals("1"))
            {
                nameUrl = "Báo cáo chi tiết/Theo doanh số chi trả/Theo thị trường/Loại hình dịch vụ/Thị trường Châu Á";
            }
            ViewBag.NameURL = nameUrl;

            if (fromDate != null && toDate != null && reportTypeID != null && marketID != null)
            {
                List<string> listData = new List<string>()
                {
                    fromDate.Value.ToString("MM/dd/yyyy"),
                    toDate.Value.ToString("MM/dd/yyyy"),
                    reportTypeID,
                    marketID
                };

                ViewData["listData"] = listData;
            }
            return View();
        }

        /// <summary>
        /// Màn hình báo cáo cho tháng
        /// </summary>
        /// <returns></returns>
        public ActionResult MarketForTotalForYear(DateTime? fromDate, DateTime? toDate, string reportTypeID, string marketID)
        {
            string nameUrl = "Báo cáo chi tiết/Theo doanh số chi trả/Theo thị trường/Loại hình dịch vụ/Tất cả";
            if (marketID.Equals("1"))
            {
                nameUrl = "Báo cáo chi tiết/Theo doanh số chi trả/Theo thị trường/Loại hình dịch vụ/Thị trường Châu Á";
            }
            ViewBag.NameURL = nameUrl;

            if (fromDate != null && toDate != null && reportTypeID != null && marketID != null)
            {
                List<string> listData = new List<string>()
                {
                    fromDate.Value.ToString("MM/dd/yyyy"),
                    toDate.Value.ToString("MM/dd/yyyy"),
                    reportTypeID,
                    marketID
                };

                ViewData["listData"] = listData;
            }
            return View();
        }

        /// <summary>
        /// Màn hình báo cáo chi tiết theo ngày
        /// </summary>
        /// <returns></returns>
        public ActionResult MarketForOne(DateTime? fromDay, DateTime? toDay, string reportTypeID, string marketID)
        {
            string nameUrl = "Báo cáo chi tiết/Theo doanh số chi trả/Theo thị trường/Loại hình dịch vụ/Từng thị trường";
            ViewBag.NameURL = nameUrl;

            if (fromDay != null && toDay != null && reportTypeID != null && marketID != null)
            {
                List<string> listData = new List<string>()
                {
                    fromDay.Value.ToString("MM/dd/yyyy"),
                    toDay.Value.ToString("MM/dd/yyyy"),
                    reportTypeID,
                    marketID
                };

                ViewData["listData"] = listData;
            }

            return View();
        }

        /// <summary>
        /// Màn hình báo cáo chi tiết theo ngày
        /// </summary>
        /// <returns></returns>
        public ActionResult MarketForOneForMonth(DateTime? fromDate, DateTime? toDate, string reportTypeID, string marketID)
        {
            string nameUrl = "Báo cáo chi tiết/Theo doanh số chi trả/Theo thị trường/Loại hình dịch vụ/Từng thị trường";
            ViewBag.NameURL = nameUrl;

            if (fromDate != null && toDate != null && reportTypeID != null && marketID != null)
            {
                List<string> listData = new List<string>()
                {
                    fromDate.Value.ToString("MM/dd/yyyy"),
                    toDate.Value.ToString("MM/dd/yyyy"),
                    reportTypeID,
                    marketID
                };

                ViewData["listData"] = listData;
            }

            return View();
        }

        /// <summary>
        /// Màn hình báo cáo chi tiết theo năm
        /// </summary>
        /// <returns></returns>
        public ActionResult MarketForOneForYear(DateTime? fromDate, DateTime? toDate, string reportTypeID, string marketID)
        {
            string nameUrl = "Báo cáo chi tiết/Theo doanh số chi trả/Theo thị trường/Loại hình dịch vụ/Từng thị trường";
            ViewBag.NameURL = nameUrl;

            if (fromDate != null && toDate != null && reportTypeID != null && marketID != null)
            {
                List<string> listData = new List<string>()
                {
                    fromDate.Value.ToString("MM/dd/yyyy"),
                    toDate.Value.ToString("MM/dd/yyyy"),
                    reportTypeID,
                    marketID
                };

                ViewData["listData"] = listData;
            }

            return View();
        }

        public ActionResult ReportDetailtGradationCompare(string gradation, int? year, string reportTypeID, string marketID)
        {
            string nameUrl = "Báo cáo chi tiết/Theo doanh số chi trả/Theo thị trường/Loại hình dịch vụ/So Sánh/Giai đoạn - Tất cả";
            if (marketID.Equals("1"))
            {
                nameUrl = "Báo cáo chi tiết/Theo doanh số chi trả/Theo thị trường/Loại hình dịch vụ/So Sánh/Giai đoạn - Thị trường Châu Á";
            }
            ViewBag.NameURL = nameUrl;
            // Khởi tạo datatable
            DataTable table = new DataTable();
            // Tạo các cột cho datatable
            table.Columns.Add("PartnerName", typeof(String));

            table.Columns.Add("CQ1", typeof(double));
            table.Columns.Add("CQ2", typeof(double));

            table.Columns.Add("CN1", typeof(double));
            table.Columns.Add("CN2", typeof(double));

            table.Columns.Add("CK1", typeof(double));
            table.Columns.Add("CK2", typeof(double));

            table.Columns.Add("TDS1", typeof(double));
            table.Columns.Add("TDS2", typeof(double));

            table.Columns.Add("MarketName", typeof(String));

            if (!string.IsNullOrEmpty(gradation) && !string.IsNullOrEmpty(marketID))
            {
                if (int.Parse(gradation) > 0 && year > 0 && reportTypeID != null)
                {
                    List<string> listData = new List<string>()
                {
                    gradation,
                    year.ToString(),
                    reportTypeID,
                    marketID
                };

                    ViewData["listData"] = listData;
                }
            }
            return View(table);
        }

        public ActionResult ReportDetailtGradationCompareForOne(string gradation, int? year, string reportTypeID, string marketID)
        {
            string nameUrl = "Báo cáo chi tiết/Theo doanh số chi trả/Theo thị trường/Loại hình dịch vụ/So sánh/Giai đoạn - Từng thị trường";
            ViewBag.NameURL = nameUrl;
            // Khởi tạo datatable
            DataTable table = new DataTable();
            // Tạo các cột cho datatable
            table.Columns.Add("PartnerName", typeof(String));

            table.Columns.Add("CQ1", typeof(double));
            table.Columns.Add("CQ2", typeof(double));

            table.Columns.Add("CN1", typeof(double));
            table.Columns.Add("CN2", typeof(double));

            table.Columns.Add("CK1", typeof(double));
            table.Columns.Add("CK2", typeof(double));

            table.Columns.Add("TDS1", typeof(double));
            table.Columns.Add("TDS2", typeof(double));

            table.Columns.Add("MarketName", typeof(String));

            if (!string.IsNullOrEmpty(gradation) && !string.IsNullOrEmpty(marketID))
            {
                if (int.Parse(gradation) > 0 && year > 0 && reportTypeID != null)
                {
                    List<string> listData = new List<string>()
                {
                    gradation,
                    year.ToString(),
                    reportTypeID,
                    marketID
                };

                    ViewData["listData"] = listData;
                }
            }

            return View(table);
        }

        public ActionResult ReportDetailtCompareMonthForAll(int? month, int? year, string reportTypeID, string marketID)
        {
            string nameUrl = "Báo cáo chi tiết/Theo doanh số chi trả/Theo thị trường/Loại hình dịch vụ/So Sánh/Tháng - Tất cả";
            if (marketID.Equals("1"))
            {
                nameUrl = "Báo cáo chi tiết/Theo doanh số chi trả/Theo thị trường/Loại hình dịch vụ/So Sánh/Tháng - Thị trường Châu Á";
            }
            ViewBag.NameURL = nameUrl;
            // Khởi tạo datatable
            DataTable table = new DataTable();
            // Tạo các cột cho datatable
            table.Columns.Add("MarketName", typeof(String));
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

            table.Columns.Add("ReportID", typeof(String));

            if (month != null && year != null)
            {
                if (month.Value > 0 && year.Value > 0 && reportTypeID != null && marketID != null)
                {
                    List<string> listData = new List<string>()
                    {
                        month.ToString(),
                        year.ToString(),
                        reportTypeID,
                        marketID
                    };

                    ViewData["listData"] = listData;
                }
            }

            return View(table);
        }

        public ActionResult ReportDetailtCompareMonthForOne(int? month, int? year, string reportTypeID, string marketID)
        {
            string nameUrl = "Báo cáo chi tiết/Theo doanh số chi trả/Theo thị trường/Loại hình dịch vụ/So sánh/Tháng - Từng thị trường";
            ViewBag.NameURL = nameUrl;
            // Khởi tạo datatable
            DataTable table = new DataTable();
            // Tạo các cột cho datatable
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


            table.Columns.Add("MarketName", typeof(String));

            if (month != null && year != null)
            {
                if (month.Value > 0 && year.Value > 0 && reportTypeID != null && marketID != null)
                {
                    List<string> listData = new List<string>()
                    {
                        month.ToString(),
                        year.ToString(),
                        reportTypeID,
                        marketID
                    };

                    ViewData["listData"] = listData;
                }
            }

            return View(table);
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
            List<ReportDetailtSTMarket> listData = new ReportBL().SearchMarketForTotalForDay(fromDay, toDay, reportTypeID, marketID);

            if (!string.IsNullOrEmpty(marketID))
            {
                if (marketID == "0")
                {
                    foreach (ReportDetailtSTMarket item in listData)
                    {
                        item.ReportID = item.MarketName;
                        item.TongDS = item.DSChiNha + item.DSChiQuay + item.DSCK;
                        item.MarketName = "Tất cả";
                    }
                }
                else
                {
                    List<string> listAsianChild = new List<string>();
                    List<ReportDetailtSTMarket> listDataConvert = new List<ReportDetailtSTMarket>();

                    foreach (ReportDetailtSTMarket item in listData)
                    {
                        item.ReportID = item.PartnerName;
                        item.TongDS = item.DSChiNha + item.DSChiQuay + item.DSCK;

                        if(!listAsianChild.Contains(item.MarketName))
                        {
                            listAsianChild.Add(item.MarketName);
                        }
                    }

                    foreach(string item in listAsianChild)
                    {
                        List<ReportDetailtSTMarket> listDataAsianChild = listData.Where(x => x.MarketName == item).ToList();

                        listDataConvert.Add(
                            new ReportDetailtSTMarket()
                            {
                                MarketName = "Châu Á",
                                ReportID = item,
                                DSChiQuay = listDataAsianChild.Sum(x => x.DSChiQuay),
                                DSChiNha = listDataAsianChild.Sum(x => x.DSChiNha),
                                DSCK = listDataAsianChild.Sum(x => x.DSCK),
                                TongDS = listDataAsianChild.Sum(x => x.TongDS),

                            }
                        );
                    }

                    if(listDataConvert.Count > 0)
                    {
                        listData = new List<ReportDetailtSTMarket>(listDataConvert);
                    }
                }
            }

            return Json(listData.ToDataSourceResult(request), JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// Search report day theo ngày nhập vào
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public ActionResult SearchMarketForTotalForMonth([DataSourceRequest]DataSourceRequest request, DateTime fromDate, DateTime toDate, string reportTypeID, string marketID)
        {
            List<ReportDetailtSTMarket> listData = new ReportBL().SearchMarketForTotalForMonth(fromDate, toDate, reportTypeID, marketID);

            if (!string.IsNullOrEmpty(marketID))
            {
                if (marketID == "0")
                {
                    foreach (ReportDetailtSTMarket item in listData)
                    {
                        item.ReportID = item.MarketName;
                        item.TongDS = item.DSChiNha + item.DSChiQuay + item.DSCK;
                        item.MarketName = "Tất cả";
                    }
                }
                else
                {
                    List<string> listAsianChild = new List<string>();
                    List<ReportDetailtSTMarket> listDataConvert = new List<ReportDetailtSTMarket>();

                    foreach (ReportDetailtSTMarket item in listData)
                    {
                        item.ReportID = item.PartnerName;
                        item.TongDS = item.DSChiNha + item.DSChiQuay + item.DSCK;

                        if (!listAsianChild.Contains(item.MarketName))
                        {
                            listAsianChild.Add(item.MarketName);
                        }
                    }

                    foreach (string item in listAsianChild)
                    {
                        List<ReportDetailtSTMarket> listDataAsianChild = listData.Where(x => x.MarketName == item).ToList();

                        listDataConvert.Add(
                            new ReportDetailtSTMarket()
                            {
                                MarketName = "Châu Á",
                                ReportID = item,
                                DSChiQuay = listDataAsianChild.Sum(x => x.DSChiQuay),
                                DSChiNha = listDataAsianChild.Sum(x => x.DSChiNha),
                                DSCK = listDataAsianChild.Sum(x => x.DSCK),
                                TongDS = listDataAsianChild.Sum(x => x.TongDS),

                            }
                        );
                    }

                    if (listDataConvert.Count > 0)
                    {
                        listData = new List<ReportDetailtSTMarket>(listDataConvert);
                    }
                }
            }

            return Json(listData.ToDataSourceResult(request), JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// Search report day theo ngày nhập vào
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public ActionResult SearchMarketForTotalForYear([DataSourceRequest]DataSourceRequest request, DateTime fromDate, DateTime toDate, string reportTypeID, string marketID)
        {
            List<ReportDetailtSTMarket> listData = new ReportBL().SearchMarketForTotalForYear(fromDate, toDate, reportTypeID, marketID);

            if (!string.IsNullOrEmpty(marketID))
            {
                if (marketID == "0")
                {
                    foreach (ReportDetailtSTMarket item in listData)
                    {
                        item.ReportID = item.MarketName;
                        item.TongDS = item.DSChiNha + item.DSChiQuay + item.DSCK;
                        item.MarketName = "Tất cả";
                    }
                }
                else
                {
                    List<string> listAsianChild = new List<string>();
                    List<ReportDetailtSTMarket> listDataConvert = new List<ReportDetailtSTMarket>();

                    foreach (ReportDetailtSTMarket item in listData)
                    {
                        item.ReportID = item.PartnerName;
                        item.TongDS = item.DSChiNha + item.DSChiQuay + item.DSCK;

                        if (!listAsianChild.Contains(item.MarketName))
                        {
                            listAsianChild.Add(item.MarketName);
                        }
                    }

                    foreach (string item in listAsianChild)
                    {
                        List<ReportDetailtSTMarket> listDataAsianChild = listData.Where(x => x.MarketName == item).ToList();

                        listDataConvert.Add(
                            new ReportDetailtSTMarket()
                            {
                                MarketName = "Châu Á",
                                ReportID = item,
                                DSChiQuay = listDataAsianChild.Sum(x => x.DSChiQuay),
                                DSChiNha = listDataAsianChild.Sum(x => x.DSChiNha),
                                DSCK = listDataAsianChild.Sum(x => x.DSCK),
                                TongDS = listDataAsianChild.Sum(x => x.TongDS),

                            }
                        );
                    }

                    if (listDataConvert.Count > 0)
                    {
                        listData = new List<ReportDetailtSTMarket>(listDataConvert);
                    }
                }
            }
            return Json(listData.ToDataSourceResult(request), JsonRequestBehavior.AllowGet);
        }

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
        public JsonResult ListMarketForPartner(string text)
        {
            List<Market> list = new ReportBL().ListMarket();
            List<Market> listMarket = new List<Market>();
            // Get danh sách các thị trường chính
            if (list.Count > 0)
            {
                listMarket = list.Where(x => x.ParentCode == null).ToList();

                if (!string.IsNullOrEmpty(text))
                {
                    listMarket = list.Where(x => x.ParentCode == null && x.MarketName.Contains(text)).ToList();
                }
            }
            return Json(listMarket, JsonRequestBehavior.AllowGet);
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
            List<string> listDataMarket = new List<string>();

            string marketName = string.Empty;

            foreach (ReportDetailtServiceType item in listData)
            {

                // Trường hợp không phải là châu Á
                if (!listDataMarket.Contains(item.MarketName))
                {
                    listDataMarket.Add(item.MarketName);
                    count = 1;
                    item.STT = (count++).ToString();
                }
                else
                {
                    item.STT = (count++).ToString();
                }

                item.TongDS = item.DSChiNha + item.DSChiQuay + item.DSCK;
            }

            return Json(listData.ToDataSourceResult(request), JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// Search report day theo ngày nhập vào
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public ActionResult SearchMarketForOneForMonth([DataSourceRequest]DataSourceRequest request, DateTime fromDate, DateTime toDate, string reportTypeID, string marketID)
        {
            List<ReportDetailtServiceType> listData = new ReportBL().SearchMarketForOneForMonth(fromDate, toDate, reportTypeID, marketID);
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
        /// Search report year cho báo cáo chi tiết
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public ActionResult SearchMarketForOneForYear([DataSourceRequest]DataSourceRequest request, DateTime fromDate, DateTime toDate, string reportTypeID, string marketID)
        {
            List<ReportDetailtServiceType> listData = new ReportBL().SearchMarketForOneForYear(fromDate, toDate, reportTypeID, marketID);
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
        /// Get data cho việc vẽ biểu đồ cột cho so sánh theo giai đoạn của Chi Quầy
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult SearchColumnChartReportForGradationForDSChiQuay([DataSourceRequest]DataSourceRequest request, int gradation, int year, string reportTypeID, string marketID)
        {
            List<ReportDetailtServiceType> listDataGradation = new ReportBL().ReportDetailtGradationCompareForAll(year, gradation, reportTypeID, marketID);
            GradationCompare[] arrayGradation = null;
            // Trường hợp tất cả thị trường
            if (marketID.Equals("0"))
            {
                // Số record của mảng
                int countArray = 1;
                arrayGradation = new GradationCompare[countArray * listDataGradation.Count];
                int count = 0;
                foreach (ReportDetailtServiceType item in listDataGradation)
                {
                    // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                    arrayGradation[count] = new GradationCompare()
                    {
                        NameGradationCompare = string.Format("Chi Quầy {0}", item.Year),
                        amount = item.DSChiQuay,
                        NameType = item.MarketName
                    };

                    count++;
                    year = year - 1;
                }
            }
            else
            {
                // List các thị trường của châu Á
                List<string> listAsianItem = new List<string>();
                foreach (ReportDetailtServiceType item in listDataGradation)
                {
                    if (!listAsianItem.Contains(item.MarketName))
                    {
                        listAsianItem.Add(item.MarketName);
                    }
                }

                // Số record của mảng
                int countArray = 2;
                arrayGradation = new GradationCompare[countArray * listAsianItem.Count];
                int count = 0;
                List<ReportDetailtServiceType> listDSMarketYear = new List<ReportDetailtServiceType>();
                List<ReportDetailtServiceType> listDSMarketLastYear = new List<ReportDetailtServiceType>();

                foreach (string item in listAsianItem)
                {
                    listDSMarketYear = listDataGradation.Where(x => x.MarketName == item && x.Year == year.ToString()).ToList();
                    listDSMarketLastYear = listDataGradation.Where(x => x.MarketName == item && x.Year == (year - 1).ToString()).ToList();

                    // Year
                    double sumDSChiQuayYear = listDSMarketYear.Sum(x => x.DSChiQuay);
                    double sumDSChiNhaYear = listDSMarketYear.Sum(x => x.DSChiNha);
                    double sumDSCKYear = listDSMarketYear.Sum(x => x.DSCK);
                    double sumTongDSYear = listDSMarketYear.Sum(x => x.TongDS);

                    // LastYear
                    double sumDSChiQuayLastYear = listDSMarketLastYear.Sum(x => x.DSChiQuay);
                    double sumDSChiNhaLastYear = listDSMarketLastYear.Sum(x => x.DSChiNha);
                    double sumDSCKLastYear = listDSMarketLastYear.Sum(x => x.DSCK);
                    double sumTongDSLastYear = listDSMarketLastYear.Sum(x => x.TongDS);

                    // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                    arrayGradation[count] = new GradationCompare()
                    {
                        NameGradationCompare = string.Format("Chi Quầy {0}", year),
                        amount = sumDSChiQuayYear,
                        NameType = item
                    };

                    count++;

                    // LastYear
                    // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                    arrayGradation[count] = new GradationCompare()
                    {
                        NameGradationCompare = string.Format("Chi Quầy {0}", year - 1),
                        amount = sumDSChiQuayLastYear,
                        NameType = item
                    };

                    count++;
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
        /// Get data cho việc vẽ biểu đồ cột cho so sánh theo giai đoạn của Chi nhà
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult SearchColumnChartReportForGradationForDSChiNha([DataSourceRequest]DataSourceRequest request, int gradation, int year, string reportTypeID, string marketID)
        {
            List<ReportDetailtServiceType> listDataGradation = new ReportBL().ReportDetailtGradationCompareForAll(year, gradation, reportTypeID, marketID);
            GradationCompare[] arrayGradation = null;
            // Trường hợp tất cả thị trường
            if (marketID.Equals("0"))
            {
                // Số record của mảng
                int countArray = 1;
                arrayGradation = new GradationCompare[countArray * listDataGradation.Count];
                int count = 0;
                foreach (ReportDetailtServiceType item in listDataGradation)
                {
                    // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                    arrayGradation[count] = new GradationCompare()
                    {
                        NameGradationCompare = string.Format("Chi Nhà {0}", item.Year),
                        amount = item.DSChiNha,
                        NameType = item.MarketName
                    };

                    count++;
                    year = year - 1;
                }
            }
            else
            {
                // List các thị trường của châu Á
                List<string> listAsianItem = new List<string>();
                foreach (ReportDetailtServiceType item in listDataGradation)
                {
                    if (!listAsianItem.Contains(item.MarketName))
                    {
                        listAsianItem.Add(item.MarketName);
                    }
                }

                // Số record của mảng
                int countArray = 2;
                arrayGradation = new GradationCompare[countArray * listAsianItem.Count];
                int count = 0;
                List<ReportDetailtServiceType> listDSMarketYear = new List<ReportDetailtServiceType>();
                List<ReportDetailtServiceType> listDSMarketLastYear = new List<ReportDetailtServiceType>();

                foreach (string item in listAsianItem)
                {
                    listDSMarketYear = listDataGradation.Where(x => x.MarketName == item && x.Year == year.ToString()).ToList();
                    listDSMarketLastYear = listDataGradation.Where(x => x.MarketName == item && x.Year == (year - 1).ToString()).ToList();

                    // Year
                    double sumDSChiQuayYear = listDSMarketYear.Sum(x => x.DSChiQuay);
                    double sumDSChiNhaYear = listDSMarketYear.Sum(x => x.DSChiNha);
                    double sumDSCKYear = listDSMarketYear.Sum(x => x.DSCK);
                    double sumTongDSYear = listDSMarketYear.Sum(x => x.TongDS);

                    // LastYear
                    double sumDSChiQuayLastYear = listDSMarketLastYear.Sum(x => x.DSChiQuay);
                    double sumDSChiNhaLastYear = listDSMarketLastYear.Sum(x => x.DSChiNha);
                    double sumDSCKLastYear = listDSMarketLastYear.Sum(x => x.DSCK);
                    double sumTongDSLastYear = listDSMarketLastYear.Sum(x => x.TongDS);

                    // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                    arrayGradation[count] = new GradationCompare()
                    {
                        NameGradationCompare = string.Format("Chi Nhà {0}", year),
                        amount = sumDSChiNhaYear,
                        NameType = item
                    };

                    count++;

                    // LastYear
                    // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                    arrayGradation[count] = new GradationCompare()
                    {
                        NameGradationCompare = string.Format("Chi Nhà {0}", year - 1),
                        amount = sumDSChiNhaLastYear,
                        NameType = item
                    };

                    count++;

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
        /// Get data cho việc vẽ biểu đồ cột cho so sánh theo giai đoạn của Chi nhà
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult SearchColumnChartReportForGradationForDSCK([DataSourceRequest]DataSourceRequest request, int gradation, int year, string reportTypeID, string marketID)
        {
            List<ReportDetailtServiceType> listDataGradation = new ReportBL().ReportDetailtGradationCompareForAll(year, gradation, reportTypeID, marketID);
            GradationCompare[] arrayGradation = null;
            // Trường hợp tất cả thị trường
            if (marketID.Equals("0"))
            {
                // Số record của mảng
                int countArray = 1;
                arrayGradation = new GradationCompare[countArray * listDataGradation.Count];
                int count = 0;
                foreach (ReportDetailtServiceType item in listDataGradation)
                {
                    // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                    arrayGradation[count] = new GradationCompare()
                    {
                        NameGradationCompare = string.Format("Chuyển khoản {0}", item.Year),
                        amount = item.DSCK,
                        NameType = item.MarketName
                    };

                    count++;
                    year = year - 1;
                }
            }
            else
            {
                // List các thị trường của châu Á
                List<string> listAsianItem = new List<string>();
                foreach (ReportDetailtServiceType item in listDataGradation)
                {
                    if (!listAsianItem.Contains(item.MarketName))
                    {
                        listAsianItem.Add(item.MarketName);
                    }
                }

                // Số record của mảng
                int countArray = 2;
                arrayGradation = new GradationCompare[countArray * listAsianItem.Count];
                int count = 0;
                List<ReportDetailtServiceType> listDSMarketYear = new List<ReportDetailtServiceType>();
                List<ReportDetailtServiceType> listDSMarketLastYear = new List<ReportDetailtServiceType>();

                foreach (string item in listAsianItem)
                {
                    listDSMarketYear = listDataGradation.Where(x => x.MarketName == item && x.Year == year.ToString()).ToList();
                    listDSMarketLastYear = listDataGradation.Where(x => x.MarketName == item && x.Year == (year - 1).ToString()).ToList();

                    // Year
                    double sumDSChiQuayYear = listDSMarketYear.Sum(x => x.DSChiQuay);
                    double sumDSChiNhaYear = listDSMarketYear.Sum(x => x.DSChiNha);
                    double sumDSCKYear = listDSMarketYear.Sum(x => x.DSCK);
                    double sumTongDSYear = listDSMarketYear.Sum(x => x.TongDS);

                    // LastYear
                    double sumDSChiQuayLastYear = listDSMarketLastYear.Sum(x => x.DSChiQuay);
                    double sumDSChiNhaLastYear = listDSMarketLastYear.Sum(x => x.DSChiNha);
                    double sumDSCKLastYear = listDSMarketLastYear.Sum(x => x.DSCK);
                    double sumTongDSLastYear = listDSMarketLastYear.Sum(x => x.TongDS);

                    // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                    arrayGradation[count] = new GradationCompare()
                    {
                        NameGradationCompare = string.Format("DSCK {0}", year),
                        amount = sumDSCKYear,
                        NameType = item
                    };

                    count++;
                    //arrayGradation[count] = new GradationCompare()
                    //{
                    //    NameGradationCompare = item,
                    //    amount = sumDSChiNhaYear,
                    //    NameType = string.Format("Chi Nhà {0}", year)
                    //};

                    //count++;
                    //arrayGradation[count] = new GradationCompare()
                    //{
                    //    NameGradationCompare = item,
                    //    amount = sumDSCKYear,
                    //    NameType = string.Format("DSCK {0}", year)
                    //};

                    //count++;

                    // LastYear
                    // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                    arrayGradation[count] = new GradationCompare()
                    {
                        NameGradationCompare = string.Format("DSCK {0}", year - 1),
                        amount = sumDSCKLastYear,
                        NameType = item
                    };

                    count++;
                    //arrayGradation[count] = new GradationCompare()
                    //{
                    //    NameGradationCompare = item,
                    //    amount = sumDSChiNhaLastYear,
                    //    NameType = string.Format("Chi Nhà {0}", year - 1)
                    //};

                    //count++;
                    //arrayGradation[count] = new GradationCompare()
                    //{
                    //    NameGradationCompare = item,
                    //    amount = sumDSCKLastYear,
                    //    NameType = string.Format("DSCK {0}", year - 1)
                    //};

                    //// Tăng count lên 1 đơn vị
                    //count++;
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
        /// Get data cho việc vẽ biểu đồ cột cho so sánh theo giai đoạn
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult SearchColumnChartReportForGradationPercent([DataSourceRequest]DataSourceRequest request, int gradation, int year, string reportTypeID, string marketID)
        {
            List<ReportDetailtServiceType> listDataGradation = new ReportBL().ReportDetailtGradationCompareForAllPercent(year, gradation, reportTypeID, marketID);

            GradationCharColumn[] arrayGradation = null;
            // Trường hợp chọn tất cả thị trường
            if (marketID.Equals("0"))
            {
                // Số mảng cần tạo
                int arrayCount = 3;
                arrayGradation = new GradationCharColumn[listDataGradation.Count * arrayCount];
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

            }
            else
            {
                // Trường hợp thị trường  Châu á
                List<string> listAsianItem = new List<string>();
                foreach (ReportDetailtServiceType item in listDataGradation)
                {
                    if (!listAsianItem.Contains(item.MarketName))
                    {
                        listAsianItem.Add(item.MarketName);
                    }
                }

                // Số record của mảng
                int countArray = 6;
                arrayGradation = new GradationCharColumn[listAsianItem.Count * countArray];
                int count = 0;
                // group theo nhóm
                int tooltip = 1;
                List<ReportDetailtServiceType> listDSMarketYear = new List<ReportDetailtServiceType>();
                List<ReportDetailtServiceType> listDSMarketLastYear = new List<ReportDetailtServiceType>();

                foreach (string item in listAsianItem)
                {
                    listDSMarketYear = listDataGradation.Where(x => x.MarketName == item && x.Year == year.ToString()).ToList();
                    listDSMarketLastYear = listDataGradation.Where(x => x.MarketName == item && x.Year == (year - 1).ToString()).ToList();

                    // Year
                    double sumDSChiQuayYear = listDSMarketYear.Sum(x => x.DSChiQuay);
                    double sumDSChiNhaYear = listDSMarketYear.Sum(x => x.DSChiNha);
                    double sumDSCKYear = listDSMarketYear.Sum(x => x.DSCK);
                    double sumTongDSYear = listDSMarketYear.Sum(x => x.TongDS);

                    // LastYear
                    double sumDSChiQuayLastYear = listDSMarketLastYear.Sum(x => x.DSChiQuay);
                    double sumDSChiNhaLastYear = listDSMarketLastYear.Sum(x => x.DSChiNha);
                    double sumDSCKLastYear = listDSMarketLastYear.Sum(x => x.DSCK);
                    double sumTongDSLastYear = listDSMarketLastYear.Sum(x => x.TongDS);

                    // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                    // Year
                    arrayGradation[count] = new GradationCharColumn()
                    {
                        Serie = item,
                        Segmento = string.Format("Chi Quầy {0}", year),
                        Valor1 = sumDSChiQuayYear,
                        Tooltip = tooltip
                    };

                    count++;

                    arrayGradation[count] = new GradationCharColumn()
                    {
                        Serie = item,
                        Segmento = string.Format("Chi Nhà {0}", year),
                        Valor1 = sumDSChiNhaYear,
                        Tooltip = tooltip
                    };

                    count++;

                    arrayGradation[count] = new GradationCharColumn()
                    {
                        Serie = item,
                        Segmento = string.Format("DSCK {0}", year),
                        Valor1 = sumDSCKYear,
                        Tooltip = tooltip
                    };

                    count++;

                    // LastYear
                    arrayGradation[count] = new GradationCharColumn()
                    {
                        Serie = item,
                        Segmento = string.Format("Chi Quầy {0}", year - 1),
                        Valor1 = sumDSChiQuayLastYear,
                        Tooltip = tooltip
                    };

                    count++;

                    arrayGradation[count] = new GradationCharColumn()
                    {
                        Serie = item,
                        Segmento = string.Format("Chi Nhà {0}", year - 1),
                        Valor1 = sumDSChiNhaLastYear,
                        Tooltip = tooltip
                    };

                    count++;

                    arrayGradation[count] = new GradationCharColumn()
                    {
                        Serie = item,
                        Segmento = string.Format("DSCK {0}", year - 1),
                        Valor1 = sumDSCKLastYear,
                        Tooltip = tooltip
                    };

                    count++;
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
        /// Get data cho việc vẽ biểu đồ cột cho so sánh theo giai đoạn
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult SearchGridReportForGradation([DataSourceRequest]DataSourceRequest request, int gradation, int year, string reportTypeID, string marketID)
        {
            List<ReportDetailtServiceType> listDataGradation = new ReportBL().ReportDetailtGradationCompareForAll(year, gradation, reportTypeID, marketID);

            // Khởi tạo datatable
            DataTable table = new DataTable();

            // Trường hợp chọn tất cả thị trường
            if (marketID.Equals("0"))
            {
                foreach (ReportDetailtServiceType item in listDataGradation)
                {
                    item.ReportID = item.MarketName;
                    item.TongDS = item.DSChiQuay + item.DSChiNha + item.DSCK;
                    item.MarketName = "Tất cả thị trường";
                }

                // Khởi tạo datatable
                table = new DataTable();
                // Tạo các cột cho datatable
                table.Columns.Add("MarketName", typeof(String));
                table.Columns.Add("CQ1", typeof(double));
                table.Columns.Add("CQ2", typeof(double));

                table.Columns.Add("CN1", typeof(double));
                table.Columns.Add("CN2", typeof(double));

                table.Columns.Add("CK1", typeof(double));
                table.Columns.Add("CK2", typeof(double));

                table.Columns.Add("TDS1", typeof(double));
                table.Columns.Add("TDS2", typeof(double));

                table.Columns.Add("ReportID", typeof(String));

                // Danh sách mã thị trường của Tất cả
                List<string> listMarket = new List<string>() { "003", "005", "001", "002", "014", "004" };

                string reportID = string.Empty;

                if (listDataGradation.Count() > 0)
                {
                    foreach (string item in listMarket)
                    {
                        // Cùng kì
                        ReportDetailtServiceType dataItemLastYear = listDataGradation.Find(x => x.MarketCode == item && x.Year == (year - 1).ToString());
                        ReportDetailtServiceType dataItemYear = listDataGradation.Find(x => x.MarketCode == item && x.Year == year.ToString());

                        if (dataItemYear != null && dataItemLastYear != null)
                        {
                            reportID = dataItemYear.ReportID;
                        }

                        // Trường hợp năm không có đối tác
                        if (dataItemLastYear == null && dataItemYear != null)
                        {
                            dataItemLastYear = new ReportDetailtServiceType();
                            dataItemLastYear.MarketName = dataItemYear.MarketName;
                            dataItemLastYear.DSChiQuay = 0;
                            dataItemLastYear.DSChiNha = 0;
                            dataItemLastYear.DSCK = 0;
                            dataItemLastYear.Year = (year - 1).ToString();

                            reportID = dataItemYear.ReportID;
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

                            reportID = dataItemLastYear.ReportID;
                        }

                        // add item vào table
                        table.Rows.Add(dataItemLastYear.MarketName
                            , dataItemLastYear.DSChiQuay, dataItemYear.DSChiQuay
                            , dataItemLastYear.DSChiNha, dataItemYear.DSChiNha
                            , dataItemLastYear.DSCK, dataItemYear.DSCK
                            , dataItemLastYear.TongDS, dataItemYear.TongDS
                            , reportID);
                    }
                }
            }
            // Trường hợp thuộc thị trường Châu Á
            else
            {
                foreach (ReportDetailtServiceType item in listDataGradation)
                {
                    item.ReportID = item.PartnerName;
                    item.TongDS = item.DSChiQuay + item.DSChiNha + item.DSCK;
                }

                // Khởi tạo datatable
                table = new DataTable();
                // Tạo các cột cho datatable
                table.Columns.Add("MarketName", typeof(String));

                table.Columns.Add("CQ1", typeof(double));
                table.Columns.Add("CQ2", typeof(double));

                table.Columns.Add("CN1", typeof(double));
                table.Columns.Add("CN2", typeof(double));

                table.Columns.Add("CK1", typeof(double));
                table.Columns.Add("CK2", typeof(double));

                table.Columns.Add("TDS1", typeof(double));
                table.Columns.Add("TDS2", typeof(double)); ;

                table.Columns.Add("ReportID", typeof(String));

                string reportID = string.Empty;

                foreach (ReportDetailtServiceType item in listDataGradation)
                {
                    // Cùng kì
                    ReportDetailtServiceType dataItemLastYear = listDataGradation.Find(x => x.PartnerCode == item.PartnerCode && x.Year == (year - 1).ToString());
                    ReportDetailtServiceType dataItemYear = listDataGradation.Find(x => x.PartnerCode == item.PartnerCode && x.Year == year.ToString());

                    reportID = item.PartnerName;

                    // Trường hợp năm trước có đối tác và năm nay không có
                    if (dataItemLastYear != null && dataItemYear == null)
                    {
                        dataItemYear = new ReportDetailtServiceType();
                        dataItemYear.PartnerName = dataItemLastYear.PartnerName;
                        dataItemYear.ReportID = dataItemLastYear.ReportID;
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
                        dataItemLastYear.ReportID = dataItemYear.ReportID;
                        dataItemLastYear.DSChiQuay = 0;
                        dataItemLastYear.DSChiNha = 0;
                        dataItemLastYear.DSCK = 0;
                        dataItemLastYear.Year = (year - 1).ToString();
                    }

                    // Check tồn tại của item
                    string value = string.Format("ReportID='{0}'", reportID);
                    DataRow[] foundRows = table.Select(value);
                    if (dataItemLastYear != null && dataItemYear != null && foundRows.Count() == 0)
                    {
                        // add item vào table
                        table.Rows.Add(item.MarketName
                            , dataItemLastYear.DSChiQuay, dataItemYear.DSChiQuay
                            , dataItemLastYear.DSChiNha, dataItemYear.DSChiNha
                            , dataItemLastYear.DSCK, dataItemYear.DSCK
                            , dataItemLastYear.TongDS, dataItemYear.TongDS
                            , reportID);
                    }
                }
            }

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
        public ActionResult SearchGridReportForGradationForOne([DataSourceRequest]DataSourceRequest request, int gradation, int year, string reportTypeID, string marketID)
        {
            List<ReportDetailtServiceType> listDataGradation = new ReportBL().ReportDetailtGradationCompareForOne(year, gradation, reportTypeID, marketID);
            List<ReportDetailtServiceType> listDataGradationConvert = new List<ReportDetailtServiceType>();

            // List thị trường
            List<string> listMarket = new List<string>();
            if (marketID.Equals("005"))
            {
                foreach(ReportDetailtServiceType item in listDataGradation)
                {
                    if (!listMarket.Contains(item.MarketName))
                    {
                        listMarket.Add(item.MarketName);
                    }
                }

                foreach(string item in listMarket)
                {
                    List<ReportDetailtServiceType> listDataItemYear = listDataGradation.Where(x => x.MarketName == item && x.Year == year.ToString()).ToList();
                    List<ReportDetailtServiceType> listDataItemLastYear = listDataGradation.Where(x => x.MarketName == item && x.Year == (year - 1).ToString()).ToList();

                    // Year
                    listDataGradationConvert.Add(
                        new ReportDetailtServiceType()
                        {
                            MarketName = item,
                            DSChiQuay = listDataItemYear.Sum(x => x.DSChiQuay),
                            DSChiNha = listDataItemYear.Sum(x => x.DSChiNha),
                            DSCK = listDataItemYear.Sum(x => x.DSCK),
                            Year = year.ToString()
                        }
                    );

                    // last year
                    listDataGradationConvert.Add(
                        new ReportDetailtServiceType()
                        {
                            MarketName = item,
                            DSChiQuay = listDataItemLastYear.Sum(x => x.DSChiQuay),
                            DSChiNha = listDataItemLastYear.Sum(x => x.DSChiNha),
                            DSCK = listDataItemLastYear.Sum(x => x.DSCK),
                            Year = (year - 1).ToString()
                        }
                    );
                }

                if (listDataGradationConvert.Count > 0)
                {
                    listDataGradation = new List<ReportDetailtServiceType>(listDataGradationConvert);
                }
            }

            // Khởi tạo datatable
            DataTable table = new DataTable();
            // Tạo các cột cho datatable
            table.Columns.Add("PartnerName", typeof(String));

            table.Columns.Add("CQ1", typeof(double));
            table.Columns.Add("CQ2", typeof(double));

            table.Columns.Add("CN1", typeof(double));
            table.Columns.Add("CN2", typeof(double));

            table.Columns.Add("CK1", typeof(double));
            table.Columns.Add("CK2", typeof(double));

            table.Columns.Add("TDS1", typeof(double));
            table.Columns.Add("TDS2", typeof(double));

            table.Columns.Add("MarketName", typeof(String));

            if (marketID.Equals("005"))
            {
                foreach (ReportDetailtServiceType item in listDataGradation)
                {
                    item.PartnerName = item.MarketName;
                    item.TongDS = item.DSChiQuay + item.DSChiNha + item.DSCK;
                    item.MarketName = "Thị trường Châu Á";
                }
            }
            else
            {
                foreach (ReportDetailtServiceType item in listDataGradation)
                {
                    item.MarketName = "Tất cả thị trường";
                    item.TongDS = item.DSChiQuay + item.DSChiNha + item.DSCK;
                }
            }

            foreach (ReportDetailtServiceType item in listDataGradation)
            {
                // Cùng kì
                ReportDetailtServiceType dataItemLastYear = listDataGradation.Find(x => x.PartnerName == item.PartnerName && x.Year == (year - 1).ToString());
                ReportDetailtServiceType dataItemYear = listDataGradation.Find(x => x.PartnerName == item.PartnerName && x.Year == year.ToString());

                // Trường hợp năm trước có đối tác và năm nay không có
                if (dataItemLastYear != null && dataItemYear == null)
                {
                    dataItemYear = new ReportDetailtServiceType();
                    dataItemYear.PartnerName = dataItemLastYear.PartnerName;
                    dataItemYear.MarketName = dataItemLastYear.MarketName;
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
                    dataItemLastYear.MarketName = dataItemYear.MarketName;
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
                    table.Rows.Add(dataItemLastYear.PartnerName
                        , dataItemLastYear.DSChiQuay, dataItemYear.DSChiQuay
                        , dataItemLastYear.DSChiNha, dataItemYear.DSChiNha
                        , dataItemLastYear.DSCK, dataItemYear.DSCK
                        , dataItemLastYear.TongDS, dataItemYear.TongDS
                        , dataItemLastYear.MarketName);
                }
            }

            DataRow row = table.NewRow();
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
        public ActionResult SearchColumnsChartGradationCompareForOneForDSChiQuay([DataSourceRequest]DataSourceRequest request, int gradation, int year, string reportTypeID, string marketID)
        {
            List<ReportDetailtServiceType> listDataGradation = new ReportBL().ReportDetailtGradationCompareForOne(year, gradation, reportTypeID, marketID);
            if (marketID.Contains("005"))
            {
                List<ReportDetailtServiceType> listDataGradationConvert = new List<ReportDetailtServiceType>();
                List<string> listMarket = new List<string>();

                foreach (ReportDetailtServiceType item in listDataGradation)
                {
                    if (!listMarket.Contains(item.MarketName))
                    {
                        listMarket.Add(item.MarketName);
                    }
                }

                foreach (string item in listMarket)
                {
                    List<ReportDetailtServiceType> listDataItemYear = listDataGradation.Where(x => x.MarketName == item && x.Year == year.ToString()).ToList();
                    List<ReportDetailtServiceType> listDataItemLastYear = listDataGradation.Where(x => x.MarketName == item && x.Year == (year - 1).ToString()).ToList();

                    // Year
                    listDataGradationConvert.Add(
                        new ReportDetailtServiceType()
                        {
                            MarketName = item,
                            PartnerName = item,
                            DSChiQuay = listDataItemYear.Sum(x => x.DSChiQuay),
                            DSChiNha = listDataItemYear.Sum(x => x.DSChiNha),
                            DSCK = listDataItemYear.Sum(x => x.DSCK),
                            TongDS = listDataItemYear.Sum(x => x.TongDS),
                            Year = year.ToString()
                        }
                    );

                    // Last Year
                    listDataGradationConvert.Add(
                        new ReportDetailtServiceType()
                        {
                            MarketName = item,
                            PartnerName = item,
                            DSChiQuay = listDataItemLastYear.Sum(x => x.DSChiQuay),
                            DSChiNha = listDataItemLastYear.Sum(x => x.DSChiNha),
                            DSCK = listDataItemLastYear.Sum(x => x.DSCK),
                            TongDS = listDataItemLastYear.Sum(x => x.TongDS),
                            Year = (year - 1).ToString()
                        }
                    );
                }

                if (listDataGradationConvert.Count > 0)
                {
                    listDataGradation = new List<ReportDetailtServiceType>(listDataGradationConvert);
                }
            }

            foreach (ReportDetailtServiceType item in listDataGradation)
            {
                ReportDetailtServiceType dataItemYear = listDataGradation.Find(x => x.PartnerName == item.PartnerName && x.Year == year.ToString());
                ReportDetailtServiceType dataItemLastYear = listDataGradation.Find(x => x.PartnerName == item.PartnerName && x.Year == (year - 1).ToString());

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
            // Số record của mảng
            int countArray = 1;
            GradationCompare[] arrayGradation = new GradationCompare[countArray * listDataGradation.Count];
            int count = 0;
            foreach (ReportDetailtServiceType item in listDataGradation)
            {
                // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                arrayGradation[count] = new GradationCompare()
                {
                    NameGradationCompare = string.Format("Chi Quầy {0}", item.Year),
                    amount = item.DSChiQuay,
                    NameType = item.PartnerName
                };

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
        /// Get data cho việc vẽ biểu đồ cột cho so sánh theo giai đoạn theo danh số Chi nhà
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult SearchColumnsChartGradationCompareForOneForDSChiNha([DataSourceRequest]DataSourceRequest request, int gradation, int year, string reportTypeID, string marketID)
        {
            List<ReportDetailtServiceType> listDataGradation = new ReportBL().ReportDetailtGradationCompareForOne(year, gradation, reportTypeID, marketID);
            if (marketID.Contains("005"))
            {
                List<ReportDetailtServiceType> listDataGradationConvert = new List<ReportDetailtServiceType>();
                List<string> listMarket = new List<string>();

                foreach (ReportDetailtServiceType item in listDataGradation)
                {
                    if (!listMarket.Contains(item.MarketName))
                    {
                        listMarket.Add(item.MarketName);
                    }
                }

                foreach (string item in listMarket)
                {
                    List<ReportDetailtServiceType> listDataItemYear = listDataGradation.Where(x => x.MarketName == item && x.Year == year.ToString()).ToList();
                    List<ReportDetailtServiceType> listDataItemLastYear = listDataGradation.Where(x => x.MarketName == item && x.Year == (year - 1).ToString()).ToList();

                    // Year
                    listDataGradationConvert.Add(
                        new ReportDetailtServiceType()
                        {
                            MarketName = item,
                            PartnerName = item,
                            DSChiQuay = listDataItemYear.Sum(x => x.DSChiQuay),
                            DSChiNha = listDataItemYear.Sum(x => x.DSChiNha),
                            DSCK = listDataItemYear.Sum(x => x.DSCK),
                            TongDS = listDataItemYear.Sum(x => x.TongDS),
                            Year = year.ToString()
                        }
                    );

                    // Last Year
                    listDataGradationConvert.Add(
                        new ReportDetailtServiceType()
                        {
                            MarketName = item,
                            PartnerName = item,
                            DSChiQuay = listDataItemLastYear.Sum(x => x.DSChiQuay),
                            DSChiNha = listDataItemLastYear.Sum(x => x.DSChiNha),
                            DSCK = listDataItemLastYear.Sum(x => x.DSCK),
                            TongDS = listDataItemLastYear.Sum(x => x.TongDS),
                            Year = (year - 1).ToString()
                        }
                    );
                }

                if (listDataGradationConvert.Count > 0)
                {
                    listDataGradation = new List<ReportDetailtServiceType>(listDataGradationConvert);
                }
            }

            foreach (ReportDetailtServiceType item in listDataGradation)
            {
                ReportDetailtServiceType dataItemYear = listDataGradation.Find(x => x.PartnerName == item.PartnerName && x.Year == year.ToString());
                ReportDetailtServiceType dataItemLastYear = listDataGradation.Find(x => x.PartnerName == item.PartnerName && x.Year == (year - 1).ToString());

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
            // Số record của mảng
            int countArray = 1;
            GradationCompare[] arrayGradation = new GradationCompare[countArray * listDataGradation.Count];
            int count = 0;
            foreach (ReportDetailtServiceType item in listDataGradation)
            {
                // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                arrayGradation[count] = new GradationCompare()
                {
                    NameGradationCompare = string.Format("Chi Nhà {0}", item.Year),
                    amount = item.DSChiNha,
                    NameType = item.PartnerName
                };

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
        /// Get data cho việc vẽ biểu đồ cột cho so sánh theo giai đoạn theo danh số Chi nhà
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult SearchColumnsChartGradationCompareForOneForDSCK([DataSourceRequest]DataSourceRequest request, int gradation, int year, string reportTypeID, string marketID)
        {
            List<ReportDetailtServiceType> listDataGradation = new ReportBL().ReportDetailtGradationCompareForOne(year, gradation, reportTypeID, marketID);
            // Trường hợp thuộc châu Á
            if (marketID.Contains("005"))
            {
                List<ReportDetailtServiceType> listDataGradationConvert = new List<ReportDetailtServiceType>();
                List<string> listMarket = new List<string>();

                foreach (ReportDetailtServiceType item in listDataGradation)
                {
                    if (!listMarket.Contains(item.MarketName))
                    {
                        listMarket.Add(item.MarketName);
                    }
                }

                foreach (string item in listMarket)
                {
                    List<ReportDetailtServiceType> listDataItemYear = listDataGradation.Where(x => x.MarketName == item && x.Year == year.ToString()).ToList();
                    List<ReportDetailtServiceType> listDataItemLastYear = listDataGradation.Where(x => x.MarketName == item && x.Year == (year - 1).ToString()).ToList();

                    // Year
                    listDataGradationConvert.Add(
                        new ReportDetailtServiceType()
                        {
                            MarketName = item,
                            PartnerName = item,
                            DSChiQuay = listDataItemYear.Sum(x => x.DSChiQuay),
                            DSChiNha = listDataItemYear.Sum(x => x.DSChiNha),
                            DSCK = listDataItemYear.Sum(x => x.DSCK),
                            TongDS = listDataItemYear.Sum(x => x.TongDS),
                            Year = year.ToString()
                        }
                    );

                    // Last Year
                    listDataGradationConvert.Add(
                        new ReportDetailtServiceType()
                        {
                            MarketName = item,
                            PartnerName = item,
                            DSChiQuay = listDataItemLastYear.Sum(x => x.DSChiQuay),
                            DSChiNha = listDataItemLastYear.Sum(x => x.DSChiNha),
                            DSCK = listDataItemLastYear.Sum(x => x.DSCK),
                            TongDS = listDataItemLastYear.Sum(x => x.TongDS),
                            Year = (year - 1).ToString()
                        }
                    );
                }

                if (listDataGradationConvert.Count > 0)
                {
                    listDataGradation = new List<ReportDetailtServiceType>(listDataGradationConvert);
                }
            }

            foreach (ReportDetailtServiceType item in listDataGradation)
            {
                ReportDetailtServiceType dataItemYear = listDataGradation.Find(x => x.PartnerName == item.PartnerName && x.Year == year.ToString());
                ReportDetailtServiceType dataItemLastYear = listDataGradation.Find(x => x.PartnerName == item.PartnerName && x.Year == (year - 1).ToString());

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

            // Số record của mảng
            int countArray = 1;
            GradationCompare[] arrayGradation = new GradationCompare[countArray * listDataGradation.Count];
            int count = 0;
            foreach (ReportDetailtServiceType item in listDataGradation)
            {
                // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                arrayGradation[count] = new GradationCompare()
                {
                    NameGradationCompare = string.Format("Chuyển Khoản {0}", item.Year),
                    amount = item.DSCK,
                    NameType = item.PartnerName
                };

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

            if(marketID.Contains("005"))
            {
                List<ReportDetailtServiceType> listDataGradationConvert = new List<ReportDetailtServiceType>();
                List<string> listMarket = new List<string>();

                foreach (ReportDetailtServiceType item in listDataGradation)
                {
                    if (!listMarket.Contains(item.MarketName))
                    {
                        listMarket.Add(item.MarketName);
                    }
                }

                foreach (string item in listMarket)
                {
                    List<ReportDetailtServiceType> listDataItemYear = listDataGradation.Where(x => x.MarketName == item && x.Year == year.ToString()).ToList();
                    List<ReportDetailtServiceType> listDataItemLastYear = listDataGradation.Where(x => x.MarketName == item && x.Year == (year - 1).ToString()).ToList();

                    // Year
                    listDataGradationConvert.Add(
                        new ReportDetailtServiceType()
                        {
                            MarketName = item,
                            PartnerName = item,
                            DSChiQuay = listDataItemYear.Sum(x => x.DSChiQuay),
                            DSChiNha = listDataItemYear.Sum(x => x.DSChiNha),
                            DSCK = listDataItemYear.Sum(x => x.DSCK),
                            TongDS = listDataItemYear.Sum(x => x.TongDS),
                            Year = year.ToString()
                        }
                    );

                    // Last Year
                    listDataGradationConvert.Add(
                        new ReportDetailtServiceType()
                        {
                            MarketName = item,
                            PartnerName = item,
                            DSChiQuay = listDataItemLastYear.Sum(x => x.DSChiQuay),
                            DSChiNha = listDataItemLastYear.Sum(x => x.DSChiNha),
                            DSCK = listDataItemLastYear.Sum(x => x.DSCK),
                            TongDS = listDataItemLastYear.Sum(x => x.TongDS),
                            Year = (year - 1).ToString()
                        }
                    );
                }

                if (listDataGradationConvert.Count > 0)
                {
                    listDataGradation = new List<ReportDetailtServiceType>(listDataGradationConvert);
                }
            }

            foreach (ReportDetailtServiceType item in listDataGradation)
            {
                ReportDetailtServiceType dataItemYear = listDataGradation.Find(x => x.PartnerName == item.PartnerName && x.Year == year.ToString());
                ReportDetailtServiceType dataItemLastYear = listDataGradation.Find(x => x.PartnerName == item.PartnerName && x.Year == (year - 1).ToString());

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

            if (marketID.Contains("005"))
            {
                List<ReportDetailtServiceType> listDataGradationConvert = new List<ReportDetailtServiceType>();
                List<string> listMarket = new List<string>();
                foreach (ReportDetailtServiceType item in listDataGradation)
                {
                    if (!listMarket.Contains(item.MarketName))
                    {
                        listMarket.Add(item.MarketName);
                    }
                }

                foreach(string item in listMarket)
                {
                    List<ReportDetailtServiceType> dataItemYear = listDataGradation.Where(x => x.MarketName == item && x.Year == year.ToString()).ToList();
                    List<ReportDetailtServiceType> dataItemLastYear = listDataGradation.Where(x => x.MarketName == item && x.Year == (year - 1).ToString()).ToList();

                    // Year
                    listDataGradationConvert.Add(
                        new ReportDetailtServiceType()
                        {
                            MarketName = item,
                            PartnerName = item,
                            DSChiQuay = dataItemYear.Sum(x => x.DSChiQuay),
                            DSChiNha = dataItemYear.Sum(x => x.DSChiNha),
                            DSCK = dataItemYear.Sum(x => x.DSCK),
                            Year = year.ToString()
                        }
                    );
                    // Last Year
                    listDataGradationConvert.Add(
                        new ReportDetailtServiceType()
                        {
                            MarketName = item,
                            PartnerName = item,
                            DSChiQuay = dataItemLastYear.Sum(x => x.DSChiQuay),
                            DSChiNha = dataItemLastYear.Sum(x => x.DSChiNha),
                            DSCK = dataItemLastYear.Sum(x => x.DSCK),
                            Year = (year - 1).ToString()
                        }
                    );
                }

                if(listDataGradationConvert.Count > 0)
                {
                    listDataGradation = new List<ReportDetailtServiceType>(listDataGradationConvert);
                }
            }

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
                ReportDetailtServiceType dataItemLastYear = listDataGradation.Find(x => x.PartnerName == item.PartnerName && x.Year == (year - 1).ToString());
                ReportDetailtServiceType dataItemYear = listDataGradation.Find(x => x.PartnerName == item.PartnerName && x.Year == year.ToString());

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
                    table.Rows.Add(dataItemLastYear.PartnerName
                        , dataItemYear.DSChiQuay - dataItemLastYear.DSChiQuay
                        , dataItemYear.DSChiNha - dataItemLastYear.DSChiNha
                        , dataItemYear.DSCK - dataItemLastYear.DSCK
                        , dataItemYear.TongDS - dataItemLastYear.TongDS);
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
        public ActionResult SearchDataGradationComparePieYear([DataSourceRequest]DataSourceRequest request, int gradation, int year, string reportTypeID, string marketID)
        {
            List<ReportDetailtServiceType> listDataGradation = new ReportBL().ReportDetailtGradationCompareForOnePercent(year, gradation, reportTypeID, marketID);

            if (marketID.Contains("005"))
            {
                List<ReportDetailtServiceType> listDataGradationConvert = new List<ReportDetailtServiceType>();
                List<string> listMarket = new List<string>();
                foreach (ReportDetailtServiceType item in listDataGradation)
                {
                    if (!listMarket.Contains(item.MarketName))
                    {
                        listMarket.Add(item.MarketName);
                    }
                }

                foreach (string item in listMarket)
                {
                    List<ReportDetailtServiceType> dataItemYear = listDataGradation.Where(x => x.MarketName == item && x.Year == year.ToString()).ToList();
                    List<ReportDetailtServiceType> dataItemLastYear = listDataGradation.Where(x => x.MarketName == item && x.Year == (year - 1).ToString()).ToList();

                    // Year
                    listDataGradationConvert.Add(
                        new ReportDetailtServiceType()
                        {
                            MarketName = item,
                            PartnerName = item,
                            DSChiQuay = dataItemYear.Sum(x => x.DSChiQuay),
                            DSChiNha = dataItemYear.Sum(x => x.DSChiNha),
                            DSCK = dataItemYear.Sum(x => x.DSCK),
                            TongDS = dataItemYear.Sum(x => x.TongDS),
                            Year = year.ToString()
                        }
                    );
                    // Last Year
                    listDataGradationConvert.Add(
                        new ReportDetailtServiceType()
                        {
                            MarketName = item,
                            PartnerName = item,
                            DSChiQuay = dataItemLastYear.Sum(x => x.DSChiQuay),
                            DSChiNha = dataItemLastYear.Sum(x => x.DSChiNha),
                            DSCK = dataItemLastYear.Sum(x => x.DSCK),
                            TongDS = dataItemLastYear.Sum(x => x.TongDS),
                            Year = (year - 1).ToString()
                        }
                    );
                }

                if (listDataGradationConvert.Count > 0)
                {
                    listDataGradation = new List<ReportDetailtServiceType>(listDataGradationConvert);
                }
            }
            
            // Get dữ liệu của năm hiện tại
            List<ReportDetailtServiceType> listData = listDataGradation.Where(x => x.Year == year.ToString()).ToList();
            
            GradationChartPie[] arrayGradation = new GradationChartPie[listData.Count()];
            int count = 0;
            List<string> listColor = new List<string>() { "#FFBF00", "#40FF00", "#2ECCFA", "#9A2EFE", "#FE2EF7", "#0000FF", "#08088A", "#B40431", "#6E6E6E" };

            foreach (ReportDetailtServiceType item in listData)
            {
                // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                arrayGradation[count] = new GradationChartPie()
                {
                    category = item.PartnerName,
                    value = Math.Round(item.TongDS, 2, MidpointRounding.ToEven),
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
            List<ReportDetailtServiceType> listDataGradation = new ReportBL().ReportDetailtGradationCompareForOnePercent(year, gradation, reportTypeID, marketID);

            if (marketID.Contains("005"))
            {
                List<ReportDetailtServiceType> listDataGradationConvert = new List<ReportDetailtServiceType>();
                List<string> listMarket = new List<string>();
                foreach (ReportDetailtServiceType item in listDataGradation)
                {
                    if (!listMarket.Contains(item.MarketName))
                    {
                        listMarket.Add(item.MarketName);
                    }
                }

                foreach (string item in listMarket)
                {
                    List<ReportDetailtServiceType> dataItemYear = listDataGradation.Where(x => x.MarketName == item && x.Year == year.ToString()).ToList();
                    List<ReportDetailtServiceType> dataItemLastYear = listDataGradation.Where(x => x.MarketName == item && x.Year == (year - 1).ToString()).ToList();

                    // Year
                    listDataGradationConvert.Add(
                        new ReportDetailtServiceType()
                        {
                            MarketName = item,
                            PartnerName = item,
                            DSChiQuay = dataItemYear.Sum(x => x.DSChiQuay),
                            DSChiNha = dataItemYear.Sum(x => x.DSChiNha),
                            DSCK = dataItemYear.Sum(x => x.DSCK),
                            TongDS = dataItemYear.Sum(x => x.TongDS),
                            Year = year.ToString()
                        }
                    );
                    // Last Year
                    listDataGradationConvert.Add(
                        new ReportDetailtServiceType()
                        {
                            MarketName = item,
                            PartnerName = item,
                            DSChiQuay = dataItemLastYear.Sum(x => x.DSChiQuay),
                            DSChiNha = dataItemLastYear.Sum(x => x.DSChiNha),
                            DSCK = dataItemLastYear.Sum(x => x.DSCK),
                            TongDS = dataItemLastYear.Sum(x => x.TongDS),
                            Year = (year - 1).ToString()
                        }
                    );
                }

                if (listDataGradationConvert.Count > 0)
                {
                    listDataGradation = new List<ReportDetailtServiceType>(listDataGradationConvert);
                }
            }

            // Get dữ liệu của năm hiện tại
            List<ReportDetailtServiceType> listData = listDataGradation.Where(x => x.Year == (year - 1).ToString()).ToList();

            GradationChartPie[] arrayGradation = new GradationChartPie[listData.Count()];
            int count = 0;
            List<string> listColor = new List<string>() { "#FFBF00", "#40FF00", "#2ECCFA", "#9A2EFE", "#FE2EF7", "#0000FF", "#08088A", "#B40431", "#6E6E6E" };

            foreach (ReportDetailtServiceType item in listData)
            {
                // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                arrayGradation[count] = new GradationChartPie()
                {
                    category = item.PartnerName,
                    value = Math.Round(item.TongDS, 2, MidpointRounding.ToEven),
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
        public ActionResult SearchDataDetailtGridGradationComparePercent([DataSourceRequest]DataSourceRequest request, int gradation, int year, string reportTypeID, string marketID)
        {
            List<ReportDetailtServiceType> listDataGradation = new ReportBL().ReportDetailtGradationCompareForOnePercent(year, gradation, reportTypeID, marketID);
            
            if(marketID.Contains("005"))
            {
                List<ReportDetailtServiceType> listDataGradationConvert = new List<ReportDetailtServiceType>();
                List<string> listMarket = new List<string>();
                foreach (ReportDetailtServiceType item in listDataGradation)
                {
                    if (!listMarket.Contains(item.MarketName))
                    {
                        listMarket.Add(item.MarketName);
                    }
                }

                foreach(string item in listMarket)
                {
                    List<ReportDetailtServiceType> listDataItemYear = listDataGradation.Where(x => x.MarketName == item && x.Year == year.ToString()).ToList();
                    List<ReportDetailtServiceType> listDataItemLastYear = listDataGradation.Where(x => x.MarketName == item && x.Year == (year - 1).ToString()).ToList();

                    // Year
                    listDataGradationConvert.Add(
                        new ReportDetailtServiceType()
                        {
                            MarketName = item,
                            PartnerName = item,
                            DSChiQuay = listDataItemYear.Sum(x => x.DSChiQuay),
                            DSChiNha = listDataItemYear.Sum(x => x.DSChiNha),
                            DSCK = listDataItemYear.Sum(x => x.DSCK),
                            TongDS = listDataItemYear.Sum(x => x.TongDS),
                            Year = year.ToString()
                        }
                    );

                    // Last Year
                    listDataGradationConvert.Add(
                        new ReportDetailtServiceType()
                        {
                            MarketName = item,
                            PartnerName = item,
                            DSChiQuay = listDataItemLastYear.Sum(x => x.DSChiQuay),
                            DSChiNha = listDataItemLastYear.Sum(x => x.DSChiNha),
                            DSCK = listDataItemLastYear.Sum(x => x.DSCK),
                            TongDS = listDataItemLastYear.Sum(x => x.TongDS),
                            Year = (year - 1).ToString()
                        }
                    );
                }

                if (listDataGradationConvert.Count > 0)
                {
                    listDataGradation = new List<ReportDetailtServiceType>(listDataGradationConvert);
                }
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
                ReportDetailtServiceType listDataYear = listDataGradation.Find(x => x.Year == year.ToString() && x.PartnerName == item.PartnerName);
                // Get dữ liệu của năm trước
                ReportDetailtServiceType listDataLastYear = listDataGradation.Find(x => x.Year == (year - 1).ToString() && x.PartnerName == item.PartnerName);

                // Check tồn tại của item
                string value = string.Format("PartnerName='{0}'", item.PartnerName);
                DataRow[] foundRows = table.Select(value);

                if (listDataYear != null && listDataLastYear != null && foundRows.Count() == 0)
                {
                    double valueYear = Math.Round(listDataYear.TongDS, 2, MidpointRounding.ToEven);
                    double valueLastYear = Math.Round(listDataLastYear.TongDS, 2, MidpointRounding.ToEven);
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
        /// Get data cho báo cáo chi tiết so sánh theo tháng
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult SearchReportDetailtCompareMonthForAll([DataSourceRequest]DataSourceRequest request, int month, int year, string reportTypeID, string marketID)
        {
            List<ReportDetailtSTMarket> listDataCompareMonth = new ReportBL().ReportDetailtCompareMonthForAll(year, month, reportTypeID, marketID);

            List<string> listMarket = new List<string>();
            // Thị trường Châu Á
            if (!marketID.Contains("0"))
            {
                List<ReportDetailtSTMarket> listDataCompareMonthConvert = new List<ReportDetailtSTMarket>();
                // List thị trường
                foreach(ReportDetailtSTMarket item in listDataCompareMonth)
                {
                    if(!listMarket.Contains(item.MarketName))
                    {
                        listMarket.Add(item.MarketName);
                    }
                }

                foreach (string item in listMarket)
                {
                    // month, year recent
                    List<ReportDetailtSTMarket> listDataItemMonth = listDataCompareMonth.Where(x => x.MarketName == item && x.Month == month.ToString() && x.Year == year.ToString()).ToList();
                    // last month, year
                    List<ReportDetailtSTMarket> listDataItemLastMonth = listDataCompareMonth.Where(x => x.MarketName == item && x.Month == (month - 1).ToString() && x.Year == year.ToString()).ToList();

                    // Trường hợp tháng 1
                    if (month == 1)
                    {
                        listDataItemLastMonth = listDataCompareMonth.Where(x => x.MarketName == item && x.Month == "12" && x.Year == (year - 1).ToString()).ToList();
                    }
                    // month, last year
                    List<ReportDetailtSTMarket> listDataItemMonthLastYear = listDataCompareMonth.Where(x => x.MarketName == item && x.Month == month.ToString() && x.Year == (year - 1).ToString()).ToList();

                    if(listDataItemMonth.Count == 0)
                    {
                        listDataItemMonth = new List<ReportDetailtSTMarket>()
                        {
                            new ReportDetailtSTMarket()
                            {
                                MarketName = item,
                                Month = month.ToString(),
                                Year = year.ToString()
                            }
                        };
                    }

                    // Last month
                    if (listDataItemLastMonth.Count == 0)
                    {
                        listDataItemLastMonth = new List<ReportDetailtSTMarket>()
                        {
                            new ReportDetailtSTMarket()
                            {
                                MarketName = item,
                                Month = (month - 1).ToString(),
                                Year = year.ToString()
                            }
                        };
                    }

                    // Month, Last year
                    if (listDataItemMonthLastYear.Count == 0)
                    {
                        listDataItemMonthLastYear = new List<ReportDetailtSTMarket>()
                        {
                            new ReportDetailtSTMarket()
                            {
                                MarketName = item,
                                Month = month.ToString(),
                                Year = (year- 1).ToString()
                            }
                        };
                    }

                    // Month recent
                    listDataCompareMonthConvert.Add(
                        new ReportDetailtSTMarket()
                        {
                            MarketName = item,
                            DSChiQuay = listDataItemMonth.Sum(x => x.DSChiQuay),
                            DSChiNha = listDataItemMonth.Sum(x => x.DSChiNha),
                            DSCK = listDataItemMonth.Sum(x => x.DSCK),
                            Month = month.ToString(),
                            Year = year.ToString()
                        }
                    );

                    // Last month
                    listDataCompareMonthConvert.Add(
                        new ReportDetailtSTMarket()
                        {
                            MarketName = item,
                            DSChiQuay = listDataItemLastMonth.Sum(x => x.DSChiQuay),
                            DSChiNha = listDataItemLastMonth.Sum(x => x.DSChiNha),
                            DSCK = listDataItemLastMonth.Sum(x => x.DSCK),
                            Month = (month - 1).ToString(),
                            Year = year.ToString()
                        }
                    );

                    // month Last year
                    listDataCompareMonthConvert.Add(
                        new ReportDetailtSTMarket()
                        {
                            MarketName = item,
                            DSChiQuay = listDataItemMonthLastYear.Sum(x => x.DSChiQuay),
                            DSChiNha = listDataItemMonthLastYear.Sum(x => x.DSChiNha),
                            DSCK = listDataItemMonthLastYear.Sum(x => x.DSCK),
                            Month = month.ToString(),
                            Year = (year - 1).ToString()
                        }
                    );
                }

                if (listDataCompareMonthConvert.Count > 0)
                {
                    listDataCompareMonth = new List<ReportDetailtSTMarket>(listDataCompareMonthConvert);
                }
            }
            

            DataTable table = new DataTable();
            // Khởi tạo datatable
            table = new DataTable();
            // Tạo các cột cho datatable
            // tháng hiện tại
            table.Columns.Add("MarketName", typeof(String));
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

            table.Columns.Add("ReportID", typeof(String));

            
            // Danh sách mã thị trường của Tất cả
            if (marketID.Contains("0"))
            {
                foreach (ReportDetailtSTMarket item in listDataCompareMonth)
                {
                    if (!listMarket.Contains(item.MarketName))
                    {
                        listMarket.Add(item.MarketName);
                    }

                    item.ReportID = item.MarketName;
                    item.TongDS = item.DSChiQuay + item.DSChiNha + item.DSCK;
                    item.MarketName = "Tất cả thị trường";
                }
            }
            else
            {
                foreach (ReportDetailtSTMarket item in listDataCompareMonth)
                {
                    item.ReportID = item.MarketName;
                    item.TongDS = item.DSChiQuay + item.DSChiNha + item.DSCK;
                    item.MarketName = "Thị trường châu Á";
                }
            }

            foreach (string item in listMarket)
            {
                // Cùng kì
                ReportDetailtSTMarket dataItemLastYear = listDataCompareMonth.Find(x => x.ReportID == item && x.Month == month.ToString() && x.Year == (year - 1).ToString());
                ReportDetailtSTMarket dataItemYear = listDataCompareMonth.Find(x => x.ReportID == item && x.Month == month.ToString() && x.Year == year.ToString());
                ReportDetailtSTMarket dataItemLastMonth = listDataCompareMonth.Find(x => x.ReportID == item && x.Month == (month - 1).ToString() && x.Year == year.ToString());
                // Trường hợp tháng 1
                if (month == 1)
                {
                    dataItemLastMonth = listDataCompareMonth.Find(x => x.MarketCode == item && x.Month == "12" && x.Year == (year - 1).ToString());
                }

                if (dataItemLastYear != null && dataItemYear != null && dataItemLastMonth != null)
                {
                    // add item vào table
                    table.Rows.Add(dataItemYear.MarketName
                        , dataItemYear.DSChiQuay, dataItemLastMonth.DSChiQuay, dataItemLastYear.DSChiQuay
                        , dataItemYear.DSChiNha, dataItemLastMonth.DSChiNha, dataItemLastYear.DSChiNha
                        , dataItemYear.DSCK, dataItemLastMonth.DSCK, dataItemLastYear.DSCK
                        , dataItemYear.TongDS, dataItemLastMonth.TongDS, dataItemLastYear.TongDS
                        , dataItemYear.ReportID);
                }
            }

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
        public ActionResult SearchReportDetailtCompareMonthForAllCompare([DataSourceRequest]DataSourceRequest request, int month, int year, string reportTypeID, string marketID)
        {
            List<ReportDetailtSTMarket> listDataCompareMonth = new ReportBL().ReportDetailtCompareMonthForAll(year, month, reportTypeID, marketID);


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

            table.Columns.Add("ReportID", typeof(String));

            // trường hợp tất cả thị trường
            if (marketID.Equals("0"))
            {
                foreach (ReportDetailtSTMarket item in listDataCompareMonth)
                {
                    item.ReportID = item.MarketName;
                    item.TongDS = item.DSChiQuay + item.DSChiNha + item.DSCK;
                    item.MarketName = "Tất cả thị trường";
                }


                foreach (ReportDetailtSTMarket item in listDataCompareMonth)
                {
                    // Cùng kì
                    ReportDetailtSTMarket dataItemLastYear = listDataCompareMonth.Find(x => x.MarketCode == item.MarketCode && x.Month == month.ToString() && x.Year == (year - 1).ToString());
                    ReportDetailtSTMarket dataItemMonth = listDataCompareMonth.Find(x => x.MarketCode == item.MarketCode && x.Month == month.ToString() && x.Year == year.ToString());
                    ReportDetailtSTMarket dataItemLastMonth = listDataCompareMonth.Find(x => x.MarketCode == item.MarketCode && x.Month == (month - 1).ToString() && x.Year == year.ToString());
                    // Trường hợp tháng 1
                    if (month == 1)
                    {
                        dataItemLastMonth = listDataCompareMonth.Find(x => x.MarketCode == item.MarketCode && x.Month == "12" && x.Year == (year - 1).ToString());
                    }

                    // Check tồn tại của item
                    string value = string.Format("ReportID='{0}'", item.ReportID);

                    DataRow[] foundRows = table.Select(value);

                    if (dataItemLastYear != null && dataItemMonth != null && dataItemLastMonth != null && foundRows.Count() == 0)
                    {
                        // add item vào table
                        table.Rows.Add(item.MarketName
                            , dataItemMonth.DSChiQuay - dataItemLastMonth.DSChiQuay, dataItemMonth.DSChiNha - dataItemLastMonth.DSChiNha, dataItemMonth.DSCK - dataItemLastMonth.DSCK, dataItemMonth.TongDS - dataItemLastMonth.TongDS
                            , dataItemMonth.DSChiQuay - dataItemLastYear.DSChiQuay, dataItemMonth.DSChiNha - dataItemLastYear.DSChiNha, dataItemMonth.DSCK - dataItemLastYear.DSCK, dataItemMonth.TongDS - dataItemLastYear.TongDS
                            , item.ReportID);
                    }
                }
            }
            else
            {
                List<string> listMarket = new List<string>();
                // Trường hợp các thị trường con của thị trường châu Á
                foreach (ReportDetailtSTMarket item in listDataCompareMonth)
                {
                    if (!listMarket.Contains(item.MarketName))
                    {
                        listMarket.Add(item.MarketName);
                    }

                    item.ReportID = item.MarketName;
                    item.TongDS = item.DSChiQuay + item.DSChiNha + item.DSCK;
                }

                foreach (string item in listMarket)
                {
                    // Cùng kì
                    List<ReportDetailtSTMarket> dataItemLastYear = listDataCompareMonth.Where(x => x.ReportID == item && x.Month == month.ToString() && x.Year == (year - 1).ToString()).ToList();
                    List<ReportDetailtSTMarket> dataItemMonth = listDataCompareMonth.Where(x => x.ReportID == item && x.Month == month.ToString() && x.Year == year.ToString()).ToList();
                    List<ReportDetailtSTMarket> dataItemLastMonth = listDataCompareMonth.Where(x => x.ReportID == item && x.Month == (month - 1).ToString() && x.Year == year.ToString()).ToList();
                    // Trường hợp tháng 1
                    if (month == 1)
                    {
                        dataItemLastMonth = listDataCompareMonth.Where(x => x.ReportID == item && x.Month == "12" && x.Year == (year - 1).ToString()).ToList();
                    }

                    // Cùng kì năm trước
                    if (dataItemLastYear.Count == 0)
                    {
                        dataItemLastYear = new List<ReportDetailtSTMarket>()
                        {
                            new ReportDetailtSTMarket()
                            {
                                ReportID = item,
                                Month = month.ToString(),
                                Year = (year - 1).ToString()
                            }
                        };
                    }

                    // Hiện tại
                    if (dataItemMonth.Count == 0)
                    {
                        dataItemMonth = new List<ReportDetailtSTMarket>()
                        {
                            new ReportDetailtSTMarket()
                            {
                                ReportID = item,
                                Month = month.ToString(),
                                Year = year.ToString()
                            }
                        };
                    }

                    // tháng trước
                    if (dataItemLastMonth.Count == 0)
                    {
                        if (month == 1)
                        {
                            dataItemLastMonth = new List<ReportDetailtSTMarket>()
                            {
                                new ReportDetailtSTMarket()
                                {
                                    ReportID = item,
                                    Month = "12",
                                    Year = (year - 1).ToString()
                                }
                            };
                        }
                        else
                        {
                            dataItemLastMonth = new List<ReportDetailtSTMarket>()
                            {
                                new ReportDetailtSTMarket()
                                {
                                    ReportID = item,
                                    Month = (month - 1).ToString(),
                                    Year = year.ToString()
                                }
                            };
                        }
                    }

                    // Check tồn tại của item
                    string value = string.Format("ReportID='{0}'", item);

                    DataRow[] foundRows = table.Select(value);

                    if (dataItemLastYear != null && dataItemMonth != null && dataItemLastMonth != null && foundRows.Count() == 0)
                    {
                        // add item vào table
                        table.Rows.Add("Thị trường Châu Á"
                            , dataItemMonth.Sum(x=>x.DSChiQuay) - dataItemLastMonth.Sum(x => x.DSChiQuay)
                            , dataItemMonth.Sum(x => x.DSChiNha) - dataItemLastMonth.Sum(x => x.DSChiNha)
                            , dataItemMonth.Sum(x => x.DSCK) - dataItemLastMonth.Sum(x => x.DSCK)
                            , dataItemMonth.Sum(x => x.TongDS) - dataItemLastMonth.Sum(x => x.TongDS)
                            , dataItemMonth.Sum(x => x.DSChiQuay) - dataItemLastYear.Sum(x => x.DSChiQuay)
                            , dataItemMonth.Sum(x => x.DSChiNha) - dataItemLastYear.Sum(x => x.DSChiNha)
                            , dataItemMonth.Sum(x => x.DSCK) - dataItemLastYear.Sum(x => x.DSCK)
                            , dataItemMonth.Sum(x => x.TongDS) - dataItemLastYear.Sum(x => x.TongDS)
                            , item);
                    }

                }
            }

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
        public ActionResult SearchColumnChartStackCompareMonthForAll([DataSourceRequest]DataSourceRequest request, int month, int year, string reportTypeID, string marketID)
        {
            List<ReportDetailtSTMarket> listDataCompareMonth = new ReportBL().ColumnChartStackCompareMonthForAllPercent(year, month, reportTypeID, marketID);

            GradationCharColumn[] arrayGradation = null;

            List<ReportDetailtSTMarket> listDataYear = listDataCompareMonth.Where(x => x.Month == month.ToString() && x.Year == year.ToString()).ToList();
            List<ReportDetailtSTMarket> listDataLastMonth = listDataCompareMonth.Where(x => x.Month == (month - 1).ToString() && x.Year == year.ToString()).ToList();
            // Trường hợp tháng 1
            if (month == 1)
            {
                listDataLastMonth = listDataCompareMonth.Where(x => x.Month == "12" && x.Year == (year - 1).ToString()).ToList();
            }
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
        /// Get data cho việc vẽ biểu đồ cột cho so sánh giai đoạn theo chi Quầy
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult SearchColumnsChartCompareMonthForAllForDSChiQuay([DataSourceRequest]DataSourceRequest request, int month, int year, string reportTypeID, string marketID)
        {
            List<ReportDetailtSTMarket> listDataCompareMonth = new ReportBL().ReportDetailtCompareMonthForAll(year, month, reportTypeID, marketID);

            GradationCompare[] arrayGradation = null;

            // Trường hợp tất cả thị trường
            if (marketID.Equals("0"))
            {
                List<ReportDetailtSTMarket> listDataYear = listDataCompareMonth.Where(x => x.Month == month.ToString() && x.Year == year.ToString()).ToList();

                List<ReportDetailtSTMarket> listDataLastMonth = listDataCompareMonth.Where(x => x.Month == (month - 1).ToString() && x.Year == year.ToString()).ToList();
                // Trường hợp tháng 1
                if (month == 1)
                {
                    listDataLastMonth = listDataCompareMonth.Where(x => x.Month == "12" && x.Year == (year - 1).ToString()).ToList();
                }

                List<ReportDetailtSTMarket> listDataLastYear = listDataCompareMonth.Where(x => x.Month == month.ToString() && x.Year == (year - 1).ToString()).ToList();

                // Trường hợp đủ dữ liệu tháng hiện tại, tháng trước và cùng kì năm trước
                if (listDataYear.Count > 0 && listDataLastMonth.Count > 0 && listDataLastYear.Count > 0)
                {
                    // Số record của mảng
                    int countArray = 1;
                    arrayGradation = new GradationCompare[countArray * listDataCompareMonth.Count];
                    int count = 0;

                    foreach (ReportDetailtSTMarket item in listDataCompareMonth)
                    {
                        // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                        arrayGradation[count] = new GradationCompare()
                        {
                            NameGradationCompare = string.Format("Chi Quầy {0}/{1}", item.Month, item.Year),
                            amount = item.DSChiQuay,
                            NameType = item.MarketName
                        };

                        count++;
                        year = year - 1;
                    }
                }
            }
            else
            {
                List<string> listAsianItem = new List<string>();
                foreach (ReportDetailtSTMarket item in listDataCompareMonth)
                {
                    if (!listAsianItem.Contains(item.MarketName))
                    {
                        listAsianItem.Add(item.MarketName);
                    }
                }

                // Số record của mảng
                int countArray = 3;
                arrayGradation = new GradationCompare[countArray * listAsianItem.Count];
                int count = 0;

                List<ReportDetailtSTMarket> listDataYear = null;
                List<ReportDetailtSTMarket> listDataLastMonth = null;
                List<ReportDetailtSTMarket> listDataLastYear = null;

                foreach (string item in listAsianItem)
                {
                    listDataYear = listDataCompareMonth.Where(x => x.MarketName == item && x.Month == month.ToString() && x.Year == year.ToString()).ToList();
                    listDataLastMonth = listDataCompareMonth.Where(x => x.MarketName == item && x.Month == (month - 1).ToString() && x.Year == year.ToString()).ToList();
                    // Trường hợp tháng 1
                    if (month == 1)
                    {
                        listDataLastMonth = listDataCompareMonth.Where(x => x.Month == "12" && x.Year == (year - 1).ToString()).ToList();
                    }
                    listDataLastYear = listDataCompareMonth.Where(x => x.MarketName == item && x.Month == month.ToString() && x.Year == (year - 1).ToString()).ToList();

                    if (listDataYear == null)
                    {
                        listDataYear = new List<ReportDetailtSTMarket>();
                    }

                    if (listDataLastMonth == null)
                    {
                        listDataLastMonth = new List<ReportDetailtSTMarket>();
                    }

                    if (listDataLastYear == null)
                    {
                        listDataLastYear = new List<ReportDetailtSTMarket>();
                    }

                    // Year
                    double sumDSChiQuayYear = listDataYear.Sum(x => x.DSChiQuay);
                    double sumDSChiNhaYear = listDataYear.Sum(x => x.DSChiNha);
                    double sumDSCKYear = listDataYear.Sum(x => x.DSCK);
                    double sumTongDSYear = listDataYear.Sum(x => x.TongDS);

                    // Last Month
                    double sumDSChiQuayLastMonth = listDataLastMonth.Sum(x => x.DSChiQuay);
                    double sumDSChiNhaLastMonth = listDataLastMonth.Sum(x => x.DSChiNha);
                    double sumDSCKLastMonth = listDataLastMonth.Sum(x => x.DSCK);
                    double sumTongDSLastMonth = listDataLastMonth.Sum(x => x.TongDS);

                    // LastYear
                    double sumDSChiQuayLastYear = listDataLastYear.Sum(x => x.DSChiQuay);
                    double sumDSChiNhaLastYear = listDataLastYear.Sum(x => x.DSChiNha);
                    double sumDSCKLastYear = listDataLastYear.Sum(x => x.DSCK);
                    double sumTongDSLastYear = listDataLastYear.Sum(x => x.TongDS);

                    // Tạo mảng insert dữ liệu để vẽ biểu đồ cột

                    // Tháng hiện tại
                    arrayGradation[count] = new GradationCompare()
                    {
                        NameGradationCompare = string.Format("Chi Quầy {0}/{1}", month, year),
                        amount = sumDSChiQuayYear,
                        NameType = item
                    };

                    count++;

                    // Tháng trước
                    arrayGradation[count] = new GradationCompare()
                    {
                        NameGradationCompare = string.Format("Chi Quầy {0}/{1}", month - 1, year),
                        amount = sumDSChiQuayLastMonth,
                        NameType = item
                    };

                    count++;

                    // Cùng kì năm trước
                    arrayGradation[count] = new GradationCompare()
                    {
                        NameGradationCompare = string.Format("Chi Quầy {0}/{1}", month, year - 1),
                        amount = sumDSChiQuayLastYear,
                        NameType = item
                    };

                    count++;
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
        /// Get data cho việc vẽ biểu đồ cột cho so sánh giai đoạn theo chi Nha
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult SearchColumnsChartCompareMonthForAllForDSChiNha([DataSourceRequest]DataSourceRequest request, int month, int year, string reportTypeID, string marketID)
        {
            List<ReportDetailtSTMarket> listDataCompareMonth = new ReportBL().ReportDetailtCompareMonthForAll(year, month, reportTypeID, marketID);

            GradationCompare[] arrayGradation = null;

            // Trường hợp tất cả thị trường
            if (marketID.Equals("0"))
            {
                List<ReportDetailtSTMarket> listDataYear = listDataCompareMonth.Where(x => x.Month == month.ToString() && x.Year == year.ToString()).ToList();
                List<ReportDetailtSTMarket> listDataLastMonth = listDataCompareMonth.Where(x => x.Month == (month - 1).ToString() && x.Year == year.ToString()).ToList();
                // Trường hợp tháng 1
                if (month == 1)
                {
                    listDataLastMonth = listDataCompareMonth.Where(x => x.Month == "12" && x.Year == (year - 1).ToString()).ToList();
                }
                List<ReportDetailtSTMarket> listDataLastYear = listDataCompareMonth.Where(x => x.Month == month.ToString() && x.Year == (year - 1).ToString()).ToList();

                // Trường hợp đủ dữ liệu tháng hiện tại, tháng trước và cùng kì năm trước
                if (listDataYear.Count > 0 && listDataLastMonth.Count > 0 && listDataLastYear.Count > 0)
                {
                    // Số record của mảng
                    int countArray = 1;
                    arrayGradation = new GradationCompare[countArray * listDataCompareMonth.Count];
                    int count = 0;

                    foreach (ReportDetailtSTMarket item in listDataCompareMonth)
                    {
                        // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                        arrayGradation[count] = new GradationCompare()
                        {
                            NameGradationCompare = string.Format("Chi Nhà {0}/{1}", item.Month, item.Year),
                            amount = item.DSChiNha,
                            NameType = item.MarketName
                        };

                        count++;
                        year = year - 1;
                    }
                }
            }
            else
            {
                List<string> listAsianItem = new List<string>();
                foreach (ReportDetailtSTMarket item in listDataCompareMonth)
                {
                    if (!listAsianItem.Contains(item.MarketName))
                    {
                        listAsianItem.Add(item.MarketName);
                    }
                }

                // Số record của mảng
                int countArray = 3;
                arrayGradation = new GradationCompare[countArray * listAsianItem.Count];
                int count = 0;

                List<ReportDetailtSTMarket> listDataYear = null;
                List<ReportDetailtSTMarket> listDataLastMonth = null;
                List<ReportDetailtSTMarket> listDataLastYear = null;

                foreach (string item in listAsianItem)
                {
                    listDataYear = listDataCompareMonth.Where(x => x.MarketName == item && x.Month == month.ToString() && x.Year == year.ToString()).ToList();
                    listDataLastMonth = listDataCompareMonth.Where(x => x.MarketName == item && x.Month == (month - 1).ToString() && x.Year == year.ToString()).ToList();
                    // Trường hợp tháng 1
                    if (month == 1)
                    {
                        listDataLastMonth = listDataCompareMonth.Where(x => x.Month == "12" && x.Year == (year - 1).ToString()).ToList();
                    }
                    listDataLastYear = listDataCompareMonth.Where(x => x.MarketName == item && x.Month == month.ToString() && x.Year == (year - 1).ToString()).ToList();

                    if (listDataYear == null)
                    {
                        listDataYear = new List<ReportDetailtSTMarket>();
                    }

                    if (listDataLastMonth == null)
                    {
                        listDataLastMonth = new List<ReportDetailtSTMarket>();
                    }

                    if (listDataLastYear == null)
                    {
                        listDataLastYear = new List<ReportDetailtSTMarket>();
                    }

                    // Year
                    double sumDSChiQuayYear = listDataYear.Sum(x => x.DSChiQuay);
                    double sumDSChiNhaYear = listDataYear.Sum(x => x.DSChiNha);
                    double sumDSCKYear = listDataYear.Sum(x => x.DSCK);
                    double sumTongDSYear = listDataYear.Sum(x => x.TongDS);

                    // Last Month
                    double sumDSChiQuayLastMonth = listDataLastMonth.Sum(x => x.DSChiQuay);
                    double sumDSChiNhaLastMonth = listDataLastMonth.Sum(x => x.DSChiNha);
                    double sumDSCKLastMonth = listDataLastMonth.Sum(x => x.DSCK);
                    double sumTongDSLastMonth = listDataLastMonth.Sum(x => x.TongDS);

                    // LastYear
                    double sumDSChiQuayLastYear = listDataLastYear.Sum(x => x.DSChiQuay);
                    double sumDSChiNhaLastYear = listDataLastYear.Sum(x => x.DSChiNha);
                    double sumDSCKLastYear = listDataLastYear.Sum(x => x.DSCK);
                    double sumTongDSLastYear = listDataLastYear.Sum(x => x.TongDS);

                    // Tạo mảng insert dữ liệu để vẽ biểu đồ cột

                    // Tháng hiện tại
                    arrayGradation[count] = new GradationCompare()
                    {
                        NameGradationCompare = string.Format("Chi Nhà {0}/{1}", month, year),
                        amount = sumDSChiNhaYear,
                        NameType = item
                    };

                    count++;

                    // Tháng trước
                    arrayGradation[count] = new GradationCompare()
                    {
                        NameGradationCompare = string.Format("Chi Nhà {0}/{1}", month - 1, year),
                        amount = sumDSChiNhaLastMonth,
                        NameType = item
                    };

                    count++;

                    // Cùng kì năm trước
                    arrayGradation[count] = new GradationCompare()
                    {
                        NameGradationCompare = string.Format("Chi Nhà {0}/{1}", month, year - 1),
                        amount = sumDSChiNhaLastYear,
                        NameType = item
                    };

                    count++;
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
        /// Get data cho việc vẽ biểu đồ cột cho so sánh giai đoạn theo chi Quầy
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult SearchColumnsChartCompareMonthForAllForDSCK([DataSourceRequest]DataSourceRequest request, int month, int year, string reportTypeID, string marketID)
        {
            List<ReportDetailtSTMarket> listDataCompareMonth = new ReportBL().ReportDetailtCompareMonthForAll(year, month, reportTypeID, marketID);

            GradationCompare[] arrayGradation = null;

            // Trường hợp tất cả thị trường
            if (marketID.Equals("0"))
            {
                List<ReportDetailtSTMarket> listDataYear = listDataCompareMonth.Where(x => x.Month == month.ToString() && x.Year == year.ToString()).ToList();
                List<ReportDetailtSTMarket> listDataLastMonth = listDataCompareMonth.Where(x => x.Month == (month - 1).ToString() && x.Year == year.ToString()).ToList();
                // Trường hợp tháng 1
                if (month == 1)
                {
                    listDataLastMonth = listDataCompareMonth.Where(x => x.Month == "12" && x.Year == (year - 1).ToString()).ToList();
                }
                List<ReportDetailtSTMarket> listDataLastYear = listDataCompareMonth.Where(x => x.Month == month.ToString() && x.Year == (year - 1).ToString()).ToList();

                // Trường hợp đủ dữ liệu tháng hiện tại, tháng trước và cùng kì năm trước
                if (listDataYear.Count > 0 && listDataLastMonth.Count > 0 && listDataLastYear.Count > 0)
                {
                    // Số record của mảng
                    int countArray = 1;
                    arrayGradation = new GradationCompare[countArray * listDataCompareMonth.Count];
                    int count = 0;

                    foreach (ReportDetailtSTMarket item in listDataCompareMonth)
                    {
                        // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                        arrayGradation[count] = new GradationCompare()
                        {
                            NameGradationCompare = string.Format("DSCK {0}/{1}", item.Month, item.Year),
                            amount = item.DSCK,
                            NameType = item.MarketName
                        };

                        count++;
                        year = year - 1;
                    }
                }
            }
            else
            {
                List<string> listAsianItem = new List<string>();
                foreach (ReportDetailtSTMarket item in listDataCompareMonth)
                {
                    if (!listAsianItem.Contains(item.MarketName))
                    {
                        listAsianItem.Add(item.MarketName);
                    }
                }

                // Số record của mảng
                int countArray = 3;
                arrayGradation = new GradationCompare[countArray * listAsianItem.Count];
                int count = 0;

                List<ReportDetailtSTMarket> listDataYear = null;
                List<ReportDetailtSTMarket> listDataLastMonth = null;
                List<ReportDetailtSTMarket> listDataLastYear = null;

                foreach (string item in listAsianItem)
                {
                    listDataYear = listDataCompareMonth.Where(x => x.MarketName == item && x.Month == month.ToString() && x.Year == year.ToString()).ToList();
                    listDataLastMonth = listDataCompareMonth.Where(x => x.MarketName == item && x.Month == (month - 1).ToString() && x.Year == year.ToString()).ToList();
                    // Trường hợp tháng 1
                    if (month == 1)
                    {
                        listDataLastMonth = listDataCompareMonth.Where(x => x.Month == "12" && x.Year == (year - 1).ToString()).ToList();
                    }
                    listDataLastYear = listDataCompareMonth.Where(x => x.MarketName == item && x.Month == month.ToString() && x.Year == (year - 1).ToString()).ToList();

                    if (listDataYear == null)
                    {
                        listDataYear = new List<ReportDetailtSTMarket>();
                    }

                    if (listDataLastMonth == null)
                    {
                        listDataLastMonth = new List<ReportDetailtSTMarket>();
                    }

                    if (listDataLastYear == null)
                    {
                        listDataLastYear = new List<ReportDetailtSTMarket>();
                    }

                    // Year
                    double sumDSChiQuayYear = listDataYear.Sum(x => x.DSChiQuay);
                    double sumDSChiNhaYear = listDataYear.Sum(x => x.DSChiNha);
                    double sumDSCKYear = listDataYear.Sum(x => x.DSCK);
                    double sumTongDSYear = listDataYear.Sum(x => x.TongDS);

                    // Last Month
                    double sumDSChiQuayLastMonth = listDataLastMonth.Sum(x => x.DSChiQuay);
                    double sumDSChiNhaLastMonth = listDataLastMonth.Sum(x => x.DSChiNha);
                    double sumDSCKLastMonth = listDataLastMonth.Sum(x => x.DSCK);
                    double sumTongDSLastMonth = listDataLastMonth.Sum(x => x.TongDS);

                    // LastYear
                    double sumDSChiQuayLastYear = listDataLastYear.Sum(x => x.DSChiQuay);
                    double sumDSChiNhaLastYear = listDataLastYear.Sum(x => x.DSChiNha);
                    double sumDSCKLastYear = listDataLastYear.Sum(x => x.DSCK);
                    double sumTongDSLastYear = listDataLastYear.Sum(x => x.TongDS);

                    // Tạo mảng insert dữ liệu để vẽ biểu đồ cột

                    // Tháng hiện tại

                    arrayGradation[count] = new GradationCompare()
                    {
                        NameGradationCompare = string.Format("DSCK {0}/{1}", month, year),
                        amount = sumDSCKYear,
                        NameType = item
                    };
                    count++;

                    // Tháng trước
                    arrayGradation[count] = new GradationCompare()
                    {
                        NameGradationCompare = string.Format("DSCK {0}/{1}", month - 1, year),
                        amount = sumDSCKLastMonth,
                        NameType = item
                    };
                    count++;

                    // Cùng kì năm trước

                    arrayGradation[count] = new GradationCompare()
                    {
                        NameGradationCompare = string.Format("DSCK {0}/{1}", month, year - 1),
                        amount = sumDSCKLastYear,
                        NameType = item
                    };
                    count++;
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

        ///// <summary>
        ///// Get data cho báo cáo so sánh giai đoạn
        ///// </summary>
        ///// <returns></returns>
        ///// <history>
        /////     [Truong Lam]   Created [10/06/2020]
        ///// </history>
        //[HttpPost]
        //public ActionResult ReportDetailtCompareMonthForOne([DataSourceRequest]DataSourceRequest request, string reportTypeID)
        //{
        //    // Danh sach của data gradation gồm key và value
        //    int toYear = DateTime.Now.Year;
        //    int toMonth = DateTime.Now.Month;
        //    string marketID = "001";

        //    List<ReportDetailtServiceType> listDataCompareMonth = new ReportBL().ReportDetailtCompareMonthForOne(toYear, toMonth, reportTypeID, marketID);

        //    foreach (ReportDetailtServiceType item in listDataCompareMonth)
        //    {
        //        item.TongDS = item.DSChiQuay + item.DSChiNha + item.DSCK;
        //    }

        //    // Danh sách mã thị trường của Tất cả
        //    List<string> listMarket = new List<string>() { "003", "005", "001", "002", "014", "004" };

        //    // Khởi tạo datatable
        //    DataTable table = new DataTable();
        //    // Tạo các cột cho datatable
        //    // tháng hiện tại
        //    table.Columns.Add("PartnerName", typeof(String));
        //    table.Columns.Add("CQ1", typeof(double));
        //    table.Columns.Add("CQ2", typeof(double));
        //    table.Columns.Add("CQ3", typeof(double));

        //    table.Columns.Add("CN1", typeof(double));
        //    table.Columns.Add("CN2", typeof(double));
        //    table.Columns.Add("CN3", typeof(double));

        //    table.Columns.Add("CK1", typeof(double));
        //    table.Columns.Add("CK2", typeof(double));
        //    table.Columns.Add("CK3", typeof(double));

        //    table.Columns.Add("TDS1", typeof(double));
        //    table.Columns.Add("TDS2", typeof(double));
        //    table.Columns.Add("TDS3", typeof(double));


        //    table.Columns.Add("ReportID", typeof(String));

        //    List<string> listTemp = new List<string>();

        //    foreach (ReportDetailtServiceType item in listDataCompareMonth)
        //    {
        //        // Cùng kì
        //        ReportDetailtServiceType dataItemLastYear = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == toMonth.ToString() && x.Year == (toYear - 1).ToString());
        //        ReportDetailtServiceType dataItemYear = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == toMonth.ToString() && x.Year == toYear.ToString());
        //        ReportDetailtServiceType dataItemLastMonth = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == (toMonth - 1).ToString() && x.Year == toYear.ToString());

        //        if (!listTemp.Contains(item.PartnerCode))
        //        {
        //            // Trường hợp năm trước có đối tác không có
        //            if (dataItemLastYear == null)
        //            {
        //                dataItemLastYear = new ReportDetailtServiceType();
        //                dataItemLastYear.PartnerName = item.PartnerName;
        //                dataItemLastYear.DSChiQuay = 0;
        //                dataItemLastYear.DSChiNha = 0;
        //                dataItemLastYear.DSCK = 0;
        //                dataItemLastYear.Year = item.Year;
        //                dataItemLastYear.Month = item.Month;
        //            }

        //            // Trường hợp năm có đối tác không có
        //            if (dataItemYear == null)
        //            {
        //                dataItemYear = new ReportDetailtServiceType();
        //                dataItemYear.PartnerName = item.PartnerName;
        //                dataItemYear.DSChiQuay = 0;
        //                dataItemYear.DSChiNha = 0;
        //                dataItemYear.DSCK = 0;
        //                dataItemYear.Year = item.Year;
        //                dataItemYear.Month = item.Month;
        //            }

        //            // Trường hợp năm có tháng trước có đối tác không có
        //            if (dataItemLastMonth == null)
        //            {
        //                dataItemLastMonth = new ReportDetailtServiceType();
        //                dataItemLastMonth.PartnerName = item.PartnerName;
        //                dataItemLastMonth.DSChiQuay = 0;
        //                dataItemLastMonth.DSChiNha = 0;
        //                dataItemLastMonth.DSCK = 0;
        //                dataItemLastMonth.Year = item.Year;
        //                dataItemLastMonth.Month = item.Month;
        //            }

        //            // Check tồn tại của item
        //            string value = string.Format("PartnerName='{0}'", item.PartnerName);
        //            DataRow[] foundRows = table.Select(value);

        //            if (dataItemLastYear != null && dataItemYear != null && dataItemLastMonth != null && foundRows.Count() == 0)
        //            {
        //                // add item vào table
        //                table.Rows.Add(dataItemYear.PartnerName
        //                    , dataItemYear.DSChiQuay, dataItemLastMonth.DSChiQuay, dataItemLastYear.DSChiQuay
        //                    , dataItemYear.DSChiNha, dataItemLastMonth.DSChiNha, dataItemLastYear.DSChiNha
        //                    , dataItemYear.DSCK, dataItemLastMonth.DSCK, dataItemLastYear.DSCK
        //                    , dataItemYear.TongDS, dataItemLastMonth.TongDS, dataItemLastYear.TongDS);
        //            }
        //        }
        //    }

        //    return Json(table.ToDataSourceResult(request), JsonRequestBehavior.AllowGet);
        //}

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

            // List Market
            List<string> listMarket = new List<string>();
            foreach (ReportDetailtServiceType item in listDataCompareMonth)
            {
                if(!listMarket.Contains(item.MarketName))
                {
                    listMarket.Add(item.MarketName);
                }
                item.ReportID = item.PartnerName;
                item.TongDS = item.DSChiQuay + item.DSChiNha + item.DSCK;
            }

            if (marketID.Contains("005"))
            {
                List<ReportDetailtServiceType> listDataCompareMonthConvert = new List<ReportDetailtServiceType>();

                foreach(string item in listMarket)
                {
                    // Month
                    List<ReportDetailtServiceType> listDataItemMonth = listDataCompareMonth.Where(x => x.MarketName == item && x.Month == month.ToString() && x.Year == year.ToString()).ToList();
                    // Last month
                    List<ReportDetailtServiceType> listDataItemLastMonth = listDataCompareMonth.Where(x => x.MarketName == item && x.Month == (month - 1).ToString() && x.Year == year.ToString()).ToList();

                    // trường hợp month = 1
                    if (month == 1)
                    {
                        listDataItemLastMonth = listDataCompareMonth.Where(x => x.MarketName == item && x.Month == "12" && x.Year == (year - 1).ToString()).ToList();
                    }
                    List<ReportDetailtServiceType> listDataItemMonthLastYear = listDataCompareMonth.Where(x => x.MarketName == item && x.Month == month.ToString() && x.Year == (year -1).ToString()).ToList();

                    if (listDataItemMonth.Count == 0)
                    {
                        listDataCompareMonthConvert.Add(
                            new ReportDetailtServiceType()
                            {
                                MarketName = "Thị trường Châu Á",
                                PartnerName = item,
                                Month = month.ToString(),
                                Year = year.ToString()
                            }
                        );
                    }

                    if (listDataItemLastMonth.Count == 0)
                    {
                        if (month == 1)
                        {
                            listDataCompareMonthConvert.Add(
                                new ReportDetailtServiceType()
                                {
                                    MarketName = "Thị trường Châu Á",
                                    PartnerName = item,
                                    Month = "12",
                                    Year = (year -1).ToString()
                                }
                            );
                        }
                        else
                        {
                            listDataCompareMonthConvert.Add(
                                new ReportDetailtServiceType()
                                {
                                    MarketName = "Thị trường Châu Á",
                                    PartnerName = item,
                                    Month = (month - 1).ToString(),
                                    Year = year.ToString()
                                }
                            );
                        }
                    }

                    if (listDataItemMonthLastYear.Count == 0)
                    {
                        listDataCompareMonthConvert.Add(
                            new ReportDetailtServiceType()
                            {
                                MarketName = "Thị trường Châu Á",
                                PartnerName = item,
                                Month = month.ToString(),
                                Year = (year - 1).ToString()
                            }
                        );
                    }

                    // Add month
                    listDataCompareMonthConvert.Add(
                        new ReportDetailtServiceType()
                        {
                            MarketName = "Thị trường Châu Á",
                            PartnerName = item,
                            DSChiQuay = listDataItemMonth.Sum(x => x.DSChiQuay),
                            DSChiNha = listDataItemMonth.Sum(x => x.DSChiNha),
                            DSCK = listDataItemMonth.Sum(x => x.DSCK),
                            TongDS = listDataItemMonth.Sum(x => x.TongDS),
                            Month = month.ToString(),
                            Year = year.ToString()
                        }
                    );

                    // Add last month
                    if (month == 1)
                    {
                        listDataCompareMonthConvert.Add(
                            new ReportDetailtServiceType()
                            {
                                MarketName = "Thị trường Châu Á",
                                PartnerName = item,
                                DSChiQuay = listDataItemLastMonth.Sum(x => x.DSChiQuay),
                                DSChiNha = listDataItemLastMonth.Sum(x => x.DSChiNha),
                                DSCK = listDataItemLastMonth.Sum(x => x.DSCK),
                                TongDS = listDataItemLastMonth.Sum(x => x.TongDS),
                                Month = "12",
                                Year = (year - 1).ToString()
                            }
                        );
                    }
                    else
                    {
                        listDataCompareMonthConvert.Add(
                            new ReportDetailtServiceType()
                            {
                                MarketName = "Thị trường Châu Á",
                                PartnerName = item,
                                DSChiQuay = listDataItemLastMonth.Sum(x => x.DSChiQuay),
                                DSChiNha = listDataItemLastMonth.Sum(x => x.DSChiNha),
                                DSCK = listDataItemLastMonth.Sum(x => x.DSCK),
                                TongDS = listDataItemLastMonth.Sum(x => x.TongDS),
                                Month = (month - 1).ToString(),
                                Year = year.ToString()
                            }
                        );
                    }

                    // Add month last year
                    listDataCompareMonthConvert.Add(
                        new ReportDetailtServiceType()
                        {
                            MarketName = "Thị trường Châu Á",
                            PartnerName = item,
                            DSChiQuay = listDataItemMonthLastYear.Sum(x => x.DSChiQuay),
                            DSChiNha = listDataItemMonthLastYear.Sum(x => x.DSChiNha),
                            DSCK = listDataItemMonthLastYear.Sum(x => x.DSCK),
                            TongDS = listDataItemMonthLastYear.Sum(x => x.TongDS),
                            Month = month.ToString(),
                            Year = (year - 1).ToString()
                        }
                    );
                }

                if (listDataCompareMonthConvert.Count > 0)
                {
                    listDataCompareMonth = new List<ReportDetailtServiceType>(listDataCompareMonthConvert);
                }
            }

            // Khởi tạo datatable
            DataTable table = new DataTable();
            // Tạo các cột cho datatable
            // tháng hiện tại
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

            table.Columns.Add("MarketName", typeof(String));


            List<string> listTemp = new List<string>();

            foreach (ReportDetailtServiceType item in listDataCompareMonth)
            {
                // Cùng kì
                ReportDetailtServiceType dataItemLastYear = listDataCompareMonth.Find(x => x.PartnerName == item.PartnerName && x.Month == month.ToString() && x.Year == (year - 1).ToString());
                ReportDetailtServiceType dataItemYear = listDataCompareMonth.Find(x => x.PartnerName == item.PartnerName && x.Month == month.ToString() && x.Year == year.ToString());
                ReportDetailtServiceType dataItemLastMonth = listDataCompareMonth.Find(x => x.PartnerName == item.PartnerName && x.Month == (month - 1).ToString() && x.Year == year.ToString());
                // Trường hợp tháng 1
                if (month == 1)
                {
                    dataItemLastMonth = listDataCompareMonth.Find(x => x.PartnerName == item.PartnerName && x.Month == "12" && x.Year == (year - 1).ToString());
                }

                // Trường hợp năm trước có đối tác không có
                if (dataItemLastYear == null)
                {
                    dataItemLastYear = new ReportDetailtServiceType();
                    dataItemLastYear.PartnerName = item.PartnerName;
                    dataItemLastYear.MarketName = item.MarketName;
                    dataItemLastYear.DSChiQuay = 0;
                    dataItemLastYear.DSChiNha = 0;
                    dataItemLastYear.DSCK = 0;
                    dataItemLastYear.Year = (year - 1).ToString();
                    dataItemLastYear.Month = month.ToString();
                }

                // Trường hợp năm có đối tác không có
                if (dataItemYear == null)
                {
                    dataItemYear = new ReportDetailtServiceType();
                    dataItemYear.PartnerName = item.PartnerName;
                    dataItemYear.MarketName = item.MarketName;
                    dataItemYear.DSChiQuay = 0;
                    dataItemYear.DSChiNha = 0;
                    dataItemYear.DSCK = 0;
                    dataItemYear.Year = year.ToString();
                    dataItemYear.Month = month.ToString();
                }

                // Trường hợp năm có tháng trước có đối tác không có
                if (dataItemLastMonth == null)
                {
                    dataItemLastMonth = new ReportDetailtServiceType();
                    dataItemLastMonth.PartnerName = item.PartnerName;
                    dataItemLastMonth.MarketName = item.MarketName;
                    dataItemLastMonth.DSChiQuay = 0;
                    dataItemLastMonth.DSChiNha = 0;
                    dataItemLastMonth.DSCK = 0;

                    if (month == 1)
                    {
                        dataItemLastMonth.Year = (year - 1).ToString();
                        dataItemLastMonth.Month = "12";
                    }
                    else
                    {
                        dataItemLastMonth.Year = (month - 1).ToString();
                        dataItemLastMonth.Month = year.ToString();
                    }
                }

                // Check tồn tại của item
                string value = string.Format("PartnerName='{0}'", item.PartnerName);
                DataRow[] foundRows = table.Select(value);

                if (dataItemLastYear != null && dataItemYear != null && dataItemLastMonth != null && foundRows.Count() == 0)
                {
                    // add item vào table
                    table.Rows.Add(item.PartnerName
                        , dataItemYear.DSChiQuay, dataItemLastMonth.DSChiQuay, dataItemLastYear.DSChiQuay
                        , dataItemYear.DSChiNha, dataItemLastMonth.DSChiNha, dataItemLastYear.DSChiNha
                        , dataItemYear.DSCK, dataItemLastMonth.DSCK, dataItemLastYear.DSCK
                        , dataItemYear.TongDS, dataItemLastMonth.TongDS, dataItemLastYear.TongDS
                        , item.MarketName);
                }

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
        public ActionResult SearchReportDetailtCompareMonthForOneCompare([DataSourceRequest]DataSourceRequest request, int month, int year, string reportTypeID, string marketID)
        {
            List<ReportDetailtServiceType> listDataCompareMonth = new ReportBL().ReportDetailtCompareMonthForOne(year, month, reportTypeID, marketID);

            // List Market
            List<string> listMarket = new List<string>();
            foreach (ReportDetailtServiceType item in listDataCompareMonth)
            {
                if (!listMarket.Contains(item.MarketName))
                {
                    listMarket.Add(item.MarketName);
                }
                item.ReportID = item.PartnerName;
                item.TongDS = item.DSChiQuay + item.DSChiNha + item.DSCK;
            }

            if (marketID.Contains("005"))
            {
                List<ReportDetailtServiceType> listDataCompareMonthConvert = new List<ReportDetailtServiceType>();

                foreach (string item in listMarket)
                {
                    // Month
                    List<ReportDetailtServiceType> listDataItemMonth = listDataCompareMonth.Where(x => x.MarketName == item && x.Month == month.ToString() && x.Year == year.ToString()).ToList();
                    // Last month
                    List<ReportDetailtServiceType> listDataItemLastMonth = listDataCompareMonth.Where(x => x.MarketName == item && x.Month == (month - 1).ToString() && x.Year == year.ToString()).ToList();

                    // trường hợp month = 1
                    if (month == 1)
                    {
                        listDataItemLastMonth = listDataCompareMonth.Where(x => x.MarketName == item && x.Month == "12" && x.Year == (year - 1).ToString()).ToList();
                    }
                    List<ReportDetailtServiceType> listDataItemMonthLastYear = listDataCompareMonth.Where(x => x.MarketName == item && x.Month == month.ToString() && x.Year == (year - 1).ToString()).ToList();

                    if (listDataItemMonth.Count == 0)
                    {
                        listDataCompareMonthConvert.Add(
                            new ReportDetailtServiceType()
                            {
                                MarketName = "Thị trường Châu Á",
                                PartnerName = item,
                                Month = month.ToString(),
                                Year = year.ToString()
                            }
                        );
                    }

                    if (listDataItemLastMonth.Count == 0)
                    {
                        if (month == 1)
                        {
                            listDataCompareMonthConvert.Add(
                                new ReportDetailtServiceType()
                                {
                                    MarketName = "Thị trường Châu Á",
                                    PartnerName = item,
                                    Month = "12",
                                    Year = (year - 1).ToString()
                                }
                            );
                        }
                        else
                        {
                            listDataCompareMonthConvert.Add(
                                new ReportDetailtServiceType()
                                {
                                    MarketName = "Thị trường Châu Á",
                                    PartnerName = item,
                                    Month = (month - 1).ToString(),
                                    Year = year.ToString()
                                }
                            );
                        }
                    }

                    if (listDataItemMonthLastYear.Count == 0)
                    {
                        listDataCompareMonthConvert.Add(
                            new ReportDetailtServiceType()
                            {
                                MarketName = "Thị trường Châu Á",
                                PartnerName = item,
                                Month = month.ToString(),
                                Year = (year - 1).ToString()
                            }
                        );
                    }

                    // Add month
                    listDataCompareMonthConvert.Add(
                        new ReportDetailtServiceType()
                        {
                            MarketName = "Thị trường Châu Á",
                            PartnerName = item,
                            DSChiQuay = listDataItemMonth.Sum(x => x.DSChiQuay),
                            DSChiNha = listDataItemMonth.Sum(x => x.DSChiNha),
                            DSCK = listDataItemMonth.Sum(x => x.DSCK),
                            TongDS = listDataItemMonth.Sum(x => x.TongDS),
                            Month = month.ToString(),
                            Year = year.ToString()
                        }
                    );

                    // Add last month
                    if (month == 1)
                    {
                        listDataCompareMonthConvert.Add(
                            new ReportDetailtServiceType()
                            {
                                MarketName = "Thị trường Châu Á",
                                PartnerName = item,
                                DSChiQuay = listDataItemLastMonth.Sum(x => x.DSChiQuay),
                                DSChiNha = listDataItemLastMonth.Sum(x => x.DSChiNha),
                                DSCK = listDataItemLastMonth.Sum(x => x.DSCK),
                                TongDS = listDataItemLastMonth.Sum(x => x.TongDS),
                                Month = "12",
                                Year = (year - 1).ToString()
                            }
                        );
                    }
                    else
                    {
                        listDataCompareMonthConvert.Add(
                            new ReportDetailtServiceType()
                            {
                                MarketName = "Thị trường Châu Á",
                                PartnerName = item,
                                DSChiQuay = listDataItemLastMonth.Sum(x => x.DSChiQuay),
                                DSChiNha = listDataItemLastMonth.Sum(x => x.DSChiNha),
                                DSCK = listDataItemLastMonth.Sum(x => x.DSCK),
                                TongDS = listDataItemLastMonth.Sum(x => x.TongDS),
                                Month = (month - 1).ToString(),
                                Year = year.ToString()
                            }
                        );
                    }

                    // Add month last year
                    listDataCompareMonthConvert.Add(
                        new ReportDetailtServiceType()
                        {
                            MarketName = "Thị trường Châu Á",
                            PartnerName = item,
                            DSChiQuay = listDataItemMonthLastYear.Sum(x => x.DSChiQuay),
                            DSChiNha = listDataItemMonthLastYear.Sum(x => x.DSChiNha),
                            DSCK = listDataItemMonthLastYear.Sum(x => x.DSCK),
                            TongDS = listDataItemMonthLastYear.Sum(x => x.TongDS),
                            Month = month.ToString(),
                            Year = (year - 1).ToString()
                        }
                    );
                }

                if (listDataCompareMonthConvert.Count > 0)
                {
                    listDataCompareMonth = new List<ReportDetailtServiceType>(listDataCompareMonthConvert);
                }
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

            table.Columns.Add("MarketName", typeof(String));

            List<string> listTemp = new List<string>();

            foreach (ReportDetailtServiceType item in listDataCompareMonth)
            {
                // Cùng kì
                ReportDetailtServiceType dataItemLastYear = listDataCompareMonth.Find(x => x.PartnerName == item.PartnerName && x.Month == month.ToString() && x.Year == (year - 1).ToString());
                ReportDetailtServiceType dataItemYear = listDataCompareMonth.Find(x => x.PartnerName == item.PartnerName && x.Month == month.ToString() && x.Year == year.ToString());
                ReportDetailtServiceType dataItemLastMonth = listDataCompareMonth.Find(x => x.PartnerName == item.PartnerName && x.Month == (month - 1).ToString() && x.Year == year.ToString());
                // Trường hợp tháng 1
                if (month == 1)
                {
                    dataItemLastMonth = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == "12" && x.Year == (year - 1).ToString());
                }

                // Trường hợp năm trước có đối tác không có
                if (dataItemLastYear == null)
                {
                    dataItemLastYear = new ReportDetailtServiceType();
                    dataItemLastYear.PartnerName = item.PartnerName;
                    dataItemLastYear.MarketName = item.MarketName;
                    dataItemLastYear.DSChiQuay = 0;
                    dataItemLastYear.DSChiNha = 0;
                    dataItemLastYear.DSCK = 0;
                    dataItemLastYear.Year = (year - 1).ToString();
                    dataItemLastYear.Month = month.ToString();
                }

                // Trường hợp năm có đối tác không có
                if (dataItemYear == null)
                {
                    dataItemYear = new ReportDetailtServiceType();
                    dataItemYear.PartnerName = item.PartnerName;
                    dataItemYear.MarketName = item.MarketName;
                    dataItemYear.DSChiQuay = 0;
                    dataItemYear.DSChiNha = 0;
                    dataItemYear.DSCK = 0;
                    dataItemYear.Year = year.ToString();
                    dataItemYear.Month = month.ToString();
                }

                // Trường hợp năm có tháng trước có đối tác không có
                if (dataItemLastMonth == null)
                {
                    dataItemLastMonth = new ReportDetailtServiceType();
                    dataItemLastMonth.PartnerName = item.PartnerName;
                    dataItemLastMonth.MarketName = item.MarketName;
                    dataItemLastMonth.DSChiQuay = 0;
                    dataItemLastMonth.DSChiNha = 0;
                    dataItemLastMonth.DSCK = 0;

                    if (month == 1)
                    {
                        dataItemLastMonth.Year = (year - 1).ToString();
                        dataItemLastMonth.Month = "12";
                    }
                    else
                    {
                        dataItemLastMonth.Year = (month - 1).ToString();
                        dataItemLastMonth.Month = year.ToString();
                    }
                }

                // Check tồn tại của item
                string value = string.Format("PartnerName='{0}'", item.PartnerName);
                DataRow[] foundRows = table.Select(value);

                if (dataItemLastYear != null && dataItemYear != null && dataItemLastMonth != null && foundRows.Count() == 0)
                {
                    // add item vào table
                    table.Rows.Add(item.PartnerName
                        , dataItemYear.DSChiQuay - dataItemLastMonth.DSChiQuay, dataItemYear.DSChiNha - dataItemLastMonth.DSChiNha, dataItemYear.DSCK - dataItemLastMonth.DSCK
                        , dataItemYear.TongDS - dataItemLastMonth.TongDS
                        , dataItemYear.DSChiQuay - dataItemLastYear.DSChiQuay, dataItemYear.DSChiNha - dataItemLastYear.DSChiNha, dataItemYear.DSCK - dataItemLastYear.DSCK
                        , dataItemYear.TongDS - dataItemLastYear.TongDS
                        , item.MarketName);
                }
            }
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
        public ActionResult SearchColumnChartCompareMonthStackForOne([DataSourceRequest]DataSourceRequest request, int month, int year, string reportTypeID, string marketID)
        {

            List<ReportDetailtServiceType> listDataCompareMonth = new ReportBL().ReportDetailtCompareMonthForOne(year, month, reportTypeID, marketID);
            List<ReportDetailtServiceType> listDataCompareMonthConvert = new List<ReportDetailtServiceType>();

            List<string> listTemp = new List<string>();

            // Trường hợp không phải là thị trường Châu á
            if (!marketID.Equals("005"))
            {
                listDataCompareMonthConvert = new List<ReportDetailtServiceType>(listDataCompareMonth);

                foreach (ReportDetailtServiceType item in listDataCompareMonth)
                {
                    item.TongDS = item.DSChiQuay + item.DSChiNha + item.DSCK;
                }

                foreach (ReportDetailtServiceType item in listDataCompareMonth)
                {
                    if (!listTemp.Contains(item.PartnerCode))
                    {
                        // Cùng kì
                        ReportDetailtServiceType dataItemLastYear = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == month.ToString() && x.Year == (year - 1).ToString());
                        ReportDetailtServiceType dataItemYear = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == month.ToString() && x.Year == year.ToString());
                        ReportDetailtServiceType dataItemLastMonth = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == (month - 1).ToString() && x.Year == year.ToString());
                        // Trường hợp tháng 1
                        if (month == 1)
                        {
                            dataItemLastMonth = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == "12" && x.Year == (year - 1).ToString());
                        }

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
                            listDataCompareMonthConvert.Add(dataItemLastYear);
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
                            listDataCompareMonthConvert.Add(dataItemYear);
                        }

                        // Trường hợp tháng trước không có
                        if (dataItemLastMonth == null)
                        {
                            dataItemLastMonth = new ReportDetailtServiceType();
                            dataItemLastMonth.PartnerName = item.PartnerName;
                            dataItemLastMonth.DSChiQuay = 0;
                            dataItemLastMonth.DSChiNha = 0;
                            dataItemLastMonth.DSCK = 0;
                            if (month == 1)
                            {
                                dataItemLastMonth.Year = (year - 1).ToString();
                                dataItemLastMonth.Month = "12";
                            }
                            else
                            {
                                dataItemLastMonth.Month = (month - 1).ToString();
                                dataItemLastMonth.Year = year.ToString();
                            }
                            listDataCompareMonthConvert.Add(dataItemLastMonth);
                        }

                        // Add partnerCode để kiểm tra
                        listTemp.Add(item.PartnerCode);
                    }
                }
            }
            else
            {
                foreach (ReportDetailtServiceType item in listDataCompareMonth)
                {
                    if (!listTemp.Contains(item.MarketName))
                    {
                        listTemp.Add(item.MarketName);
                    }

                    item.TongDS = item.DSChiQuay + item.DSChiNha + item.DSCK;
                }

                foreach (string item in listTemp)
                {
                    List<ReportDetailtServiceType> listDataItemLastYear = listDataCompareMonth.Where(x => x.MarketName == item && x.Month == month.ToString() && x.Year == (year - 1).ToString()).ToList();
                    List<ReportDetailtServiceType> listDataItemYear = listDataCompareMonth.Where(x => x.MarketName == item && x.Month == month.ToString() && x.Year == year.ToString()).ToList();
                    List<ReportDetailtServiceType> listDataItemLastMonth = listDataCompareMonth.Where(x => x.MarketName == item && x.Month == (month - 1).ToString() && x.Year == year.ToString()).ToList();
                    // Trường hợp tháng 1
                    if (month == 1)
                    {
                        listDataItemLastMonth = listDataCompareMonth.Where(x => x.MarketName == item && x.Month == "12" && x.Year == (year - 1).ToString()).ToList();
                    }
                    
                    double sumTongDSLastYear = listDataItemLastYear.Sum(x => x.TongDS);
                    double sumTongDSYear = listDataItemYear.Sum(x => x.TongDS);
                    double sumTongDSLastMonth = listDataItemLastMonth.Sum(x => x.TongDS);

                    listDataCompareMonthConvert.Add(

                        // Cùng kì năm trước
                        new ReportDetailtServiceType()
                        {
                            PartnerName = item,
                            Month = month.ToString(),
                            Year = (year - 1).ToString(),
                            TongDS = sumTongDSLastYear
                        }
                    );

                    if (month == 1)
                    {
                        listDataCompareMonthConvert.Add(
                            // Tháng trước
                            new ReportDetailtServiceType()
                            {
                                PartnerName = item,
                                Month = "12",
                                Year = (year - 1).ToString(),
                                TongDS = sumTongDSLastMonth
                            }
                        );
                    }
                    else
                    {
                        listDataCompareMonthConvert.Add(
                            // Tháng trước
                            new ReportDetailtServiceType()
                            {
                                PartnerName = item,
                                Month = (month - 1).ToString(),
                                Year = year.ToString(),
                                TongDS = sumTongDSLastMonth
                            }
                        );
                    }
                    
                    listDataCompareMonthConvert.Add(
                        // Tháng hiện tại
                        new ReportDetailtServiceType()
                        {
                            PartnerName = item,
                            Month = month.ToString(),
                            Year = year.ToString(),
                            TongDS = sumTongDSYear
                        }
                    );
                }
            }


            // Số mảng cần tạo
            int arrayCount = 1;
            GradationCharColumn[] arrayGradation = new GradationCharColumn[listDataCompareMonthConvert.Count * arrayCount];
            int count = 0;

            foreach (ReportDetailtServiceType item in listDataCompareMonthConvert)
            {
                // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                arrayGradation[count] = new GradationCharColumn()
                {
                    Serie = string.Format("Tổng tháng {0}/{1}", item.Month, item.Year),
                    Segmento = item.PartnerName,
                    Valor1 = item.TongDS
                };
                count++;
            }

            if (listDataCompareMonthConvert.Count == 0)
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
        /// Get data cho việc vẽ biểu đồ cột cho so sánh giai đoạn theo Doanh số Chi Quầy
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult SearchColumnsChartCompareMonthForOneDSChiQuay([DataSourceRequest]DataSourceRequest request, int month, int year, string reportTypeID, string marketID)
        {

            List<ReportDetailtServiceType> listDataCompareMonth = new ReportBL().ReportDetailtCompareMonthForOne(year, month, reportTypeID, marketID);
            List<ReportDetailtServiceType> listDataCompareMonthConvert = new List<ReportDetailtServiceType>();

            // Trường hợp không phải là thị trường Châu á
            if (!marketID.Equals("005"))
            {
                listDataCompareMonthConvert = new List<ReportDetailtServiceType>(listDataCompareMonth);

                // Khởi tạo list tạm
                List<string> listTemp = new List<string>();
                foreach (ReportDetailtServiceType item in listDataCompareMonth)
                {
                    if (listTemp.Contains(item.PartnerName))
                    {
                        continue;
                    }

                    ReportDetailtServiceType dataItemLastYear = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == month.ToString() && x.Year == (year - 1).ToString());
                    ReportDetailtServiceType dataItemYear = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == month.ToString() && x.Year == year.ToString());
                    ReportDetailtServiceType dataItemLastMonth = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == (month - 1).ToString() && x.Year == year.ToString());
                    // Trường hợp tháng 1
                    if (month == 1)
                    {
                        dataItemLastMonth = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == "12" && x.Year == (year - 1).ToString());
                    }

                    // Cùng kì năm trước
                    if (dataItemLastYear == null)
                    {
                        dataItemLastYear = new ReportDetailtServiceType();
                        dataItemLastYear.PartnerCode = item.PartnerCode;
                        dataItemLastYear.PartnerName = item.PartnerName;
                        dataItemLastYear.Month = month.ToString();
                        dataItemLastYear.Year = (year - 1).ToString();
                        dataItemLastYear.MarketCode = item.MarketCode;
                        dataItemLastYear.MarketName = item.MarketName;
                        // Add vào list mới
                        listDataCompareMonthConvert.Add(dataItemLastYear);
                    }

                    // Hiện tại
                    if (dataItemYear == null)
                    {
                        dataItemYear = new ReportDetailtServiceType();
                        dataItemYear.PartnerCode = item.PartnerCode;
                        dataItemYear.PartnerName = item.PartnerName;
                        dataItemYear.Month = month.ToString();
                        dataItemYear.Year = year.ToString();
                        dataItemYear.MarketCode = item.MarketCode;
                        dataItemYear.MarketName = item.MarketName;
                        // Add vào list mới
                        listDataCompareMonthConvert.Add(dataItemYear);
                    }

                    // tháng trước
                    if (dataItemLastMonth == null)
                    {
                        dataItemLastMonth = new ReportDetailtServiceType();
                        dataItemLastMonth.PartnerCode = item.PartnerCode;
                        dataItemLastMonth.PartnerName = item.PartnerName;
                        if (month == 1)
                        {
                            dataItemLastMonth.Year = (year - 1).ToString();
                            dataItemLastMonth.Month = "12";
                        }
                        else
                        {
                            dataItemLastMonth.Month = (month - 1).ToString();
                            dataItemLastMonth.Year = year.ToString();
                        }
                        dataItemLastMonth.MarketCode = item.MarketCode;
                        dataItemLastMonth.MarketName = item.MarketName;
                        // Add vào list mới
                        listDataCompareMonthConvert.Add(dataItemLastMonth);
                    }

                    listTemp.Add(item.PartnerName);
                }
            }
            else
            {
                // Danh sách các thị trường
                List<string> listMarket = new List<string>();
                // truyền các thị trường vào list tạm
                foreach (ReportDetailtServiceType item in listDataCompareMonth)
                {
                    if (listMarket.Contains(item.MarketName))
                    {
                        continue;
                    }
                    listMarket.Add(item.MarketName);
                }

                foreach (string item in listMarket)
                {
                    List<ReportDetailtServiceType> listDataItemLastYear = listDataCompareMonth.Where(x => x.MarketName == item && x.Month == month.ToString() && x.Year == (year - 1).ToString()).ToList();
                    List<ReportDetailtServiceType> listDataItemYear = listDataCompareMonth.Where(x => x.MarketName == item && x.Month == month.ToString() && x.Year == year.ToString()).ToList();
                    List<ReportDetailtServiceType> listDataItemLastMonth = listDataCompareMonth.Where(x => x.MarketName == item && x.Month == (month - 1).ToString() && x.Year == year.ToString()).ToList();
                    // Trường hợp tháng 1
                    if (month == 1)
                    {
                        listDataItemLastMonth = listDataCompareMonth.Where(x => x.MarketName == item && x.Month == "12" && x.Year == (year - 1).ToString()).ToList();
                    }

                    double sumDSChiQuayLastYear = listDataItemLastYear.Sum(x => x.DSChiQuay);
                    double sumDSChiQuayYear = listDataItemYear.Sum(x => x.DSChiQuay);
                    double sumDSChiQuayLastMonth = listDataItemLastMonth.Sum(x => x.DSChiQuay);

                    listDataCompareMonthConvert.Add(

                        // Cùng kì năm trước
                        new ReportDetailtServiceType()
                        {
                            PartnerName = item,
                            Month = month.ToString(),
                            Year = (year - 1).ToString(),
                            DSChiQuay = sumDSChiQuayLastYear
                        }
                    );

                    if (month == 1)
                    {
                        listDataCompareMonthConvert.Add(
                            // Tháng trước
                            new ReportDetailtServiceType()
                            {
                                PartnerName = item,
                                Month = "12",
                                Year = (year - 1).ToString(),
                                DSChiQuay = sumDSChiQuayLastMonth
                            }
                        );
                    }
                    else
                    {
                        listDataCompareMonthConvert.Add(
                            // Tháng trước
                            new ReportDetailtServiceType()
                            {
                                PartnerName = item,
                                Month = (month - 1).ToString(),
                                Year = year.ToString(),
                                DSChiQuay = sumDSChiQuayLastMonth
                            }
                        );
                    }
                    

                    listDataCompareMonthConvert.Add(
                        // Tháng hiện tại
                        new ReportDetailtServiceType()
                        {
                            PartnerName = item,
                            Month = month.ToString(),
                            Year = year.ToString(),
                            DSChiQuay = sumDSChiQuayYear
                        }
                    );
                }
            }

            // Số record của mảng
            int countArray = 1;
            GradationCompare[] arrayGradation = new GradationCompare[countArray * listDataCompareMonthConvert.Count];
            int count = 0;
            foreach (ReportDetailtServiceType item in listDataCompareMonthConvert)
            {
                // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                arrayGradation[count] = new GradationCompare()
                {
                    NameGradationCompare = string.Format("Chi Quầy {0}/{1}", item.Month, item.Year),
                    amount = item.DSChiQuay,
                    NameType = item.PartnerName
                };

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
        /// Get data cho việc vẽ biểu đồ cột cho so sánh giai đoạn theo Doanh số Chi Nhà
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult SearchColumnsChartCompareMonthForOneDSChiNha([DataSourceRequest]DataSourceRequest request, int month, int year, string reportTypeID, string marketID)
        {

            List<ReportDetailtServiceType> listDataCompareMonth = new ReportBL().ReportDetailtCompareMonthForOne(year, month, reportTypeID, marketID);
            List<ReportDetailtServiceType> listDataCompareMonthConvert = new List<ReportDetailtServiceType>();

            // Trường hợp không phải là thị trường Châu á
            if (!marketID.Equals("005"))
            {
                listDataCompareMonthConvert = new List<ReportDetailtServiceType>(listDataCompareMonth);

                // Khởi tạo list tạm
                List<string> listTemp = new List<string>();
                foreach (ReportDetailtServiceType item in listDataCompareMonth)
                {
                    if (listTemp.Contains(item.PartnerName))
                    {
                        continue;
                    }

                    ReportDetailtServiceType dataItemLastYear = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == month.ToString() && x.Year == (year - 1).ToString());
                    ReportDetailtServiceType dataItemYear = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == month.ToString() && x.Year == year.ToString());
                    ReportDetailtServiceType dataItemLastMonth = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == (month - 1).ToString() && x.Year == year.ToString());
                    // Trường hợp tháng 1
                    if (month == 1)
                    {
                        dataItemLastMonth = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == "12" && x.Year == (year - 1).ToString());
                    }

                    // Cùng kì năm trước
                    if (dataItemLastYear == null)
                    {
                        dataItemLastYear = new ReportDetailtServiceType();
                        dataItemLastYear.PartnerCode = item.PartnerCode;
                        dataItemLastYear.PartnerName = item.PartnerName;
                        dataItemLastYear.Month = month.ToString();
                        dataItemLastYear.Year = (year - 1).ToString();
                        dataItemLastYear.MarketCode = item.MarketCode;
                        dataItemLastYear.MarketName = item.MarketName;
                        // Add vào list mới
                        listDataCompareMonthConvert.Add(dataItemLastYear);
                    }

                    // Hiện tại
                    if (dataItemYear == null)
                    {
                        dataItemYear = new ReportDetailtServiceType();
                        dataItemYear.PartnerCode = item.PartnerCode;
                        dataItemYear.PartnerName = item.PartnerName;
                        dataItemYear.Month = month.ToString();
                        dataItemYear.Year = year.ToString();
                        dataItemYear.MarketCode = item.MarketCode;
                        dataItemYear.MarketName = item.MarketName;
                        // Add vào list mới
                        listDataCompareMonthConvert.Add(dataItemYear);
                    }

                    // tháng trước
                    if (dataItemLastMonth == null)
                    {
                        dataItemLastMonth = new ReportDetailtServiceType();
                        dataItemLastMonth.PartnerCode = item.PartnerCode;
                        dataItemLastMonth.PartnerName = item.PartnerName;
                        if (month == 1)
                        {
                            dataItemLastMonth.Year = (year - 1).ToString();
                            dataItemLastMonth.Month = "12";
                        }
                        else
                        {
                            dataItemLastMonth.Month = (month - 1).ToString();
                            dataItemLastMonth.Year = year.ToString();
                        }
                        dataItemLastMonth.MarketCode = item.MarketCode;
                        dataItemLastMonth.MarketName = item.MarketName;
                        // Add vào list mới
                        listDataCompareMonthConvert.Add(dataItemLastMonth);
                    }

                    listTemp.Add(item.PartnerName);
                }
            }
            else
            {
                // Danh sách các thị trường
                List<string> listMarket = new List<string>();
                // truyền các thị trường vào list tạm
                foreach (ReportDetailtServiceType item in listDataCompareMonth)
                {
                    if (listMarket.Contains(item.MarketName))
                    {
                        continue;
                    }
                    listMarket.Add(item.MarketName);
                }

                foreach (string item in listMarket)
                {
                    List<ReportDetailtServiceType> listDataItemLastYear = listDataCompareMonth.Where(x => x.MarketName == item && x.Month == month.ToString() && x.Year == (year - 1).ToString()).ToList();
                    List<ReportDetailtServiceType> listDataItemYear = listDataCompareMonth.Where(x => x.MarketName == item && x.Month == month.ToString() && x.Year == year.ToString()).ToList();
                    List<ReportDetailtServiceType> listDataItemLastMonth = listDataCompareMonth.Where(x => x.MarketName == item && x.Month == (month - 1).ToString() && x.Year == year.ToString()).ToList();
                    // Trường hợp tháng 1
                    if (month == 1)
                    {
                        listDataItemLastMonth = listDataCompareMonth.Where(x => x.MarketName == item && x.Month == "12" && x.Year == (year - 1).ToString()).ToList();
                    }

                    double sumDSChiNhaLastYear = listDataItemLastYear.Sum(x => x.DSChiNha);
                    double sumDSChiNhaYear = listDataItemYear.Sum(x => x.DSChiNha);
                    double sumDSChiNhaLastMonth = listDataItemLastMonth.Sum(x => x.DSChiNha);

                    listDataCompareMonthConvert.Add(

                        // Cùng kì năm trước
                        new ReportDetailtServiceType()
                        {
                            PartnerName = item,
                            Month = month.ToString(),
                            Year = (year - 1).ToString(),
                            DSChiNha = sumDSChiNhaLastYear
                        }
                    );

                    if (month == 1)
                    {
                        listDataCompareMonthConvert.Add(
                            // Tháng trước
                            new ReportDetailtServiceType()
                            {
                                PartnerName = item,
                                Month = "12",
                                Year = (year - 1).ToString(),
                                DSChiNha = sumDSChiNhaLastMonth
                            }
                        );
                    }
                    else
                    {
                        listDataCompareMonthConvert.Add(
                            // Tháng trước
                            new ReportDetailtServiceType()
                            {
                                PartnerName = item,
                                Month = (month - 1).ToString(),
                                Year = year.ToString(),
                                DSChiNha = sumDSChiNhaLastMonth
                            }
                        );
                    }

                    listDataCompareMonthConvert.Add(
                        // Tháng hiện tại
                        new ReportDetailtServiceType()
                        {
                            PartnerName = item,
                            Month = month.ToString(),
                            Year = year.ToString(),
                            DSChiNha = sumDSChiNhaYear
                        }
                    );
                }
            }

            // Số record của mảng
            int countArray = 1;
            GradationCompare[] arrayGradation = new GradationCompare[countArray * listDataCompareMonthConvert.Count];
            int count = 0;
            foreach (ReportDetailtServiceType item in listDataCompareMonthConvert)
            {
                // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                arrayGradation[count] = new GradationCompare()
                {
                    NameGradationCompare = string.Format("Chi Nhà {0}/{1}", item.Month, item.Year),
                    amount = item.DSChiNha,
                    NameType = item.PartnerName
                };

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
        /// Get data cho việc vẽ biểu đồ cột cho so sánh giai đoạn theo Doanh số Chi Nhà
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult SearchColumnsChartCompareMonthForOneDSCK([DataSourceRequest]DataSourceRequest request, int month, int year, string reportTypeID, string marketID)
        {
            List<ReportDetailtServiceType> listDataCompareMonth = new ReportBL().ReportDetailtCompareMonthForOne(year, month, reportTypeID, marketID);
            List<ReportDetailtServiceType> listDataCompareMonthConvert = new List<ReportDetailtServiceType>();

            // Trường hợp không phải là thị trường Châu á
            if (!marketID.Equals("005"))
            {
                listDataCompareMonthConvert = new List<ReportDetailtServiceType>(listDataCompareMonth);

                // Khởi tạo list tạm
                List<string> listTemp = new List<string>();
                foreach (ReportDetailtServiceType item in listDataCompareMonth)
                {
                    if (listTemp.Contains(item.PartnerName))
                    {
                        continue;
                    }

                    ReportDetailtServiceType dataItemLastYear = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == month.ToString() && x.Year == (year - 1).ToString());
                    ReportDetailtServiceType dataItemYear = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == month.ToString() && x.Year == year.ToString());
                    ReportDetailtServiceType dataItemLastMonth = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == (month - 1).ToString() && x.Year == year.ToString());
                    // Trường hợp tháng 1
                    if (month == 1)
                    {
                        dataItemLastMonth = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == "12" && x.Year == (year - 1).ToString());
                    }

                    // Cùng kì năm trước
                    if (dataItemLastYear == null)
                    {
                        dataItemLastYear = new ReportDetailtServiceType();
                        dataItemLastYear.PartnerCode = item.PartnerCode;
                        dataItemLastYear.PartnerName = item.PartnerName;
                        dataItemLastYear.Month = month.ToString();
                        dataItemLastYear.Year = (year - 1).ToString();
                        dataItemLastYear.MarketCode = item.MarketCode;
                        dataItemLastYear.MarketName = item.MarketName;
                        // Add vào list mới
                        listDataCompareMonthConvert.Add(dataItemLastYear);
                    }

                    // Hiện tại
                    if (dataItemYear == null)
                    {
                        dataItemYear = new ReportDetailtServiceType();
                        dataItemYear.PartnerCode = item.PartnerCode;
                        dataItemYear.PartnerName = item.PartnerName;
                        dataItemYear.Month = month.ToString();
                        dataItemYear.Year = year.ToString();
                        dataItemYear.MarketCode = item.MarketCode;
                        dataItemYear.MarketName = item.MarketName;
                        // Add vào list mới
                        listDataCompareMonthConvert.Add(dataItemYear);
                    }

                    // tháng trước
                    if (dataItemLastMonth == null)
                    {
                        dataItemLastMonth = new ReportDetailtServiceType();
                        dataItemLastMonth.PartnerCode = item.PartnerCode;
                        dataItemLastMonth.PartnerName = item.PartnerName;
                        if (month == 1)
                        {
                            dataItemLastMonth.Year = (year - 1).ToString();
                            dataItemLastMonth.Month = "12";
                        }
                        else
                        {
                            dataItemLastMonth.Month = (month - 1).ToString();
                            dataItemLastMonth.Year = year.ToString();
                        }
                        dataItemLastMonth.MarketCode = item.MarketCode;
                        dataItemLastMonth.MarketName = item.MarketName;
                        // Add vào list mới
                        listDataCompareMonthConvert.Add(dataItemLastMonth);
                    }

                    listTemp.Add(item.PartnerName);
                }
            }
            else
            {
                // Danh sách các thị trường
                List<string> listMarket = new List<string>();
                // truyền các thị trường vào list tạm
                foreach (ReportDetailtServiceType item in listDataCompareMonth)
                {
                    if (listMarket.Contains(item.MarketName))
                    {
                        continue;
                    }
                    listMarket.Add(item.MarketName);
                }

                foreach (string item in listMarket)
                {
                    List<ReportDetailtServiceType> listDataItemLastYear = listDataCompareMonth.Where(x => x.MarketName == item && x.Month == month.ToString() && x.Year == (year - 1).ToString()).ToList();
                    List<ReportDetailtServiceType> listDataItemYear = listDataCompareMonth.Where(x => x.MarketName == item && x.Month == month.ToString() && x.Year == year.ToString()).ToList();
                    List<ReportDetailtServiceType> listDataItemLastMonth = listDataCompareMonth.Where(x => x.MarketName == item && x.Month == (month - 1).ToString() && x.Year == year.ToString()).ToList();
                    // Trường hợp tháng 1
                    if (month == 1)
                    {
                        listDataItemLastMonth = listDataCompareMonth.Where(x => x.MarketName == item && x.Month == "12" && x.Year == (year - 1).ToString()).ToList();
                    }

                    double sumDSCKLastYear = listDataItemLastYear.Sum(x => x.DSCK);
                    double sumDSCKYear = listDataItemYear.Sum(x => x.DSCK);
                    double sumDSCKLastMonth = listDataItemLastMonth.Sum(x => x.DSCK);

                    listDataCompareMonthConvert.Add(

                        // Cùng kì năm trước
                        new ReportDetailtServiceType()
                        {
                            PartnerName = item,
                            Month = month.ToString(),
                            Year = (year - 1).ToString(),
                            DSCK = sumDSCKLastYear
                        }
                    );

                    if (month == 1)
                    {
                        listDataCompareMonthConvert.Add(
                            // Tháng trước
                            new ReportDetailtServiceType()
                            {
                                PartnerName = item,
                                Month = "12",
                                Year = (year - 1).ToString(),
                                DSCK = sumDSCKLastMonth
                            }
                        );
                    }
                    else
                    {
                        listDataCompareMonthConvert.Add(
                            // Tháng trước
                            new ReportDetailtServiceType()
                            {
                                PartnerName = item,
                                Month = (month - 1).ToString(),
                                Year = year.ToString(),
                                DSCK = sumDSCKLastMonth
                            }
                        );
                    }


                    listDataCompareMonthConvert.Add(
                        // Tháng hiện tại
                        new ReportDetailtServiceType()
                        {
                            PartnerName = item,
                            Month = month.ToString(),
                            Year = year.ToString(),
                            DSCK = sumDSCKYear
                        }
                    );
                }
            }

            // Số record của mảng
            int countArray = 1;
            GradationCompare[] arrayGradation = new GradationCompare[countArray * listDataCompareMonthConvert.Count];
            int count = 0;
            foreach (ReportDetailtServiceType item in listDataCompareMonthConvert)
            {
                // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                arrayGradation[count] = new GradationCompare()
                {
                    NameGradationCompare = string.Format("Chuyển khoản {0}/{1}", item.Month, item.Year),
                    amount = item.DSCK,
                    NameType = item.PartnerName
                };

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

            GradationChartPie[] arrayGradation = null;
            int count = 0;
            // Trường hợp không phải là thị trường Châu á
            if (!marketID.Equals("005"))
            {
                // Get dữ liệu của năm hiện tại
                List<ReportDetailtServiceType> listData = listDataCompareMonth.Where(x => x.Year == year.ToString() && x.Month == month.ToString()).ToList();

                foreach (ReportDetailtServiceType item in listData)
                {
                    item.TongDS = item.DSChiQuay + item.DSChiNha + item.DSCK;
                }

                double sumTongDS = listData.Sum(x => x.TongDS);

                arrayGradation = new GradationChartPie[listData.Count()];
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
            }
            else
            {
                List<string> listTemp = new List<string>();
                List<ReportDetailtServiceType> listDataItemYear = listDataCompareMonth.Where(x => x.Month == month.ToString() && x.Year == year.ToString()).ToList();

                foreach (ReportDetailtServiceType item in listDataItemYear)
                {
                    if (!listTemp.Contains(item.MarketName))
                    {
                        listTemp.Add(item.MarketName);
                    }

                    item.TongDS = item.DSChiQuay + item.DSChiNha + item.DSCK;
                }

                // Trong doanh số trong tháng hiện tại
                double sumTongDSForMonth = listDataItemYear.Sum(x => x.TongDS);

                arrayGradation = new GradationChartPie[listTemp.Count()];
                List<string> listColor = new List<string>() { "#FFBF00", "#40FF00", "#2ECCFA", "#9A2EFE", "#FE2EF7", "#0000FF", "#08088A", "#B40431", "#6E6E6E" };

                foreach (string item in listTemp)
                {
                    List<ReportDetailtServiceType> listDataItem = listDataCompareMonth.Where(x => x.MarketName == item).ToList();

                    double sumTongDSYear = listDataItem.Sum(x => x.TongDS);

                    // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                    arrayGradation[count] = new GradationChartPie()
                    {
                        category = item,
                        value = sumTongDSForMonth == 0 ? 0 : Math.Round((sumTongDSYear / sumTongDSForMonth) * 100, 2, MidpointRounding.ToEven),
                        color = listColor[count]
                    };
                    count++;
                }

            }


            if (arrayGradation == null)
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

            GradationChartPie[] arrayGradation = null;
            int count = 0;
            // Trường hợp không phải là thị trường Châu á
            if (!marketID.Equals("005"))
            {
                // Get dữ liệu của năm hiện tại
                List<ReportDetailtServiceType> listData = listDataCompareMonth.Where(x => x.Year == year.ToString() && x.Month == (month - 1).ToString()).ToList();

                // trường hợp với get dữ liệu tháng 1
                if (month == 1)
                {
                    listData = listDataCompareMonth.Where(x => x.Year == (year - 1).ToString() && x.Month == "12").ToList();
                }

                foreach (ReportDetailtServiceType item in listData)
                {
                    item.TongDS = item.DSChiQuay + item.DSChiNha + item.DSCK;
                }

                double sumTongDS = listData.Sum(x => x.TongDS);

                arrayGradation = new GradationChartPie[listData.Count()];
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
            }
            else
            {
                List<string> listTemp = new List<string>();
                List<ReportDetailtServiceType> listDataItemLastMonth = listDataCompareMonth.Where(x => x.Month == (month - 1).ToString() && x.Year == year.ToString()).ToList();
                if (month == 1)
                {
                    listDataItemLastMonth = listDataCompareMonth.Where(x => x.Year == (year - 1).ToString() && x.Month == "12").ToList();
                }

                foreach (ReportDetailtServiceType item in listDataItemLastMonth)
                {
                    if (!listTemp.Contains(item.MarketName))
                    {
                        listTemp.Add(item.MarketName);
                    }

                    item.TongDS = item.DSChiQuay + item.DSChiNha + item.DSCK;
                }

                // Trong doanh số trong tháng hiện tại
                double sumTongDSForLastMonth = listDataItemLastMonth.Sum(x => x.TongDS);

                arrayGradation = new GradationChartPie[listTemp.Count()];
                List<string> listColor = new List<string>() { "#FFBF00", "#40FF00", "#2ECCFA", "#9A2EFE", "#FE2EF7", "#0000FF", "#08088A", "#B40431", "#6E6E6E" };

                foreach (string item in listTemp)
                {
                    List<ReportDetailtServiceType> listDataItem = listDataCompareMonth.Where(x => x.MarketName == item).ToList();

                    double sumTongDSLastMonth = listDataItem.Sum(x => x.TongDS);

                    // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                    arrayGradation[count] = new GradationChartPie()
                    {
                        category = item,
                        value = sumTongDSForLastMonth == 0 ? 0 : Math.Round((sumTongDSLastMonth / sumTongDSForLastMonth) * 100, 2, MidpointRounding.ToEven),
                        color = listColor[count]
                    };
                    count++;
                }

            }


            if (arrayGradation == null)
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
        /// Get data cho việc vẽ biểu đồ tròn cho so sánh giai đoạn theo cùng kì năm trước
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult SearchDataCompareMonthPieLastYear([DataSourceRequest]DataSourceRequest request, int month, int year, string reportTypeID, string marketID)
        {

            List<ReportDetailtServiceType> listDataCompareMonth = new ReportBL().ReportDetailtCompareMonthForOne(year, month, reportTypeID, marketID);

            GradationChartPie[] arrayGradation = null;
            int count = 0;
            // Trường hợp không phải là thị trường Châu á
            if (!marketID.Equals("005"))
            {
                // Get dữ liệu của năm hiện tại
                List<ReportDetailtServiceType> listData = listDataCompareMonth.Where(x => x.Year == (year - 1).ToString() && x.Month == month.ToString()).ToList();

                foreach (ReportDetailtServiceType item in listData)
                {
                    item.TongDS = item.DSChiQuay + item.DSChiNha + item.DSCK;
                }

                double sumTongDS = listData.Sum(x => x.TongDS);

                arrayGradation = new GradationChartPie[listData.Count()];
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
            }
            else
            {
                List<string> listTemp = new List<string>();
                List<ReportDetailtServiceType> listDataItemMonthLastYear = listDataCompareMonth.Where(x => x.Year == (year - 1).ToString() && x.Month == month.ToString()).ToList();

                foreach (ReportDetailtServiceType item in listDataItemMonthLastYear)
                {
                    if (!listTemp.Contains(item.MarketName))
                    {
                        listTemp.Add(item.MarketName);
                    }

                    item.TongDS = item.DSChiQuay + item.DSChiNha + item.DSCK;
                }

                // Trong doanh số trong tháng hiện tại
                double sumTongDSForMonthLastYear = listDataItemMonthLastYear.Sum(x => x.TongDS);

                arrayGradation = new GradationChartPie[listTemp.Count()];
                List<string> listColor = new List<string>() { "#FFBF00", "#40FF00", "#2ECCFA", "#9A2EFE", "#FE2EF7", "#0000FF", "#08088A", "#B40431", "#6E6E6E" };

                foreach (string item in listTemp)
                {
                    List<ReportDetailtServiceType> listDataItem = listDataCompareMonth.Where(x => x.MarketName == item).ToList();

                    double sumTongDSLastMonth = listDataItem.Sum(x => x.TongDS);

                    // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                    arrayGradation[count] = new GradationChartPie()
                    {
                        category = item,
                        value = sumTongDSForMonthLastYear == 0 ? 0 : Math.Round((sumTongDSLastMonth / sumTongDSForMonthLastYear) * 100, 2, MidpointRounding.ToEven),
                        color = listColor[count]
                    };
                    count++;
                }

            }


            if (arrayGradation == null)
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

            // List Market
            List<string> listMarket = new List<string>();
            foreach (ReportDetailtServiceType item in listDataCompareMonth)
            {
                if (!listMarket.Contains(item.MarketName))
                {
                    listMarket.Add(item.MarketName);
                }
                item.ReportID = item.PartnerName;
                item.TongDS = item.DSChiQuay + item.DSChiNha + item.DSCK;
            }

            if (marketID.Contains("005"))
            {
                List<ReportDetailtServiceType> listDataCompareMonthConvert = new List<ReportDetailtServiceType>();

                foreach (string item in listMarket)
                {
                    // Month
                    List<ReportDetailtServiceType> listDataItemMonth = listDataCompareMonth.Where(x => x.MarketName == item && x.Month == month.ToString() && x.Year == year.ToString()).ToList();
                    // Last month
                    List<ReportDetailtServiceType> listDataItemLastMonth = listDataCompareMonth.Where(x => x.MarketName == item && x.Month == (month - 1).ToString() && x.Year == year.ToString()).ToList();

                    // trường hợp month = 1
                    if (month == 1)
                    {
                        listDataItemLastMonth = listDataCompareMonth.Where(x => x.MarketName == item && x.Month == "12" && x.Year == (year - 1).ToString()).ToList();
                    }
                    List<ReportDetailtServiceType> listDataItemMonthLastYear = listDataCompareMonth.Where(x => x.MarketName == item && x.Month == month.ToString() && x.Year == (year - 1).ToString()).ToList();

                    if (listDataItemMonth.Count == 0)
                    {
                        listDataCompareMonthConvert.Add(
                            new ReportDetailtServiceType()
                            {
                                MarketName = "Thị trường Châu Á",
                                PartnerName = item,
                                Month = month.ToString(),
                                Year = year.ToString()
                            }
                        );
                    }

                    if (listDataItemLastMonth.Count == 0)
                    {
                        if (month == 1)
                        {
                            listDataCompareMonthConvert.Add(
                                new ReportDetailtServiceType()
                                {
                                    MarketName = "Thị trường Châu Á",
                                    PartnerName = item,
                                    Month = "12",
                                    Year = (year - 1).ToString()
                                }
                            );
                        }
                        else
                        {
                            listDataCompareMonthConvert.Add(
                                new ReportDetailtServiceType()
                                {
                                    MarketName = "Thị trường Châu Á",
                                    PartnerName = item,
                                    Month = (month - 1).ToString(),
                                    Year = year.ToString()
                                }
                            );
                        }
                    }

                    if (listDataItemMonthLastYear.Count == 0)
                    {
                        listDataCompareMonthConvert.Add(
                            new ReportDetailtServiceType()
                            {
                                MarketName = "Thị trường Châu Á",
                                PartnerName = item,
                                Month = month.ToString(),
                                Year = (year - 1).ToString()
                            }
                        );
                    }

                    // Add month
                    listDataCompareMonthConvert.Add(
                        new ReportDetailtServiceType()
                        {
                            MarketName = "Thị trường Châu Á",
                            PartnerName = item,
                            DSChiQuay = listDataItemMonth.Sum(x => x.DSChiQuay),
                            DSChiNha = listDataItemMonth.Sum(x => x.DSChiNha),
                            DSCK = listDataItemMonth.Sum(x => x.DSCK),
                            TongDS = listDataItemMonth.Sum(x => x.TongDS),
                            Month = month.ToString(),
                            Year = year.ToString()
                        }
                    );

                    // Add last month
                    if (month == 1)
                    {
                        listDataCompareMonthConvert.Add(
                            new ReportDetailtServiceType()
                            {
                                MarketName = "Thị trường Châu Á",
                                PartnerName = item,
                                DSChiQuay = listDataItemLastMonth.Sum(x => x.DSChiQuay),
                                DSChiNha = listDataItemLastMonth.Sum(x => x.DSChiNha),
                                DSCK = listDataItemLastMonth.Sum(x => x.DSCK),
                                TongDS = listDataItemLastMonth.Sum(x => x.TongDS),
                                Month = "12",
                                Year = (year - 1).ToString()
                            }
                        );
                    }
                    else
                    {
                        listDataCompareMonthConvert.Add(
                            new ReportDetailtServiceType()
                            {
                                MarketName = "Thị trường Châu Á",
                                PartnerName = item,
                                DSChiQuay = listDataItemLastMonth.Sum(x => x.DSChiQuay),
                                DSChiNha = listDataItemLastMonth.Sum(x => x.DSChiNha),
                                DSCK = listDataItemLastMonth.Sum(x => x.DSCK),
                                TongDS = listDataItemLastMonth.Sum(x => x.TongDS),
                                Month = (month - 1).ToString(),
                                Year = year.ToString()
                            }
                        );
                    }

                    // Add month last year
                    listDataCompareMonthConvert.Add(
                        new ReportDetailtServiceType()
                        {
                            MarketName = "Thị trường Châu Á",
                            PartnerName = item,
                            DSChiQuay = listDataItemMonthLastYear.Sum(x => x.DSChiQuay),
                            DSChiNha = listDataItemMonthLastYear.Sum(x => x.DSChiNha),
                            DSCK = listDataItemMonthLastYear.Sum(x => x.DSCK),
                            TongDS = listDataItemMonthLastYear.Sum(x => x.TongDS),
                            Month = month.ToString(),
                            Year = (year - 1).ToString()
                        }
                    );
                }

                if (listDataCompareMonthConvert.Count > 0)
                {
                    listDataCompareMonth = new List<ReportDetailtServiceType>(listDataCompareMonthConvert);
                }
            }
            
            List<ReportDetailtServiceType> listDataCommpareMonthClone = new List<ReportDetailtServiceType>(listDataCompareMonth);

            double sumTongDSYear = listDataCompareMonth.Where(x => x.Year == year.ToString() && x.Month == month.ToString()).Sum(x => x.TongDS);
            double sumTongDSLastYear = listDataCompareMonth.Where(x => x.Year == (year - 1).ToString() && x.Month == month.ToString()).Sum(x => x.TongDS);
            double sumTongDSLastMonth = listDataCompareMonth.Where(x => x.Year == year.ToString() && x.Month == (month - 1).ToString()).Sum(x => x.TongDS);
            if (month == 1)
            {
                sumTongDSLastMonth = listDataCompareMonth.Where(x => x.Year == (year - 1).ToString() && x.Month == "12").Sum(x => x.TongDS);
            }
            List<ReportDetailtServiceType> listDataConvert = new List<ReportDetailtServiceType>();

            // Khởi tạo datatable
            DataTable table = new DataTable();
            // Tạo các cột cho datatable
            //table.Columns.Add("STT", typeof(String));
            table.Columns.Add("PartnerName", typeof(String));
            table.Columns.Add("LK1", typeof(double));
            table.Columns.Add("LK2", typeof(double));
            table.Columns.Add("LK3", typeof(double));

            table.Columns.Add("MarketName", typeof(String));

            List<string> listTemp = new List<string>();

            foreach (ReportDetailtServiceType item in listDataCommpareMonthClone)
            {
                // Cùng kì
                ReportDetailtServiceType dataItemLastYear = listDataCompareMonth.Find(x => x.PartnerName == item.PartnerName && x.Month == month.ToString() && x.Year == (year - 1).ToString());
                ReportDetailtServiceType dataItemYear = listDataCompareMonth.Find(x => x.PartnerName == item.PartnerName && x.Month == month.ToString() && x.Year == year.ToString());
                ReportDetailtServiceType dataItemLastMonth = listDataCompareMonth.Find(x => x.PartnerName == item.PartnerName && x.Month == (month - 1).ToString() && x.Year == year.ToString());
                // Trường hợp tháng 1
                if (month == 1)
                {
                    dataItemLastMonth = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == "12" && x.Year == (year - 1).ToString());
                }
                // Trường hợp năm không có đối tác
                if (dataItemLastYear == null)
                {
                    dataItemLastYear = new ReportDetailtServiceType();
                    dataItemLastYear.PartnerName = item.PartnerName;
                    dataItemLastYear.MarketName = item.MarketName;
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
                    dataItemYear.MarketName = item.MarketName;
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
                    dataItemLastMonth.MarketName = item.MarketName;
                    dataItemLastMonth.DSChiQuay = 0;
                    dataItemLastMonth.DSChiNha = 0;
                    dataItemLastMonth.DSCK = 0;
                    dataItemLastMonth.TongDS = 0;
                    if (month == 1)
                    {
                        dataItemLastMonth.Year = (year - 1).ToString();
                        dataItemLastMonth.Month = "12";
                    }
                    else
                    {
                        dataItemLastMonth.Month = (month - 1).ToString();
                        dataItemLastMonth.Year = year.ToString();
                    }
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

                    table.Rows.Add(dataItemYear.PartnerName, valueYear, valueLastMonth, valueLastYear, item.MarketName);
                }
            }

            return Json(table.ToDataSourceResult(request), JsonRequestBehavior.AllowGet);
        }
    }
}
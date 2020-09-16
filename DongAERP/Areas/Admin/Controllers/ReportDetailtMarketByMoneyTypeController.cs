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
    public class ReportDetailtMarketByMoneyTypeController : Controller
    {
        // GET: Admin/ReportDetailtMarketByMoneyType
        public ActionResult Index()
        {
            return View();
        }

        /// <summary>
        /// Báo cáo thị trường theo loại tiền chi trả tất cả
        /// </summary>
        /// <returns></returns>
        public ActionResult MarketForTotal(DateTime? fromDay, DateTime? toDay, string reportTypeID, string marketID)
        {
            string nameUrl = "Báo cáo chi tiết/Theo doanh số chi trả/Theo thị trường/Loại tiền/Tất cả";
            if (marketID.Equals("1"))
            {
                nameUrl = "Báo cáo chi tiết/Theo doanh số chi trả/Theo thị trường/Loại tiền/Thị trường Châu Á";
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
            string nameUrl = "Báo cáo chi tiết/Theo doanh số chi trả/Theo thị trường/Loại tiền/Tất cả";
            if (marketID.Equals("1"))
            {
                nameUrl = "Báo cáo chi tiết/Theo doanh số chi trả/Theo thị trường/Loại tiền/Thị trường Châu Á";
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
        /// Màn hình báo cáo cho năm
        /// </summary>
        /// <returns></returns>
        public ActionResult MarketForTotalForYear(DateTime? fromDate, DateTime? toDate, string reportTypeID, string marketID)
        {
            string nameUrl = "Báo cáo chi tiết/Theo doanh số chi trả/Theo thị trường/Loại tiền/Tất cả";
            if (marketID.Equals("1"))
            {
                nameUrl = "Báo cáo chi tiết/Theo doanh số chi trả/Theo thị trường/Loại tiền/Thị trường Châu Á";
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
        /// Báo cáo thị trường theo loại tiền chi trả từng thị trường
        /// </summary>
        /// <returns></returns>
        public ActionResult MarketForOne(DateTime? fromDay, DateTime? toDay, string reportTypeID, string marketID)
        {
            string nameUrl = "Báo cáo chi tiết/Theo doanh số chi trả/Theo thị trường/Loại tiền chi trả/Từng thị trường";
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
        /// Màn hình báo cáo chi tiết theo tháng
        /// </summary>
        /// <returns></returns>
        public ActionResult MarketForOneForMonth(DateTime? fromDate, DateTime? toDate, string reportTypeID, string marketID)
        {
            string nameUrl = "Báo cáo chi tiết/Theo doanh số chi trả/Theo thị trường/Loại hình tiền/Từng thị trường";
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
        /// Màn hình báo cáo chi tiết theo tháng
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
            string nameUrl = "Báo cáo chi tiết/Theo doanh số chi trả/Theo thị trường/Loại tiền chi trả/So sánh - Giai đoạn - Tất cả thị trường";
            if (marketID.Equals("1"))
            {
                nameUrl = "Báo cáo chi tiết/Theo doanh số chi trả/Theo thị trường/Loại tiền chi trả/So sánh - Giai đoạn - Thị trường Châu Á";
            }
            ViewBag.NameURL = nameUrl;
            ViewBag.NameURL = nameUrl;
            // Khởi tạo datatable
            DataTable table = new DataTable();
            // Tạo các cột cho datatable
            table.Columns.Add("MarketName", typeof(String));
            table.Columns.Add("VND1", typeof(double));
            table.Columns.Add("VND2", typeof(double));

            table.Columns.Add("USD1", typeof(double));
            table.Columns.Add("USD2", typeof(double));

            table.Columns.Add("EUR1", typeof(double));
            table.Columns.Add("EUR2", typeof(double));

            table.Columns.Add("CAD1", typeof(double));
            table.Columns.Add("CAD2", typeof(double));

            table.Columns.Add("AUD1", typeof(double));
            table.Columns.Add("AUD2", typeof(double));

            table.Columns.Add("GBP1", typeof(double));
            table.Columns.Add("GBP2", typeof(double));

            table.Columns.Add("TDS1", typeof(double));
            table.Columns.Add("TDS2", typeof(double));


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
            string nameUrl = "Báo cáo chi tiết/Theo doanh số chi trả/Theo thị trường/Loại tiền chi trả/So sánh - Giai đoạn - Từng thị trường";
            ViewBag.NameURL = nameUrl;
            // Khởi tạo datatable
            DataTable table = new DataTable();
            // Tạo các cột cho datatable
            table.Columns.Add("PartnerName", typeof(String));
            table.Columns.Add("VND1", typeof(double));
            table.Columns.Add("VND2", typeof(double));

            table.Columns.Add("USD1", typeof(double));
            table.Columns.Add("USD2", typeof(double));

            table.Columns.Add("EUR1", typeof(double));
            table.Columns.Add("EUR2", typeof(double));

            table.Columns.Add("CAD1", typeof(double));
            table.Columns.Add("CAD2", typeof(double));

            table.Columns.Add("AUD1", typeof(double));
            table.Columns.Add("AUD2", typeof(double));

            table.Columns.Add("GBP1", typeof(double));
            table.Columns.Add("GBP2", typeof(double));

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
            string nameUrl = "Báo cáo chi tiết/Theo doanh số chi trả/Theo thị trường/Loại tiền chi trả/So sánh - Theo tháng - Tất cả thị trường";
            if (marketID.Equals("1"))
            {
                nameUrl = "Báo cáo chi tiết/Theo doanh số chi trả/Theo thị trường/Loại tiền chi trả/So sánh - Theo tháng - Thị trường Châu Á";
            }
            ViewBag.NameURL = nameUrl;
            // Khởi tạo datatable
            DataTable table = new DataTable();
            // Tạo các cột cho datatable
            table.Columns.Add("MarketName", typeof(String));
            table.Columns.Add("VND1", typeof(double));
            table.Columns.Add("VND2", typeof(double));
            table.Columns.Add("VND3", typeof(double));

            table.Columns.Add("USD1", typeof(double));
            table.Columns.Add("USD2", typeof(double));
            table.Columns.Add("USD3", typeof(double));

            table.Columns.Add("EUR1", typeof(double));
            table.Columns.Add("EUR2", typeof(double));
            table.Columns.Add("EUR3", typeof(double));

            table.Columns.Add("CAD1", typeof(double));
            table.Columns.Add("CAD2", typeof(double));
            table.Columns.Add("CAD3", typeof(double));

            table.Columns.Add("AUD1", typeof(double));
            table.Columns.Add("AUD2", typeof(double));
            table.Columns.Add("AUD3", typeof(double));

            table.Columns.Add("GBP1", typeof(double));
            table.Columns.Add("GBP2", typeof(double));
            table.Columns.Add("GBP3", typeof(double));

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
            string nameUrl = "Báo cáo chi tiết/Theo doanh số chi trả/Theo thị trường/Loại tiền chi trả/So sánh - Theo tháng - Từng thị trường";
            ViewBag.NameURL = nameUrl;
            // Khởi tạo datatable
            DataTable table = new DataTable();
            // Tạo các cột cho datatable
            table.Columns.Add("PartnerName", typeof(String));
            table.Columns.Add("VND1", typeof(double));
            table.Columns.Add("VND2", typeof(double));
            table.Columns.Add("VND3", typeof(double));

            table.Columns.Add("USD1", typeof(double));
            table.Columns.Add("USD2", typeof(double));
            table.Columns.Add("USD3", typeof(double));

            table.Columns.Add("EUR1", typeof(double));
            table.Columns.Add("EUR2", typeof(double));
            table.Columns.Add("EUR3", typeof(double));

            table.Columns.Add("CAD1", typeof(double));
            table.Columns.Add("CAD2", typeof(double));
            table.Columns.Add("CAD3", typeof(double));

            table.Columns.Add("AUD1", typeof(double));
            table.Columns.Add("AUD2", typeof(double));
            table.Columns.Add("AUD3", typeof(double));

            table.Columns.Add("GBP1", typeof(double));
            table.Columns.Add("GBP2", typeof(double));
            table.Columns.Add("GBP3", typeof(double));

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
            List<ReportDetailtForTotalMoneyType> listData = new ReportBL().SearchReportDetailtMTForAll(fromDay, toDay, reportTypeID, marketID);

            foreach (ReportDetailtForTotalMoneyType item in listData)
            {
                item.ReportID = item.MarketName;
                if (marketID.Equals("0"))
                {
                    item.MarketName = "";
                }
                item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;
            }

            listData.Add(
                new ReportDetailtForTotalMoneyType()
                {
                    ReportID ="Tổng",
                    VND = listData.Sum(x => x.VND),
                    USD = listData.Sum(x => x.USD),
                    EUR = listData.Sum(x => x.EUR),
                    CAD = listData.Sum(x => x.CAD),
                    AUD = listData.Sum(x => x.AUD),
                    GBP = listData.Sum(x => x.GBP),
                }
            );
            
            return Json(listData.ToDataSourceResult(request), JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// Search report detailt theo tháng của loại tiền
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public ActionResult SearchMarketForTotalForMonth([DataSourceRequest]DataSourceRequest request, DateTime fromDate, DateTime toDate, string reportTypeID, string marketID)
        {
            List<ReportDetailtForTotalMoneyType> listData = new ReportBL().SearchReportDetailtMTForAllForMonth(fromDate, toDate, reportTypeID, marketID);

            foreach (ReportDetailtForTotalMoneyType item in listData)
            {
                item.ReportID = item.MarketName;
                if (marketID.Equals("0"))
                {
                    item.MarketName = "";
                }
                item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;
            }

            listData.Add(
                new ReportDetailtForTotalMoneyType()
                {
                    ReportID = "Tổng",
                    VND = listData.Sum(x => x.VND),
                    USD = listData.Sum(x => x.USD),
                    EUR = listData.Sum(x => x.EUR),
                    CAD = listData.Sum(x => x.CAD),
                    AUD = listData.Sum(x => x.AUD),
                    GBP = listData.Sum(x => x.GBP),
                }
            );

            return Json(listData.ToDataSourceResult(request), JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// Search report detailt theo năm của loại tiền
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public ActionResult SearchMarketForTotalForYear([DataSourceRequest]DataSourceRequest request, DateTime fromDate, DateTime toDate, string reportTypeID, string marketID)
        {
            List<ReportDetailtForTotalMoneyType> listData = new ReportBL().SearchReportDetailtMTForAllForYear(fromDate, toDate, reportTypeID, marketID);

            foreach (ReportDetailtForTotalMoneyType item in listData)
            {
                item.ReportID = item.MarketName;
                if (marketID.Equals("0"))
                {
                    item.MarketName = "";
                }
                item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;
            }

            listData.Add(
                new ReportDetailtForTotalMoneyType()
                {
                    ReportID = "Tổng",
                    VND = listData.Sum(x => x.VND),
                    USD = listData.Sum(x => x.USD),
                    EUR = listData.Sum(x => x.EUR),
                    CAD = listData.Sum(x => x.CAD),
                    AUD = listData.Sum(x => x.AUD),
                    GBP = listData.Sum(x => x.GBP),
                }
            );

            return Json(listData.ToDataSourceResult(request), JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// Search report day theo ngày nhập vào
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public ActionResult SearchMarketForTotalConvert([DataSourceRequest]DataSourceRequest request, DateTime fromDay, DateTime toDay, string reportTypeID, string marketID)
        {
            List<ReportDetailtForTotalMoneyType> listData = new ReportBL().SearchReportDetailtMTForAllConvert(fromDay, toDay, reportTypeID, marketID);

            foreach (ReportDetailtForTotalMoneyType item in listData)
            {
                item.ReportID = item.MarketName;
                if (marketID.Equals("0"))
                {
                    item.MarketName = "";
                }
                item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;
            }

            listData.Add(
                new ReportDetailtForTotalMoneyType()
                {
                    ReportID = "Tổng",
                    VND = listData.Sum(x => x.VND),
                    USD = listData.Sum(x => x.USD),
                    EUR = listData.Sum(x => x.EUR),
                    CAD = listData.Sum(x => x.CAD),
                    AUD = listData.Sum(x => x.AUD),
                    GBP = listData.Sum(x => x.GBP),
                    TongDS = listData.Sum(x => x.TongDS)
                }
            );

            return Json(listData.ToDataSourceResult(request), JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// Search report day theo ngày nhập vào
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public ActionResult SearchMarketForTotalForMonthConvert([DataSourceRequest]DataSourceRequest request, DateTime fromDate, DateTime toDate, string reportTypeID, string marketID)
        {
            List<ReportDetailtForTotalMoneyType> listData = new ReportBL().SearchReportDetailtMTForAllForMonthConvert(fromDate, toDate, reportTypeID, marketID);

            foreach (ReportDetailtForTotalMoneyType item in listData)
            {
                item.ReportID = item.MarketName;
                if (marketID.Equals("0"))
                {
                    item.MarketName = "";
                }
                item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;
            }

            listData.Add(
                new ReportDetailtForTotalMoneyType()
                {
                    ReportID = "Tổng",
                    VND = listData.Sum(x => x.VND),
                    USD = listData.Sum(x => x.USD),
                    EUR = listData.Sum(x => x.EUR),
                    CAD = listData.Sum(x => x.CAD),
                    AUD = listData.Sum(x => x.AUD),
                    GBP = listData.Sum(x => x.GBP),
                    TongDS = listData.Sum(x => x.TongDS)
                }
            );

            return Json(listData.ToDataSourceResult(request), JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// Search report day theo ngày nhập vào
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public ActionResult SearchMarketForTotalForYearConvert([DataSourceRequest]DataSourceRequest request, DateTime fromDate, DateTime toDate, string reportTypeID, string marketID)
        {
            List<ReportDetailtForTotalMoneyType> listData = new ReportBL().SearchReportDetailtMTForAllForYearConvert(fromDate, toDate, reportTypeID, marketID);

            foreach (ReportDetailtForTotalMoneyType item in listData)
            {
                item.ReportID = item.MarketName;
                if (marketID.Equals("0"))
                {
                    item.MarketName = "Tất cả thị trường";
                }
                item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;
            }

            listData.Add(
                new ReportDetailtForTotalMoneyType()
                {
                    ReportID = "Tổng",
                    VND = listData.Sum(x => x.VND),
                    USD = listData.Sum(x => x.USD),
                    EUR = listData.Sum(x => x.EUR),
                    CAD = listData.Sum(x => x.CAD),
                    AUD = listData.Sum(x => x.AUD),
                    GBP = listData.Sum(x => x.GBP),
                    TongDS = listData.Sum(x => x.TongDS)
                }
            );

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
            List<ReportDetailtForTotalMoneyType> listData = new ReportBL().SearchReportDetailtMTForOne(fromDay, toDay, reportTypeID, marketID);

            foreach (ReportDetailtForTotalMoneyType item in listData)
            {
                item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;
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
            List<ReportDetailtForTotalMoneyType> listData = new ReportBL().SearchReportDetailtMTForOneForMonth(fromDate, toDate, reportTypeID, marketID);

            foreach (ReportDetailtForTotalMoneyType item in listData)
            {
                item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;
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
        public ActionResult SearchMarketForOneForYear([DataSourceRequest]DataSourceRequest request, DateTime fromDate, DateTime toDate, string reportTypeID, string marketID)
        {
            List<ReportDetailtForTotalMoneyType> listData = new ReportBL().SearchReportDetailtMTForOneForYear(fromDate, toDate, reportTypeID, marketID);

            foreach (ReportDetailtForTotalMoneyType item in listData)
            {
                item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;
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
        public ActionResult SearchMarketForOneConvert([DataSourceRequest]DataSourceRequest request, DateTime fromDay, DateTime toDay, string reportTypeID, string marketID)
        {
            List<ReportDetailtForTotalMoneyType> listData = new ReportBL().SearchReportDetailtMTForOneConvert(fromDay, toDay, reportTypeID, marketID);

            foreach (ReportDetailtForTotalMoneyType item in listData)
            {
                item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;
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
        public ActionResult SearchMarketForOneForMonthConvert([DataSourceRequest]DataSourceRequest request, DateTime fromDate, DateTime toDate, string reportTypeID, string marketID)
        {
            List<ReportDetailtForTotalMoneyType> listData = new ReportBL().SearchReportDetailtMTForOneForMonthConvert(fromDate, toDate, reportTypeID, marketID);

            foreach (ReportDetailtForTotalMoneyType item in listData)
            {
                item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;
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
        public ActionResult SearchMarketForOneForYearConvert([DataSourceRequest]DataSourceRequest request, DateTime fromDate, DateTime toDate, string reportTypeID, string marketID)
        {
            List<ReportDetailtForTotalMoneyType> listData = new ReportBL().SearchReportDetailtMTForOneForYearConvert(fromDate, toDate, reportTypeID, marketID);

            foreach (ReportDetailtForTotalMoneyType item in listData)
            {
                item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;
            }
            return Json(listData.ToDataSourceResult(request), JsonRequestBehavior.AllowGet);
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
            List<ReportDetailtForTotalMoneyType> listDataGradation = new ReportBL().ReportDetailtMTGradationCompareForAllConvert(year, gradation, reportTypeID, marketID);
            List<ReportDetailtForTotalMoneyType> listDataGradationConvert = new List<ReportDetailtForTotalMoneyType>();

            // Khởi tạo datatable
            DataTable table = new DataTable();
            // Tạo các cột cho datatable
            table.Columns.Add("MarketName", typeof(String));
            table.Columns.Add("VND1", typeof(double));
            table.Columns.Add("VND2", typeof(double));

            table.Columns.Add("USD1", typeof(double));
            table.Columns.Add("USD2", typeof(double));

            table.Columns.Add("EUR1", typeof(double));
            table.Columns.Add("EUR2", typeof(double));

            table.Columns.Add("CAD1", typeof(double));
            table.Columns.Add("CAD2", typeof(double));

            table.Columns.Add("AUD1", typeof(double));
            table.Columns.Add("AUD2", typeof(double));

            table.Columns.Add("GBP1", typeof(double));
            table.Columns.Add("GBP2", typeof(double));

            table.Columns.Add("TDS1", typeof(double));
            table.Columns.Add("TDS2", typeof(double));

            // Trường hợp chọn tất cả thị trường
            if (marketID.Equals("0"))
            {
                foreach (ReportDetailtForTotalMoneyType item in listDataGradation)
                {
                    item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;
                }

                // Danh sách mã thị trường của Tất cả
                List<string> listMarket = new List<string>() { "003", "005", "001", "002", "014", "004" };
                 
                foreach (string item in listMarket)
                {
                    // Cùng kì
                    ReportDetailtForTotalMoneyType dataItemLastYear = listDataGradation.Find(x => x.MarketCode == item && x.Year == (year - 1).ToString());
                    ReportDetailtForTotalMoneyType dataItemYear = listDataGradation.Find(x => x.MarketCode == item && x.Year == year.ToString());

                    // add item vào table
                    table.Rows.Add(dataItemLastYear.MarketName
                        , dataItemLastYear.VND, dataItemYear.VND, dataItemLastYear.USD, dataItemYear.USD, dataItemLastYear.EUR, dataItemYear.EUR
                        , dataItemLastYear.CAD, dataItemYear.CAD, dataItemLastYear.AUD, dataItemYear.AUD, dataItemLastYear.GBP, dataItemYear.GBP
                        , dataItemLastYear.TongDS, dataItemYear.TongDS
                        );
                }

                DataRow row = table.NewRow();
                row["MarketName"] = "Tổng";
                row["VND1"] = table.Compute("Sum(VND1)", "");
                row["VND2"] = table.Compute("Sum(VND2)", "");

                row["USD1"] = table.Compute("Sum(USD1)", "");
                row["USD2"] = table.Compute("Sum(USD2)", "");

                row["EUR1"] = table.Compute("Sum(EUR1)", "");
                row["EUR2"] = table.Compute("Sum(EUR2)", "");

                row["CAD1"] = table.Compute("Sum(CAD1)", "");
                row["CAD2"] = table.Compute("Sum(CAD2)", "");

                row["AUD1"] = table.Compute("Sum(AUD1)", "");
                row["AUD2"] = table.Compute("Sum(AUD2)", "");

                row["GBP1"] = table.Compute("Sum(GBP1)", "");
                row["GBP2"] = table.Compute("Sum(GBP2)", "");

                row["TDS1"] = table.Compute("Sum(TDS1)", "");
                row["TDS2"] = table.Compute("Sum(TDS2)", "");
                table.Rows.Add(row);

            }
            else
            {
                // List thị trường
                List<string> listMarket = new List<string>();

                foreach (ReportDetailtForTotalMoneyType item in listDataGradation)
                {
                    if (!listMarket.Contains(item.MarketName))
                    {
                        listMarket.Add(item.MarketName);
                    }
                    item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;
                }

                foreach(string item in listMarket)
                {
                    List<ReportDetailtForTotalMoneyType> listDetailtYear = listDataGradation.Where(x => x.MarketName == item && x.Year == year.ToString()).ToList();
                    List<ReportDetailtForTotalMoneyType> listDetailtLastYear = listDataGradation.Where(x => x.MarketName == item && x.Year == (year - 1).ToString()).ToList();

                    if(listDetailtYear.Count == 0)
                    {
                        listDetailtYear = new List<ReportDetailtForTotalMoneyType>();
                    }

                    if (listDetailtLastYear.Count == 0)
                    {
                        listDetailtLastYear = new List<ReportDetailtForTotalMoneyType>();
                    }

                    // Year
                    double sumVNDYear = listDetailtYear.Sum(x => x.VND);
                    double sumUSDYear = listDetailtYear.Sum(x => x.USD);
                    double sumEURYear = listDetailtYear.Sum(x => x.EUR);
                    double sumCADYear = listDetailtYear.Sum(x => x.CAD);
                    double sumAUDYear = listDetailtYear.Sum(x => x.AUD);
                    double sumGBPYear = listDetailtYear.Sum(x => x.GBP);
                    double sumTongDSYear = sumVNDYear + sumUSDYear + sumEURYear + sumCADYear + sumAUDYear + sumGBPYear;

                    // LastYear
                    double sumVNDLastYear = listDetailtLastYear.Sum(x => x.VND);
                    double sumUSDLastYear = listDetailtLastYear.Sum(x => x.USD);
                    double sumEURLastYear = listDetailtLastYear.Sum(x => x.EUR);
                    double sumCADLastYear = listDetailtLastYear.Sum(x => x.CAD);
                    double sumAUDLastYear = listDetailtLastYear.Sum(x => x.AUD);
                    double sumGBPLastYear = listDetailtLastYear.Sum(x => x.GBP);
                    double sumTongDSLastYear = sumVNDLastYear + sumUSDLastYear + sumEURLastYear + sumCADLastYear + sumAUDLastYear + sumGBPLastYear;

                    table.Rows.Add(item
                        , sumVNDLastYear, sumVNDYear, sumUSDYear, sumUSDLastYear, sumEURLastYear, sumEURYear
                        , sumCADLastYear, sumCADYear, sumAUDLastYear, sumAUDYear, sumGBPLastYear, sumGBPYear
                        , sumTongDSLastYear, sumTongDSYear
                        );
                }

                DataRow row = table.NewRow();
                row["MarketName"] = "Tổng";
                row["VND1"] = table.Compute("Sum(VND1)", "");
                row["VND2"] = table.Compute("Sum(VND2)", "");

                row["USD1"] = table.Compute("Sum(USD1)", "");
                row["USD2"] = table.Compute("Sum(USD2)", "");

                row["EUR1"] = table.Compute("Sum(EUR1)", "");
                row["EUR2"] = table.Compute("Sum(EUR2)", "");

                row["CAD1"] = table.Compute("Sum(CAD1)", "");
                row["CAD2"] = table.Compute("Sum(CAD2)", "");

                row["AUD1"] = table.Compute("Sum(AUD1)", "");
                row["AUD2"] = table.Compute("Sum(AUD2)", "");

                row["GBP1"] = table.Compute("Sum(GBP1)", "");
                row["GBP2"] = table.Compute("Sum(GBP2)", "");

                row["TDS1"] = table.Compute("Sum(TDS1)", "");
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
        public ActionResult SearchGradationCompareForAllConvertCompare([DataSourceRequest]DataSourceRequest request, int gradation, int year, string reportTypeID, string marketID)
        {
            List<ReportDetailtForTotalMoneyType> listDataGradation = new ReportBL().ReportDetailtMTGradationCompareForAllConvert(year, gradation, reportTypeID, marketID);

            foreach (ReportDetailtForTotalMoneyType item in listDataGradation)
            {
                item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;
            }

            // Danh sách mã thị trường của Tất cả
            List<string> listMarket = new List<string>() { "003", "005", "001", "002", "014", "004" };

            // Khởi tạo datatable
            DataTable table = new DataTable();
            // Tạo các cột cho datatable
            table.Columns.Add("MarketName", typeof(String));
            table.Columns.Add("VND1", typeof(double));
            table.Columns.Add("USD1", typeof(double));
            table.Columns.Add("EUR1", typeof(double));
            table.Columns.Add("CAD1", typeof(double));
            table.Columns.Add("AUD1", typeof(double));
            table.Columns.Add("GBP1", typeof(double));
            table.Columns.Add("TDS1", typeof(double));

            // Trường hợp chọn tất cả thị trường
            if (marketID.Equals("0"))
            {
                foreach (string item in listMarket)
                {
                    // Cùng kì
                    ReportDetailtForTotalMoneyType dataItemLastYear = listDataGradation.Find(x => x.MarketCode == item && x.Year == (year - 1).ToString());
                    ReportDetailtForTotalMoneyType dataItemYear = listDataGradation.Find(x => x.MarketCode == item && x.Year == year.ToString());

                    // Trường hợp năm trước có đối tác và năm nay không có
                    if (dataItemLastYear != null && dataItemYear == null)
                    {
                        dataItemYear = new ReportDetailtForTotalMoneyType();
                        dataItemYear.MarketCode = dataItemLastYear.MarketCode;
                        dataItemYear.MarketName = dataItemLastYear.MarketName;
                        dataItemYear.VND = 0;
                        dataItemYear.USD = 0;
                        dataItemYear.EUR = 0;
                        dataItemYear.CAD = 0;
                        dataItemYear.AUD = 0;
                        dataItemYear.GBP = 0;
                        dataItemYear.Year = year.ToString();
                    }

                    // Trường hợp năm trước không có đối tác và năm nay có đối tác
                    if (dataItemYear != null && dataItemLastYear == null)
                    {
                        dataItemLastYear = new ReportDetailtForTotalMoneyType();
                        dataItemLastYear.MarketCode = dataItemYear.MarketCode;
                        dataItemLastYear.MarketName = dataItemYear.MarketName;
                        dataItemLastYear.VND = 0;
                        dataItemLastYear.USD = 0;
                        dataItemLastYear.EUR = 0;
                        dataItemLastYear.CAD = 0;
                        dataItemLastYear.AUD = 0;
                        dataItemLastYear.GBP = 0;
                        dataItemLastYear.Year = (year - 1).ToString();
                    }

                    double sumVND = dataItemYear.VND - dataItemLastYear.VND;
                    double sumUSD = dataItemYear.USD - dataItemLastYear.USD;
                    double sumEUR = dataItemYear.EUR - dataItemLastYear.EUR;
                    double sumCAD = dataItemYear.CAD - dataItemLastYear.CAD;
                    double sumAUD = dataItemYear.AUD - dataItemLastYear.AUD;
                    double sumGBP = dataItemYear.GBP - dataItemLastYear.GBP;
                    double sumTongDS = sumVND + sumUSD + sumEUR + sumCAD + sumAUD + sumGBP;

                    // add item vào table
                    table.Rows.Add(dataItemLastYear.MarketName
                        , Math.Round(sumVND, 2, MidpointRounding.ToEven)
                        , Math.Round(sumUSD, 2, MidpointRounding.ToEven)
                        , Math.Round(sumEUR, 2, MidpointRounding.ToEven)
                        , Math.Round(sumCAD, 2, MidpointRounding.ToEven)
                        , Math.Round(sumAUD, 2, MidpointRounding.ToEven)
                        , Math.Round(sumGBP, 2, MidpointRounding.ToEven)
                        , Math.Round(sumTongDS, 2, MidpointRounding.ToEven)
                        );
                }

                DataRow row = table.NewRow();
                row["MarketName"] = "Tổng";
                row["VND1"] = table.Compute("Sum(VND1)", "");
                row["USD1"] = table.Compute("Sum(USD1)", "");
                row["EUR1"] = table.Compute("Sum(EUR1)", "");
                row["CAD1"] = table.Compute("Sum(CAD1)", "");
                row["AUD1"] = table.Compute("Sum(AUD1)", "");
                row["GBP1"] = table.Compute("Sum(GBP1)", "");
                row["TDS1"] = table.Compute("Sum(TDS1)", "");
                table.Rows.Add(row);
            }
            else
            {
                // List thị trường
                listMarket = new List<string>();

                foreach (ReportDetailtForTotalMoneyType item in listDataGradation)
                {
                    if (!listMarket.Contains(item.MarketName))
                    {
                        listMarket.Add(item.MarketName);
                    }
                    item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;
                }

                foreach (string item in listMarket)
                {
                    List<ReportDetailtForTotalMoneyType> listDetailtYear = listDataGradation.Where(x => x.MarketName == item && x.Year == year.ToString()).ToList();
                    List<ReportDetailtForTotalMoneyType> listDetailtLastYear = listDataGradation.Where(x => x.MarketName == item && x.Year == (year - 1).ToString()).ToList();

                    // Year
                    double sumVNDYear = listDetailtYear.Sum(x => x.VND);
                    double sumUSDYear = listDetailtYear.Sum(x => x.USD);
                    double sumEURYear = listDetailtYear.Sum(x => x.EUR);
                    double sumCADYear = listDetailtYear.Sum(x => x.CAD);
                    double sumAUDYear = listDetailtYear.Sum(x => x.AUD);
                    double sumGBPYear = listDetailtYear.Sum(x => x.GBP);
                    double sumTongDSYear = sumVNDYear + sumUSDYear + sumEURYear + sumCADYear + sumAUDYear + sumGBPYear;

                    // LastYear
                    double sumVNDLastYear = listDetailtLastYear.Sum(x => x.VND);
                    double sumUSDLastYear = listDetailtLastYear.Sum(x => x.USD);
                    double sumEURLastYear = listDetailtLastYear.Sum(x => x.EUR);
                    double sumCADLastYear = listDetailtLastYear.Sum(x => x.CAD);
                    double sumAUDLastYear = listDetailtLastYear.Sum(x => x.AUD);
                    double sumGBPLastYear = listDetailtLastYear.Sum(x => x.GBP);
                    double sumTongDSLastYear = sumVNDLastYear + sumUSDLastYear + sumEURLastYear + sumCADLastYear + sumAUDLastYear + sumGBPLastYear;

                    double sumVND = sumVNDYear - sumVNDLastYear;
                    double sumUSD = sumUSDYear - sumUSDLastYear;
                    double sumEUR = sumEURYear - sumEURLastYear;
                    double sumCAD = sumCADYear - sumCADLastYear;
                    double sumAUD = sumAUDYear - sumAUDLastYear;
                    double sumGBP = sumGBPYear - sumGBPLastYear;
                    double sumTongDS = sumVND + sumUSD + sumEUR + sumCAD + sumAUD + sumGBP;

                    // add item vào table
                    table.Rows.Add(item
                        , Math.Round(sumVND, 2, MidpointRounding.ToEven)
                        , Math.Round(sumUSD, 2, MidpointRounding.ToEven)
                        , Math.Round(sumEUR, 2, MidpointRounding.ToEven)
                        , Math.Round(sumCAD, 2, MidpointRounding.ToEven)
                        , Math.Round(sumAUD, 2, MidpointRounding.ToEven)
                        , Math.Round(sumGBP, 2, MidpointRounding.ToEven)
                        , Math.Round(sumTongDS, 2, MidpointRounding.ToEven)
                        );
                }
                
                DataRow row = table.NewRow();
                row["MarketName"] = "Tổng";
                row["VND1"] = table.Compute("Sum(VND1)", "");
                row["USD1"] = table.Compute("Sum(USD1)", "");
                row["EUR1"] = table.Compute("Sum(EUR1)", "");
                row["CAD1"] = table.Compute("Sum(CAD1)", "");
                row["AUD1"] = table.Compute("Sum(AUD1)", "");
                row["GBP1"] = table.Compute("Sum(GBP1)", "");
                row["TDS1"] = table.Compute("Sum(TDS1)", "");
                table.Rows.Add(row);
            }

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
        public ActionResult SearchColumnChartReportForGradationPercent([DataSourceRequest]DataSourceRequest request, int gradation, int year, string reportTypeID, string marketID)
        {
            List<ReportDetailtForTotalMoneyType> listDataGradation = new ReportBL().ReportDetailtMTGradationCompareForAllPercent(year, gradation, reportTypeID, marketID);

            // Số mảng cần tạo
            int arrayCount = 6;
            GradationCharColumn[] arrayGradation = new GradationCharColumn[listDataGradation.Count * arrayCount];
            int count = 0;
            // group theo nhóm
            int tooltip = 1;
            foreach (ReportDetailtForTotalMoneyType item in listDataGradation)
            {
                // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                arrayGradation[count] = new GradationCharColumn()
                {
                    Serie = item.MarketName,
                    Segmento = string.Format("VND {0}", item.Year),
                    Valor1 = item.VND,
                    Tooltip = tooltip
                };

                count++;
                arrayGradation[count] = new GradationCharColumn()
                {
                    Serie = item.MarketName,
                    Segmento = string.Format("USD {0}", item.Year),
                    Valor1 = item.USD,
                    Tooltip = tooltip
                };

                count++;
                arrayGradation[count] = new GradationCharColumn()
                {
                    Serie = item.MarketName,
                    Segmento = string.Format("EUR {0}", item.Year),
                    Valor1 = item.EUR,
                    Tooltip = tooltip
                };
                count++;

                arrayGradation[count] = new GradationCharColumn()
                {
                    Serie = item.MarketName,
                    Segmento = string.Format("CAD {0}", item.Year),
                    Valor1 = item.CAD,
                    Tooltip = tooltip
                };
                count++;

                arrayGradation[count] = new GradationCharColumn()
                {
                    Serie = item.MarketName,
                    Segmento = string.Format("AUD {0}", item.Year),
                    Valor1 = item.AUD,
                    Tooltip = tooltip
                };
                count++;

                arrayGradation[count] = new GradationCharColumn()
                {
                    Serie = item.MarketName,
                    Segmento = string.Format("GBP {0}", item.Year),
                    Valor1 = item.GBP,
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
        /// Get data cho việc vẽ biểu đồ cột cho so sánh theo giai đoạn - Tất cẳ thị trường
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult SearchColumnChartReportForGradationVND([DataSourceRequest]DataSourceRequest request, int gradation, int year, string reportTypeID, string marketID)
        {
            List<ReportDetailtForTotalMoneyType> listDataGradation = new ReportBL().ReportDetailtMTGradationCompareForAllConvert(year, gradation, reportTypeID, marketID);

            GradationCompare[] arrayGradation = null;

            // Trường hợp chọn tất cả thị trường
            if (marketID.Equals("0"))
            {
                // Số record của mảng
                int countArray = 1;
                arrayGradation = new GradationCompare[countArray * listDataGradation.Count];
                int count = 0;
                foreach (ReportDetailtForTotalMoneyType item in listDataGradation)
                {
                    // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                    arrayGradation[count] = new GradationCompare()
                    {
                        NameGradationCompare = string.Format("VND năm {0}", item.Year),
                        amount = item.VND,
                        NameType = item.MarketName
                    };
                    count++;

                    year = year - 1;
                }
            }
            else
            {
                // Trường hợp thị trường  Châu á
                List<string> listAsianItem = new List<string>();
                foreach (ReportDetailtForTotalMoneyType item in listDataGradation)
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

                List<ReportDetailtForTotalMoneyType> listDSMarketYear = new List<ReportDetailtForTotalMoneyType>();
                List<ReportDetailtForTotalMoneyType> listDSMarketLastYear = new List<ReportDetailtForTotalMoneyType>();

                foreach (string item in listAsianItem)
                {
                    listDSMarketYear = listDataGradation.Where(x => x.MarketName == item && x.Year == year.ToString()).ToList();
                    listDSMarketLastYear = listDataGradation.Where(x => x.MarketName == item && x.Year == (year - 1).ToString()).ToList();

                    double sumVNDYear = listDSMarketYear.Sum(x => x.VND);
                    double sumVNDLastYear = listDSMarketLastYear.Sum(x => x.VND);

                    // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                    arrayGradation[count] = new GradationCompare()
                    {
                        NameGradationCompare = string.Format("VND năm {0}", year),
                        amount = sumVNDYear,
                        NameType = item
                    };
                    count++;

                    // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                    arrayGradation[count] = new GradationCompare()
                    {
                        NameGradationCompare = string.Format("VND năm {0}", year - 1),
                        amount = sumVNDLastYear,
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
        /// Get data cho việc vẽ biểu đồ cột cho so sánh theo giai đoạn - Tất cẳ thị trường - USD
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult SearchColumnChartReportForGradationUSD([DataSourceRequest]DataSourceRequest request, int gradation, int year, string reportTypeID, string marketID)
        {
            List<ReportDetailtForTotalMoneyType> listDataGradation = new ReportBL().ReportDetailtMTGradationCompareForAllConvert(year, gradation, reportTypeID, marketID);

            GradationCompare[] arrayGradation = null;

            // Trường hợp chọn tất cả thị trường
            if (marketID.Equals("0"))
            {
                // Số record của mảng
                int countArray = 1;
                arrayGradation = new GradationCompare[countArray * listDataGradation.Count];
                int count = 0;
                foreach (ReportDetailtForTotalMoneyType item in listDataGradation)
                {
                    // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                    arrayGradation[count] = new GradationCompare()
                    {
                        NameGradationCompare = string.Format("USD năm {0}", item.Year),
                        amount = item.USD,
                        NameType = item.MarketName
                    };
                    count++;

                    year = year - 1;
                }
            }
            else
            {
                // Trường hợp thị trường  Châu á
                List<string> listAsianItem = new List<string>();
                foreach (ReportDetailtForTotalMoneyType item in listDataGradation)
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

                List<ReportDetailtForTotalMoneyType> listDSMarketYear = new List<ReportDetailtForTotalMoneyType>();
                List<ReportDetailtForTotalMoneyType> listDSMarketLastYear = new List<ReportDetailtForTotalMoneyType>();

                foreach (string item in listAsianItem)
                {
                    listDSMarketYear = listDataGradation.Where(x => x.MarketName == item && x.Year == year.ToString()).ToList();
                    listDSMarketLastYear = listDataGradation.Where(x => x.MarketName == item && x.Year == (year - 1).ToString()).ToList();

                    double sumUSDYear = listDSMarketYear.Sum(x => x.USD);
                    double sumUSDLastYear = listDSMarketLastYear.Sum(x => x.USD);

                    // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                    arrayGradation[count] = new GradationCompare()
                    {
                        NameGradationCompare = string.Format("USD năm {0}", year),
                        amount = sumUSDYear,
                        NameType = item
                    };
                    count++;

                    // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                    arrayGradation[count] = new GradationCompare()
                    {
                        NameGradationCompare = string.Format("USD năm {0}", year - 1),
                        amount = sumUSDLastYear,
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
        /// Get data cho việc vẽ biểu đồ cột cho so sánh theo giai đoạn - Tất cẳ thị trường - EUR
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult SearchColumnChartReportForGradationEUR([DataSourceRequest]DataSourceRequest request, int gradation, int year, string reportTypeID, string marketID)
        {
            List<ReportDetailtForTotalMoneyType> listDataGradation = new ReportBL().ReportDetailtMTGradationCompareForAllConvert(year, gradation, reportTypeID, marketID);

            GradationCompare[] arrayGradation = null;

            // Trường hợp chọn tất cả thị trường
            if (marketID.Equals("0"))
            {
                // Số record của mảng
                int countArray = 1;
                arrayGradation = new GradationCompare[countArray * listDataGradation.Count];
                int count = 0;
                foreach (ReportDetailtForTotalMoneyType item in listDataGradation)
                {
                    // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                    arrayGradation[count] = new GradationCompare()
                    {
                        NameGradationCompare = string.Format("EUR năm {0}", item.Year),
                        amount = item.EUR,
                        NameType = item.MarketName
                    };
                    count++;

                    year = year - 1;
                }
            }
            else
            {
                // Trường hợp thị trường  Châu á
                List<string> listAsianItem = new List<string>();
                foreach (ReportDetailtForTotalMoneyType item in listDataGradation)
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

                List<ReportDetailtForTotalMoneyType> listDSMarketYear = new List<ReportDetailtForTotalMoneyType>();
                List<ReportDetailtForTotalMoneyType> listDSMarketLastYear = new List<ReportDetailtForTotalMoneyType>();

                foreach (string item in listAsianItem)
                {
                    listDSMarketYear = listDataGradation.Where(x => x.MarketName == item && x.Year == year.ToString()).ToList();
                    listDSMarketLastYear = listDataGradation.Where(x => x.MarketName == item && x.Year == (year - 1).ToString()).ToList();

                    double sumEURYear = listDSMarketYear.Sum(x => x.EUR);
                    double sumEURLastYear = listDSMarketLastYear.Sum(x => x.EUR);

                    // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                    arrayGradation[count] = new GradationCompare()
                    {
                        NameGradationCompare = string.Format("EUR năm {0}", year),
                        amount = sumEURYear,
                        NameType = item
                    };
                    count++;

                    // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                    arrayGradation[count] = new GradationCompare()
                    {
                        NameGradationCompare = string.Format("EUR năm {0}", year - 1),
                        amount = sumEURLastYear,
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
        /// Get data cho việc vẽ biểu đồ cột cho so sánh theo giai đoạn - Tất cẳ thị trường - CAD
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult SearchColumnChartReportForGradationCAD([DataSourceRequest]DataSourceRequest request, int gradation, int year, string reportTypeID, string marketID)
        {
            List<ReportDetailtForTotalMoneyType> listDataGradation = new ReportBL().ReportDetailtMTGradationCompareForAllConvert(year, gradation, reportTypeID, marketID);
            
            GradationCompare[] arrayGradation = null;

            // Trường hợp chọn tất cả thị trường
            if (marketID.Equals("0"))
            {
                // Số record của mảng
                int countArray = 1;
                arrayGradation = new GradationCompare[countArray * listDataGradation.Count];
                int count = 0;
                foreach (ReportDetailtForTotalMoneyType item in listDataGradation)
                {
                    // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                    arrayGradation[count] = new GradationCompare()
                    {
                        NameGradationCompare = string.Format("CAD năm {0}", item.Year),
                        amount = item.CAD,
                        NameType = item.MarketName
                    };
                    count++;

                    year = year - 1;
                }
            }
            else
            {
                // Trường hợp thị trường  Châu á
                List<string> listAsianItem = new List<string>();
                foreach (ReportDetailtForTotalMoneyType item in listDataGradation)
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

                List<ReportDetailtForTotalMoneyType> listDSMarketYear = new List<ReportDetailtForTotalMoneyType>();
                List<ReportDetailtForTotalMoneyType> listDSMarketLastYear = new List<ReportDetailtForTotalMoneyType>();

                foreach (string item in listAsianItem)
                {
                    listDSMarketYear = listDataGradation.Where(x => x.MarketName == item && x.Year == year.ToString()).ToList();
                    listDSMarketLastYear = listDataGradation.Where(x => x.MarketName == item && x.Year == (year - 1).ToString()).ToList();

                    double sumCADYear = listDSMarketYear.Sum(x => x.CAD);
                    double sumCADLastYear = listDSMarketLastYear.Sum(x => x.CAD);

                    // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                    arrayGradation[count] = new GradationCompare()
                    {
                        NameGradationCompare = string.Format("CAD năm {0}", year),
                        amount = sumCADYear,
                        NameType = item
                    };
                    count++;

                    // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                    arrayGradation[count] = new GradationCompare()
                    {
                        NameGradationCompare = string.Format("CAD năm {0}", year - 1),
                        amount = sumCADLastYear,
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
        /// Get data cho việc vẽ biểu đồ cột cho so sánh theo giai đoạn - Tất cẳ thị trường - AUD
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult SearchColumnChartReportForGradationAUD([DataSourceRequest]DataSourceRequest request, int gradation, int year, string reportTypeID, string marketID)
        {
            List<ReportDetailtForTotalMoneyType> listDataGradation = new ReportBL().ReportDetailtMTGradationCompareForAllConvert(year, gradation, reportTypeID, marketID);
            GradationCompare[] arrayGradation = null;

            // Trường hợp chọn tất cả thị trường
            if (marketID.Equals("0"))
            {
                // Số record của mảng
                int countArray = 1;
                arrayGradation = new GradationCompare[countArray * listDataGradation.Count];
                int count = 0;
                foreach (ReportDetailtForTotalMoneyType item in listDataGradation)
                {
                    // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                    arrayGradation[count] = new GradationCompare()
                    {
                        NameGradationCompare = string.Format("AUD năm {0}", item.Year),
                        amount = item.AUD,
                        NameType = item.MarketName
                    };
                    count++;

                    year = year - 1;
                }
            }
            else
            {
                // Trường hợp thị trường  Châu á
                List<string> listAsianItem = new List<string>();
                foreach (ReportDetailtForTotalMoneyType item in listDataGradation)
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

                List<ReportDetailtForTotalMoneyType> listDSMarketYear = new List<ReportDetailtForTotalMoneyType>();
                List<ReportDetailtForTotalMoneyType> listDSMarketLastYear = new List<ReportDetailtForTotalMoneyType>();

                foreach (string item in listAsianItem)
                {
                    listDSMarketYear = listDataGradation.Where(x => x.MarketName == item && x.Year == year.ToString()).ToList();
                    listDSMarketLastYear = listDataGradation.Where(x => x.MarketName == item && x.Year == (year - 1).ToString()).ToList();

                    double sumAUDYear = listDSMarketYear.Sum(x => x.AUD);
                    double sumAUDLastYear = listDSMarketLastYear.Sum(x => x.AUD);

                    // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                    arrayGradation[count] = new GradationCompare()
                    {
                        NameGradationCompare = string.Format("AUD năm {0}", year),
                        amount = sumAUDYear,
                        NameType = item
                    };
                    count++;

                    // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                    arrayGradation[count] = new GradationCompare()
                    {
                        NameGradationCompare = string.Format("AUD năm {0}", year - 1),
                        amount = sumAUDLastYear,
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
        /// Get data cho việc vẽ biểu đồ cột cho so sánh theo giai đoạn - Tất cẳ thị trường - GBP
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult SearchColumnChartReportForGradationGBP([DataSourceRequest]DataSourceRequest request, int gradation, int year, string reportTypeID, string marketID)
        {
            List<ReportDetailtForTotalMoneyType> listDataGradation = new ReportBL().ReportDetailtMTGradationCompareForAllConvert(year, gradation, reportTypeID, marketID);

            GradationCompare[] arrayGradation = null;

            // Trường hợp chọn tất cả thị trường
            if (marketID.Equals("0"))
            {
                // Số record của mảng
                int countArray = 1;
                arrayGradation = new GradationCompare[countArray * listDataGradation.Count];
                int count = 0;
                foreach (ReportDetailtForTotalMoneyType item in listDataGradation)
                {
                    // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                    arrayGradation[count] = new GradationCompare()
                    {
                        NameGradationCompare = string.Format("GBP năm {0}", item.Year),
                        amount = item.GBP,
                        NameType = item.MarketName
                    };
                    count++;

                    year = year - 1;
                }
            }
            else
            {
                // Trường hợp thị trường  Châu á
                List<string> listAsianItem = new List<string>();
                foreach (ReportDetailtForTotalMoneyType item in listDataGradation)
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

                List<ReportDetailtForTotalMoneyType> listDSMarketYear = new List<ReportDetailtForTotalMoneyType>();
                List<ReportDetailtForTotalMoneyType> listDSMarketLastYear = new List<ReportDetailtForTotalMoneyType>();

                foreach (string item in listAsianItem)
                {
                    listDSMarketYear = listDataGradation.Where(x => x.MarketName == item && x.Year == year.ToString()).ToList();
                    listDSMarketLastYear = listDataGradation.Where(x => x.MarketName == item && x.Year == (year - 1).ToString()).ToList();

                    double sumGBPYear = listDSMarketYear.Sum(x => x.GBP);
                    double sumGBPLastYear = listDSMarketLastYear.Sum(x => x.GBP);

                    // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                    arrayGradation[count] = new GradationCompare()
                    {
                        NameGradationCompare = string.Format("GBP năm {0}", year),
                        amount = sumGBPYear,
                        NameType = item
                    };
                    count++;

                    // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                    arrayGradation[count] = new GradationCompare()
                    {
                        NameGradationCompare = string.Format("GBP năm {0}", year - 1),
                        amount = sumGBPLastYear,
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
        /// Get data cho việc vẽ biểu đồ cột cho so sánh theo giai đoạn
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult SearchGridReportForGradationForOne([DataSourceRequest]DataSourceRequest request, int gradation, int year, string reportTypeID, string marketID)
        {
            List<ReportDetailtForTotalMoneyType> listDataGradation = new ReportBL().ReportDetailtMTGradationCompareForOneConvert(year, gradation, reportTypeID, marketID);

            foreach (ReportDetailtForTotalMoneyType item in listDataGradation)
            {
                item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;
            }

            // Khởi tạo datatable
            DataTable table = new DataTable();
            // Tạo các cột cho datatable
            table.Columns.Add("PartnerName", typeof(String));
            table.Columns.Add("VND1", typeof(double));
            table.Columns.Add("VND2", typeof(double));

            table.Columns.Add("USD1", typeof(double));
            table.Columns.Add("USD2", typeof(double));

            table.Columns.Add("EUR1", typeof(double));
            table.Columns.Add("EUR2", typeof(double));

            table.Columns.Add("CAD1", typeof(double));
            table.Columns.Add("CAD2", typeof(double));

            table.Columns.Add("AUD1", typeof(double));
            table.Columns.Add("AUD2", typeof(double));

            table.Columns.Add("GBP1", typeof(double));
            table.Columns.Add("GBP2", typeof(double));

            table.Columns.Add("TDS1", typeof(double));
            table.Columns.Add("TDS2", typeof(double));

            table.Columns.Add("MarketName", typeof(String));


            foreach (ReportDetailtForTotalMoneyType item in listDataGradation)
            {

                // Cùng kì
                ReportDetailtForTotalMoneyType dataItemLastYear = listDataGradation.Find(x => x.PartnerCode == item.PartnerCode && x.Year == (year - 1).ToString());
                ReportDetailtForTotalMoneyType dataItemYear = listDataGradation.Find(x => x.PartnerCode == item.PartnerCode && x.Year == year.ToString());

                // Trường hợp năm trước có đối tác và năm nay không có
                if (dataItemLastYear != null && dataItemYear == null)
                {
                    dataItemYear = new ReportDetailtForTotalMoneyType();
                    dataItemYear.PartnerCode = dataItemLastYear.PartnerCode;
                    dataItemYear.PartnerName = dataItemLastYear.PartnerName;
                    dataItemYear.Year = year.ToString();
                }

                // Trường hợp năm trước không có đối tác và năm nay có đối tác
                if (dataItemYear != null && dataItemLastYear == null)
                {
                    dataItemLastYear = new ReportDetailtForTotalMoneyType();
                    dataItemLastYear.PartnerCode = dataItemYear.PartnerCode;
                    dataItemLastYear.PartnerName = dataItemYear.PartnerName;
                    dataItemLastYear.Year = (year - 1).ToString();
                }

                // Check tồn tại của item
                string value = string.Format("PartnerName='{0}'", item.PartnerName);
                DataRow[] foundRows = table.Select(value);
                if (dataItemLastYear != null && dataItemYear != null && foundRows.Count() == 0)
                {
                    // add item vào table
                    table.Rows.Add(item.PartnerName
                        , dataItemLastYear.VND, dataItemYear.VND, dataItemLastYear.USD, dataItemYear.USD, dataItemLastYear.EUR, dataItemYear.EUR
                        , dataItemLastYear.CAD, dataItemYear.CAD, dataItemLastYear.AUD, dataItemYear.AUD, dataItemLastYear.GBP, dataItemYear.GBP
                        , dataItemLastYear.TongDS, dataItemYear.TongDS
                        , item.MarketName
                        );
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
        public ActionResult SearchGridReportForGradationForOneCompare([DataSourceRequest]DataSourceRequest request, int gradation, int year, string reportTypeID, string marketID)
        {
            List<ReportDetailtForTotalMoneyType> listDataGradation = new ReportBL().ReportDetailtMTGradationCompareForOneConvert(year, gradation, reportTypeID, marketID);

            foreach (ReportDetailtForTotalMoneyType item in listDataGradation)
            {
                item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;
            }

            // Khởi tạo datatable
            DataTable table = new DataTable();
            // Tạo các cột cho datatable
            table.Columns.Add("PartnerName", typeof(String));
            table.Columns.Add("VND1", typeof(double));
            table.Columns.Add("USD1", typeof(double));
            table.Columns.Add("EUR1", typeof(double));
            table.Columns.Add("CAD1", typeof(double));
            table.Columns.Add("AUD1", typeof(double));
            table.Columns.Add("GBP1", typeof(double));
            table.Columns.Add("TDS1", typeof(double));

            table.Columns.Add("MarketName", typeof(string));

            foreach (ReportDetailtForTotalMoneyType item in listDataGradation)
            {

                // Cùng kì
                ReportDetailtForTotalMoneyType dataItemLastYear = listDataGradation.Find(x => x.PartnerCode == item.PartnerCode && x.Year == (year - 1).ToString());
                ReportDetailtForTotalMoneyType dataItemYear = listDataGradation.Find(x => x.PartnerCode == item.PartnerCode && x.Year == year.ToString());

                // Trường hợp năm trước có đối tác và năm nay không có
                if (dataItemLastYear != null && dataItemYear == null)
                {
                    dataItemYear = new ReportDetailtForTotalMoneyType();
                    dataItemYear.PartnerCode = dataItemLastYear.PartnerCode;
                    dataItemYear.PartnerName = dataItemLastYear.PartnerName;
                    dataItemYear.Year = year.ToString();
                }

                // Trường hợp năm trước không có đối tác và năm nay có đối tác
                if (dataItemYear != null && dataItemLastYear == null)
                {
                    dataItemLastYear = new ReportDetailtForTotalMoneyType();
                    dataItemLastYear.PartnerCode = dataItemYear.PartnerCode;
                    dataItemLastYear.PartnerName = dataItemYear.PartnerName;
                    dataItemLastYear.Year = (year - 1).ToString();
                }

                double sumVND = dataItemYear.VND - dataItemLastYear.VND;
                double sumUSD = dataItemYear.USD - dataItemLastYear.USD;
                double sumEUR = dataItemYear.EUR - dataItemLastYear.EUR;
                double sumCAD = dataItemYear.CAD - dataItemLastYear.CAD;
                double sumAUD = dataItemYear.AUD - dataItemLastYear.AUD;
                double sumGBP = dataItemYear.GBP - dataItemLastYear.GBP;
                double sumTongDS = sumVND + sumUSD + sumEUR + sumCAD + sumAUD + sumGBP;

                // Check tồn tại của item
                string value = string.Format("PartnerName='{0}'", item.PartnerName);
                DataRow[] foundRows = table.Select(value);
                if (dataItemLastYear != null && dataItemYear != null && foundRows.Count() == 0)
                {
                    // add item vào table
                    table.Rows.Add(item.PartnerName
                        , Math.Round(sumVND, 2, MidpointRounding.ToEven)
                        , Math.Round(sumUSD, 2, MidpointRounding.ToEven)
                        , Math.Round(sumEUR, 2, MidpointRounding.ToEven)
                        , Math.Round(sumCAD, 2, MidpointRounding.ToEven)
                        , Math.Round(sumAUD, 2, MidpointRounding.ToEven)
                        , Math.Round(sumGBP, 2, MidpointRounding.ToEven)
                        , Math.Round(sumTongDS, 2, MidpointRounding.ToEven)
                        , item.MarketName
                        );
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
        public ActionResult SearchColumnChartGradationCompareStackForOne([DataSourceRequest]DataSourceRequest request, int gradation, int year, string reportTypeID, string marketID)
        {
            List<ReportDetailtForTotalMoneyType> listDataGradation = new ReportBL().ReportDetailtMTGradationCompareForOnePercent(year, gradation, reportTypeID, marketID);
            List<ReportDetailtForTotalMoneyType> listDataGradationClone = new List<ReportDetailtForTotalMoneyType>(listDataGradation);

            foreach (ReportDetailtForTotalMoneyType item in listDataGradationClone)
            {
                ReportDetailtForTotalMoneyType dataItemYear = listDataGradation.Find(x => x.PartnerCode == item.PartnerCode && x.Year == year.ToString());
                ReportDetailtForTotalMoneyType dataItemLastYear = listDataGradation.Find(x => x.PartnerCode == item.PartnerCode && x.Year == (year - 1).ToString());

                // Trường hợp năm trước có đối tác và năm nay không có
                if (dataItemLastYear != null && dataItemYear == null)
                {
                    dataItemYear = new ReportDetailtForTotalMoneyType();
                    dataItemYear.PartnerCode = dataItemLastYear.PartnerCode;
                    dataItemYear.PartnerName = dataItemLastYear.PartnerName;
                    dataItemYear.Year = year.ToString();

                    // Add item thiếu vào
                    listDataGradation.Add(dataItemYear);
                }

                // Trường hợp năm trước không có đối tác và năm nay có đối tác
                if (dataItemYear != null && dataItemLastYear == null)
                {
                    dataItemLastYear = new ReportDetailtForTotalMoneyType();
                    dataItemLastYear.PartnerCode = dataItemYear.PartnerCode;
                    dataItemLastYear.PartnerName = dataItemYear.PartnerName;
                    dataItemLastYear.Year = (year - 1).ToString();

                    // Add item thiếu vào
                    listDataGradation.Add(dataItemLastYear);
                }
            }

            // Số mảng cần tạo
            int arrayCount = 6;
            GradationCharColumn[] arrayGradation = new GradationCharColumn[listDataGradation.Count * arrayCount];
            int count = 0;
            foreach (ReportDetailtForTotalMoneyType item in listDataGradation)
            {
                string nameSerie = item.PartnerName;
                if (marketID.Contains("005"))
                {
                    nameSerie = item.MarketName;
                }
                // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                arrayGradation[count] = new GradationCharColumn()
                {
                    Serie = nameSerie,
                    Segmento = string.Format("VND năm {0}", item.Year),
                    Valor1 = item.VND,
                    //Tooltip = tooltip
                };
                count++;

                arrayGradation[count] = new GradationCharColumn()
                {
                    Serie = nameSerie,
                    Segmento = string.Format("USD năm {0}", item.Year),
                    Valor1 = item.USD,
                    //Tooltip = tooltip
                };
                count++;

                arrayGradation[count] = new GradationCharColumn()
                {
                    Serie = nameSerie,
                    Segmento = string.Format("EUR năm {0}", item.Year),
                    Valor1 = item.EUR,
                    //Tooltip = tooltip
                };
                count++;

                arrayGradation[count] = new GradationCharColumn()
                {
                    Serie = nameSerie,
                    Segmento = string.Format("CAD năm {0}", item.Year),
                    Valor1 = item.CAD,
                    //Tooltip = tooltip
                };
                count++;

                arrayGradation[count] = new GradationCharColumn()
                {
                    Serie = nameSerie,
                    Segmento = string.Format("AUD năm {0}", item.Year),
                    Valor1 = item.AUD,
                    //Tooltip = tooltip
                };
                count++;

                arrayGradation[count] = new GradationCharColumn()
                {
                    Serie = nameSerie,
                    Segmento = string.Format("GBP năm {0}", item.Year),
                    Valor1 = item.GBP,
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
            List<ReportDetailtForTotalMoneyType> listDataGradation = new ReportBL().ReportDetailtMTGradationCompareForOneConvert(year, gradationID, reportTypeID, marketID);

            // Số record của mảng
            int countArray = 6;
            GradationCompare[] arrayGradation = new GradationCompare[countArray * listDataGradation.Count];
            int count = 0;
            foreach (ReportDetailtForTotalMoneyType item in listDataGradation)
            {
                // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                arrayGradation[count] = new GradationCompare()
                {
                    NameGradationCompare = "VND",
                    amount = item.VND,
                    NameType = string.Format("{0} {1}", item.PartnerName, item.Year),
                };

                count++;
                arrayGradation[count] = new GradationCompare()
                {
                    NameGradationCompare = "USD",
                    amount = item.USD,
                    NameType = string.Format("{0} {1}", item.PartnerName, item.Year),
                };

                count++;
                arrayGradation[count] = new GradationCompare()
                {
                    NameGradationCompare = "EUR",
                    amount = item.EUR,
                    NameType = string.Format("{0} {1}", item.PartnerName, item.Year),
                };

                count++;
                arrayGradation[count] = new GradationCompare()
                {
                    NameGradationCompare = "CAD",
                    amount = item.CAD,
                    NameType = string.Format("{0} {1}", item.PartnerName, item.Year),
                };

                count++;
                arrayGradation[count] = new GradationCompare()
                {
                    NameGradationCompare = "AUD",
                    amount = item.AUD,
                    NameType = string.Format("{0} {1}", item.PartnerName, item.Year),
                };

                count++;
                arrayGradation[count] = new GradationCompare()
                {
                    NameGradationCompare = "GBP",
                    amount = item.GBP,
                    NameType = string.Format("{0} {1}", item.PartnerName, item.Year),
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
        /// Get data cho việc vẽ biểu đồ cột cho so sánh theo giai đoạn - VND
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult SearchColumnsChartGradationCompareForOneVND([DataSourceRequest]DataSourceRequest request, int gradation, int year, string reportTypeID, string marketID)
        {
            List<ReportDetailtForTotalMoneyType> listDataGradation = new ReportBL().ReportDetailtMTGradationCompareForOneConvert(year, gradation, reportTypeID, marketID);

            GradationCompare[] arrayGradation = null;
            if (!marketID.Contains("005"))
            {
                // Số record của mảng
                int countArray = 1;
                arrayGradation = new GradationCompare[countArray * listDataGradation.Count];

                int count = 0;
                foreach (ReportDetailtForTotalMoneyType item in listDataGradation)
                {
                    // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                    arrayGradation[count] = new GradationCompare()
                    {
                        NameGradationCompare = string.Format("VND năm {0}", item.Year),
                        amount = item.VND,
                        NameType = item.PartnerName
                    };

                    count++;
                    year = year - 1;
                }
            }
            else
            {
                // Trường hợp thị trường  Châu á
                List<string> listAsianItem = new List<string>();
                foreach (ReportDetailtForTotalMoneyType item in listDataGradation)
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
                
                List<ReportDetailtForTotalMoneyType> listDSMarketYear = new List<ReportDetailtForTotalMoneyType>();
                List<ReportDetailtForTotalMoneyType> listDSMarketLastYear = new List<ReportDetailtForTotalMoneyType>();

                foreach (string item in listAsianItem)
                {
                    listDSMarketYear = listDataGradation.Where(x => x.MarketName == item && x.Year == year.ToString()).ToList();
                    listDSMarketLastYear = listDataGradation.Where(x => x.MarketName == item && x.Year == (year - 1).ToString()).ToList();

                    double sumVNDYear = listDSMarketYear.Sum(x => x.VND);
                    double sumVNDLastYear = listDSMarketLastYear.Sum(x => x.VND);

                    // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                    arrayGradation[count] = new GradationCompare()
                    {
                        NameGradationCompare = string.Format("VND năm {0}", year),
                        amount = sumVNDYear,
                        NameType = item
                    };
                    count++;

                    // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                    arrayGradation[count] = new GradationCompare()
                    {
                        NameGradationCompare = string.Format("VND năm {0}", year - 1),
                        amount = sumVNDLastYear,
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
        /// Get data cho việc vẽ biểu đồ cột cho so sánh theo giai đoạn - VND
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult SearchColumnsChartGradationCompareForOneUSD([DataSourceRequest]DataSourceRequest request, int gradation, int year, string reportTypeID, string marketID)
        {
            List<ReportDetailtForTotalMoneyType> listDataGradation = new ReportBL().ReportDetailtMTGradationCompareForOneConvert(year, gradation, reportTypeID, marketID);

            GradationCompare[] arrayGradation = null;
            if (!marketID.Contains("005"))
            {
                // Số record của mảng
                int countArray = 1;
                arrayGradation = new GradationCompare[countArray * listDataGradation.Count];

                int count = 0;
                foreach (ReportDetailtForTotalMoneyType item in listDataGradation)
                {
                    // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                    arrayGradation[count] = new GradationCompare()
                    {
                        NameGradationCompare = string.Format("USD năm {0}", item.Year),
                        amount = item.USD,
                        NameType = item.PartnerName
                    };

                    count++;
                    year = year - 1;
                }
            }
            else
            {
                // Trường hợp thị trường  Châu á
                List<string> listAsianItem = new List<string>();
                foreach (ReportDetailtForTotalMoneyType item in listDataGradation)
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

                List<ReportDetailtForTotalMoneyType> listDSMarketYear = new List<ReportDetailtForTotalMoneyType>();
                List<ReportDetailtForTotalMoneyType> listDSMarketLastYear = new List<ReportDetailtForTotalMoneyType>();

                foreach (string item in listAsianItem)
                {
                    listDSMarketYear = listDataGradation.Where(x => x.MarketName == item && x.Year == year.ToString()).ToList();
                    listDSMarketLastYear = listDataGradation.Where(x => x.MarketName == item && x.Year == (year - 1).ToString()).ToList();

                    double sumUSDYear = listDSMarketYear.Sum(x => x.USD);
                    double sumUSDLastYear = listDSMarketLastYear.Sum(x => x.USD);

                    // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                    arrayGradation[count] = new GradationCompare()
                    {
                        NameGradationCompare = string.Format("USD năm {0}", year),
                        amount = sumUSDYear,
                        NameType = item
                    };
                    count++;

                    // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                    arrayGradation[count] = new GradationCompare()
                    {
                        NameGradationCompare = string.Format("USD năm {0}", year - 1),
                        amount = sumUSDLastYear,
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
        /// Get data cho việc vẽ biểu đồ cột cho so sánh theo giai đoạn - EUR
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult SearchColumnsChartGradationCompareForOneEUR([DataSourceRequest]DataSourceRequest request, int gradation, int year, string reportTypeID, string marketID)
        {
            List<ReportDetailtForTotalMoneyType> listDataGradation = new ReportBL().ReportDetailtMTGradationCompareForOneConvert(year, gradation, reportTypeID, marketID);

            GradationCompare[] arrayGradation = null;
            if (!marketID.Contains("005"))
            {
                // Số record của mảng
                int countArray = 1;
                arrayGradation = new GradationCompare[countArray * listDataGradation.Count];

                int count = 0;
                foreach (ReportDetailtForTotalMoneyType item in listDataGradation)
                {
                    // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                    arrayGradation[count] = new GradationCompare()
                    {
                        NameGradationCompare = string.Format("EUR năm {0}", item.Year),
                        amount = item.EUR,
                        NameType = item.PartnerName
                    };

                    count++;
                    year = year - 1;
                }
            }
            else
            {
                // Trường hợp thị trường  Châu á
                List<string> listAsianItem = new List<string>();
                foreach (ReportDetailtForTotalMoneyType item in listDataGradation)
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

                List<ReportDetailtForTotalMoneyType> listDSMarketYear = new List<ReportDetailtForTotalMoneyType>();
                List<ReportDetailtForTotalMoneyType> listDSMarketLastYear = new List<ReportDetailtForTotalMoneyType>();

                foreach (string item in listAsianItem)
                {
                    listDSMarketYear = listDataGradation.Where(x => x.MarketName == item && x.Year == year.ToString()).ToList();
                    listDSMarketLastYear = listDataGradation.Where(x => x.MarketName == item && x.Year == (year - 1).ToString()).ToList();

                    double sumEURYear = listDSMarketYear.Sum(x => x.EUR);
                    double sumEURLastYear = listDSMarketLastYear.Sum(x => x.EUR);

                    // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                    arrayGradation[count] = new GradationCompare()
                    {
                        NameGradationCompare = string.Format("EUR năm {0}", year),
                        amount = sumEURYear,
                        NameType = item
                    };
                    count++;

                    // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                    arrayGradation[count] = new GradationCompare()
                    {
                        NameGradationCompare = string.Format("EUR năm {0}", year - 1),
                        amount = sumEURLastYear,
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
        /// Get data cho việc vẽ biểu đồ cột cho so sánh theo giai đoạn - CAD
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult SearchColumnsChartGradationCompareForOneCAD([DataSourceRequest]DataSourceRequest request, int gradation, int year, string reportTypeID, string marketID)
        {
            List<ReportDetailtForTotalMoneyType> listDataGradation = new ReportBL().ReportDetailtMTGradationCompareForOneConvert(year, gradation, reportTypeID, marketID);

            GradationCompare[] arrayGradation = null;
            if (!marketID.Contains("005"))
            {
                // Số record của mảng
                int countArray = 1;
                arrayGradation = new GradationCompare[countArray * listDataGradation.Count];

                int count = 0;
                foreach (ReportDetailtForTotalMoneyType item in listDataGradation)
                {
                    // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                    arrayGradation[count] = new GradationCompare()
                    {
                        NameGradationCompare = string.Format("CAD năm {0}", item.Year),
                        amount = item.CAD,
                        NameType = item.PartnerName
                    };

                    count++;
                    year = year - 1;
                }
            }
            else
            {
                // Trường hợp thị trường  Châu á
                List<string> listAsianItem = new List<string>();
                foreach (ReportDetailtForTotalMoneyType item in listDataGradation)
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

                List<ReportDetailtForTotalMoneyType> listDSMarketYear = new List<ReportDetailtForTotalMoneyType>();
                List<ReportDetailtForTotalMoneyType> listDSMarketLastYear = new List<ReportDetailtForTotalMoneyType>();

                foreach (string item in listAsianItem)
                {
                    listDSMarketYear = listDataGradation.Where(x => x.MarketName == item && x.Year == year.ToString()).ToList();
                    listDSMarketLastYear = listDataGradation.Where(x => x.MarketName == item && x.Year == (year - 1).ToString()).ToList();

                    double sumCADYear = listDSMarketYear.Sum(x => x.CAD);
                    double sumCADLastYear = listDSMarketLastYear.Sum(x => x.CAD);

                    // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                    arrayGradation[count] = new GradationCompare()
                    {
                        NameGradationCompare = string.Format("CAD năm {0}", year),
                        amount = sumCADYear,
                        NameType = item
                    };
                    count++;

                    // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                    arrayGradation[count] = new GradationCompare()
                    {
                        NameGradationCompare = string.Format("CAD năm {0}", year - 1),
                        amount = sumCADLastYear,
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
        /// Get data cho việc vẽ biểu đồ cột cho so sánh theo giai đoạn - AUD
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult SearchColumnsChartGradationCompareForOneAUD([DataSourceRequest]DataSourceRequest request, int gradation, int year, string reportTypeID, string marketID)
        {
            List<ReportDetailtForTotalMoneyType> listDataGradation = new ReportBL().ReportDetailtMTGradationCompareForOneConvert(year, gradation, reportTypeID, marketID);

            GradationCompare[] arrayGradation = null;
            if (!marketID.Contains("005"))
            {
                // Số record của mảng
                int countArray = 1;
                arrayGradation = new GradationCompare[countArray * listDataGradation.Count];

                int count = 0;
                foreach (ReportDetailtForTotalMoneyType item in listDataGradation)
                {
                    // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                    arrayGradation[count] = new GradationCompare()
                    {
                        NameGradationCompare = string.Format("AUD năm {0}", item.Year),
                        amount = item.AUD,
                        NameType = item.PartnerName
                    };

                    count++;
                    year = year - 1;
                }
            }
            else
            {
                // Trường hợp thị trường  Châu á
                List<string> listAsianItem = new List<string>();
                foreach (ReportDetailtForTotalMoneyType item in listDataGradation)
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

                List<ReportDetailtForTotalMoneyType> listDSMarketYear = new List<ReportDetailtForTotalMoneyType>();
                List<ReportDetailtForTotalMoneyType> listDSMarketLastYear = new List<ReportDetailtForTotalMoneyType>();

                foreach (string item in listAsianItem)
                {
                    listDSMarketYear = listDataGradation.Where(x => x.MarketName == item && x.Year == year.ToString()).ToList();
                    listDSMarketLastYear = listDataGradation.Where(x => x.MarketName == item && x.Year == (year - 1).ToString()).ToList();

                    double sumAUDYear = listDSMarketYear.Sum(x => x.AUD);
                    double sumAUDLastYear = listDSMarketLastYear.Sum(x => x.AUD);

                    // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                    arrayGradation[count] = new GradationCompare()
                    {
                        NameGradationCompare = string.Format("AUD năm {0}", year),
                        amount = sumAUDYear,
                        NameType = item
                    };
                    count++;

                    // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                    arrayGradation[count] = new GradationCompare()
                    {
                        NameGradationCompare = string.Format("AUD năm {0}", year - 1),
                        amount = sumAUDLastYear,
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
        /// Get data cho việc vẽ biểu đồ cột cho so sánh theo giai đoạn - GBP
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult SearchColumnsChartGradationCompareForOneGBP([DataSourceRequest]DataSourceRequest request, int gradation, int year, string reportTypeID, string marketID)
        {
            List<ReportDetailtForTotalMoneyType> listDataGradation = new ReportBL().ReportDetailtMTGradationCompareForOneConvert(year, gradation, reportTypeID, marketID);

            GradationCompare[] arrayGradation = null;
            if (!marketID.Contains("005"))
            {
                // Số record của mảng
                int countArray = 1;
                arrayGradation = new GradationCompare[countArray * listDataGradation.Count];

                int count = 0;
                foreach (ReportDetailtForTotalMoneyType item in listDataGradation)
                {
                    // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                    arrayGradation[count] = new GradationCompare()
                    {
                        NameGradationCompare = string.Format("GBP năm {0}", item.Year),
                        amount = item.GBP,
                        NameType = item.PartnerName
                    };

                    count++;
                    year = year - 1;
                }
            }
            else
            {
                // Trường hợp thị trường  Châu á
                List<string> listAsianItem = new List<string>();
                foreach (ReportDetailtForTotalMoneyType item in listDataGradation)
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

                List<ReportDetailtForTotalMoneyType> listDSMarketYear = new List<ReportDetailtForTotalMoneyType>();
                List<ReportDetailtForTotalMoneyType> listDSMarketLastYear = new List<ReportDetailtForTotalMoneyType>();

                foreach (string item in listAsianItem)
                {
                    listDSMarketYear = listDataGradation.Where(x => x.MarketName == item && x.Year == year.ToString()).ToList();
                    listDSMarketLastYear = listDataGradation.Where(x => x.MarketName == item && x.Year == (year - 1).ToString()).ToList();

                    double sumGBPYear = listDSMarketYear.Sum(x => x.GBP);
                    double sumGBPLastYear = listDSMarketLastYear.Sum(x => x.GBP);

                    // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                    arrayGradation[count] = new GradationCompare()
                    {
                        NameGradationCompare = string.Format("GBP năm {0}", year),
                        amount = sumGBPYear,
                        NameType = item
                    };
                    count++;

                    // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                    arrayGradation[count] = new GradationCompare()
                    {
                        NameGradationCompare = string.Format("GBP năm {0}", year - 1),
                        amount = sumGBPLastYear,
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
        /// Get data cho báo cáo chi tiết so sánh theo tháng
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult SearchReportDetailtCompareMonthForAll([DataSourceRequest]DataSourceRequest request, int month, int year, string reportTypeID, string marketID)
        {
            List<ReportDetailtForTotalMoneyType> listDataCompareMonth = new ReportBL().ReportDetailtMTCompareMonthForAllConvert(year, month, reportTypeID, marketID);

            foreach (ReportDetailtForTotalMoneyType item in listDataCompareMonth)
            {
                item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;
            }

            // Danh sách mã thị trường của Tất cả
            List<string> listMarket = new List<string>() { "003", "005", "001", "002", "014", "004" };

            // Khởi tạo datatable
            DataTable table = new DataTable();
            // Tạo các cột cho datatable
            table.Columns.Add("MarketName", typeof(String));
            table.Columns.Add("VND1", typeof(double));
            table.Columns.Add("VND2", typeof(double));
            table.Columns.Add("VND3", typeof(double));

            table.Columns.Add("USD1", typeof(double));
            table.Columns.Add("USD2", typeof(double));
            table.Columns.Add("USD3", typeof(double));

            table.Columns.Add("EUR1", typeof(double));
            table.Columns.Add("EUR2", typeof(double));
            table.Columns.Add("EUR3", typeof(double));

            table.Columns.Add("CAD1", typeof(double));
            table.Columns.Add("CAD2", typeof(double));
            table.Columns.Add("CAD3", typeof(double));

            table.Columns.Add("AUD1", typeof(double));
            table.Columns.Add("AUD2", typeof(double));
            table.Columns.Add("AUD3", typeof(double));

            table.Columns.Add("GBP1", typeof(double));
            table.Columns.Add("GBP2", typeof(double));
            table.Columns.Add("GBP3", typeof(double));

            table.Columns.Add("TDS1", typeof(double));
            table.Columns.Add("TDS2", typeof(double));
            table.Columns.Add("TDS3", typeof(double));


            table.Columns.Add("ReportID", typeof(String));

            foreach (ReportDetailtForTotalMoneyType item in listDataCompareMonth)
            {
                ReportDetailtForTotalMoneyType dataItemLastYear = listDataCompareMonth.Find(x => x.MarketCode == item.MarketCode && x.Month == month.ToString() && x.Year == (year - 1).ToString());
                ReportDetailtForTotalMoneyType dataItemYear = listDataCompareMonth.Find(x => x.MarketCode == item.MarketCode && x.Month == month.ToString() && x.Year == year.ToString());
                ReportDetailtForTotalMoneyType dataItemLastMonth = listDataCompareMonth.Find(x => x.MarketCode == item.MarketCode && x.Month == (month - 1).ToString() && x.Year == year.ToString());

                // Trường hợp tháng 1
                if (month == 1)
                {
                    dataItemLastMonth = listDataCompareMonth.Find(x => x.MarketCode == item.MarketCode && x.Month == "12" && x.Year == (year - 1).ToString());
                }

                // Check tồn tại của item
                string value = string.Format("MarketName='{0}'", item.MarketName);
                DataRow[] foundRows = table.Select(value);

                if (dataItemLastYear != null && dataItemYear != null && dataItemLastMonth != null && foundRows.Count() == 0)
                {
                    // add item vào table
                    table.Rows.Add(dataItemYear.MarketName
                        , dataItemLastYear.VND, dataItemLastMonth.VND, dataItemYear.VND
                        , dataItemLastYear.USD, dataItemLastMonth.USD, dataItemYear.USD
                        , dataItemLastYear.EUR, dataItemLastMonth.EUR, dataItemYear.EUR
                        , dataItemLastYear.CAD, dataItemLastMonth.CAD, dataItemYear.CAD
                        , dataItemLastYear.AUD, dataItemLastMonth.AUD, dataItemYear.AUD
                        , dataItemLastYear.GBP, dataItemLastMonth.GBP, dataItemYear.GBP
                        , dataItemLastYear.TongDS, dataItemLastMonth.TongDS, dataItemYear.TongDS
                        );
                }
            }
            // Cùng kì
            

            DataRow row = table.NewRow();
            row["MarketName"] = "Tổng";
            row["VND1"] = table.Compute("Sum(VND1)", "");
            row["VND2"] = table.Compute("Sum(VND2)", "");
            row["VND3"] = table.Compute("Sum(VND3)", "");

            row["USD1"] = table.Compute("Sum(USD1)", "");
            row["USD2"] = table.Compute("Sum(USD2)", "");
            row["USD3"] = table.Compute("Sum(USD3)", "");

            row["EUR1"] = table.Compute("Sum(EUR1)", "");
            row["EUR2"] = table.Compute("Sum(EUR2)", "");
            row["EUR3"] = table.Compute("Sum(EUR3)", "");

            row["CAD1"] = table.Compute("Sum(CAD1)", "");
            row["CAD2"] = table.Compute("Sum(CAD2)", "");
            row["CAD3"] = table.Compute("Sum(CAD3)", "");

            row["AUD1"] = table.Compute("Sum(AUD1)", "");
            row["AUD2"] = table.Compute("Sum(AUD2)", "");
            row["AUD3"] = table.Compute("Sum(AUD3)", "");

            row["GBP1"] = table.Compute("Sum(GBP1)", "");
            row["GBP2"] = table.Compute("Sum(GBP2)", "");
            row["GBP3"] = table.Compute("Sum(GBP3)", "");

            row["TDS1"] = table.Compute("Sum(TDS1)", "");
            row["TDS2"] = table.Compute("Sum(TDS2)", "");
            row["TDS3"] = table.Compute("Sum(TDS3)", "");

            table.Rows.Add(row);

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
            List<ReportDetailtForTotalMoneyType> listDataCompareMonth = new ReportBL().ReportDetailtMTCompareMonthForAllConvert(year, month, reportTypeID, marketID);

            foreach (ReportDetailtForTotalMoneyType item in listDataCompareMonth)
            {
                item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;
            }

            // Khởi tạo datatable
            DataTable table = new DataTable();
            // Tạo các cột cho datatable
            table.Columns.Add("MarketName", typeof(String));

            // So sánh với tháng trước
            table.Columns.Add("VND1", typeof(double));
            table.Columns.Add("VND2", typeof(double));

            table.Columns.Add("USD1", typeof(double));
            table.Columns.Add("USD2", typeof(double));

            table.Columns.Add("EUR1", typeof(double));
            table.Columns.Add("EUR2", typeof(double));

            table.Columns.Add("CAD1", typeof(double));
            table.Columns.Add("CAD2", typeof(double));

            table.Columns.Add("AUD1", typeof(double));
            table.Columns.Add("AUD2", typeof(double));

            table.Columns.Add("GBP1", typeof(double));
            table.Columns.Add("GBP2", typeof(double));

            table.Columns.Add("TDS1", typeof(double));
            table.Columns.Add("TDS2", typeof(double));

            foreach (ReportDetailtForTotalMoneyType item in listDataCompareMonth)
            {
                // Cùng kì
                ReportDetailtForTotalMoneyType dataItemLastYear = listDataCompareMonth.Find(x => x.MarketCode == item.MarketCode && x.Month == month.ToString() && x.Year == (year - 1).ToString());
                ReportDetailtForTotalMoneyType dataItemYear = listDataCompareMonth.Find(x => x.MarketCode == item.MarketCode && x.Month == month.ToString() && x.Year == year.ToString());
                ReportDetailtForTotalMoneyType dataItemLastMonth = listDataCompareMonth.Find(x => x.MarketCode == item.MarketCode && x.Month == (month - 1).ToString() && x.Year == year.ToString());
                // Trường hợp tháng 1
                if (month == 1)
                {
                    dataItemLastMonth = listDataCompareMonth.Find(x => x.MarketCode == item.MarketCode && x.Month == "12" && x.Year == (year - 1).ToString());
                }

                // Cùng kì năm trước
                if (dataItemLastYear == null)
                {
                    dataItemLastYear.MarketCode = item.MarketCode;
                    dataItemLastYear.MarketName = item.MarketName;
                }

                // Tháng hiện tại
                if (dataItemYear == null)
                {
                    dataItemYear.MarketCode = item.MarketCode;
                    dataItemYear.MarketName = item.MarketName;
                }

                // Tháng trước
                if (dataItemLastMonth == null)
                {
                    dataItemLastMonth.MarketCode = item.MarketCode;
                    dataItemLastMonth.MarketName = item.MarketName;
                }

                // So với cùng kì năm trước
                double sumVNDLastYear = Math.Round(dataItemYear.VND - dataItemLastYear.VND, 2, MidpointRounding.ToEven);
                double sumUSDLastYear = Math.Round(dataItemYear.USD - dataItemLastYear.USD, 2, MidpointRounding.ToEven);
                double sumEURLastYear = Math.Round(dataItemYear.EUR - dataItemLastYear.EUR, 2, MidpointRounding.ToEven);
                double sumCADLastYear = Math.Round(dataItemYear.CAD - dataItemLastYear.CAD, 2, MidpointRounding.ToEven);
                double sumAUDLastYear = Math.Round(dataItemYear.AUD - dataItemLastYear.AUD, 2, MidpointRounding.ToEven);
                double sumGBPLastYear = Math.Round(dataItemYear.GBP - dataItemLastYear.GBP, 2, MidpointRounding.ToEven);

                double sumTDSLastYear = sumVNDLastYear + sumUSDLastYear + sumEURLastYear + sumCADLastYear + sumAUDLastYear + sumGBPLastYear;

                // so với tháng trước
                double sumVNDLastMonth = Math.Round(dataItemYear.VND - dataItemLastMonth.VND, 2, MidpointRounding.ToEven);
                double sumUSDLastMonth = Math.Round(dataItemYear.USD - dataItemLastMonth.USD, 2, MidpointRounding.ToEven);
                double sumEURLastMonth = Math.Round(dataItemYear.EUR - dataItemLastMonth.EUR, 2, MidpointRounding.ToEven);
                double sumCADLastMonth = Math.Round(dataItemYear.CAD - dataItemLastMonth.CAD, 2, MidpointRounding.ToEven);
                double sumAUDLastMonth = Math.Round(dataItemYear.AUD - dataItemLastMonth.AUD, 2, MidpointRounding.ToEven);
                double sumGBPLastMonth = Math.Round(dataItemYear.GBP - dataItemLastMonth.GBP, 2, MidpointRounding.ToEven);

                double sumTDSLastMonth = sumVNDLastMonth + sumUSDLastMonth + sumEURLastMonth + sumCADLastMonth + sumAUDLastMonth + sumGBPLastMonth;

                // Check tồn tại của item
                string value = string.Format("MarketName='{0}'", item.MarketName);
                DataRow[] foundRows = table.Select(value);

                if (dataItemLastYear != null && dataItemYear != null && dataItemLastMonth != null && foundRows.Count() == 0)
                {
                    // add item vào table
                    table.Rows.Add(dataItemYear.MarketName
                        , sumVNDLastMonth, sumVNDLastYear
                        , sumUSDLastMonth, sumUSDLastYear
                        , sumEURLastMonth, sumEURLastYear
                        , sumCADLastMonth, sumCADLastYear
                        , sumAUDLastMonth, sumAUDLastYear
                        , sumGBPLastMonth, sumGBPLastYear
                        , Math.Round(sumTDSLastMonth, 2, MidpointRounding.ToEven), Math.Round(sumTDSLastYear, 2, MidpointRounding.ToEven)
                        );
                }
            }

            DataRow row = table.NewRow();
            row["MarketName"] = "Tổng";
            row["VND1"] = table.Compute("Sum(VND1)", "");
            row["VND2"] = table.Compute("Sum(VND2)", "");

            row["USD1"] = table.Compute("Sum(USD1)", "");
            row["USD2"] = table.Compute("Sum(USD2)", "");

            row["EUR1"] = table.Compute("Sum(EUR1)", "");
            row["EUR2"] = table.Compute("Sum(EUR2)", "");

            row["CAD1"] = table.Compute("Sum(CAD1)", "");
            row["CAD2"] = table.Compute("Sum(CAD2)", "");

            row["AUD1"] = table.Compute("Sum(AUD1)", "");
            row["AUD2"] = table.Compute("Sum(AUD2)", "");

            row["GBP1"] = table.Compute("Sum(GBP1)", "");
            row["GBP2"] = table.Compute("Sum(GBP2)", "");

            row["TDS1"] = table.Compute("Sum(TDS1)", "");
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
        public ActionResult SearchColumnChartStackCompareMonthForAll([DataSourceRequest]DataSourceRequest request, int month, int year, string reportTypeID, string marketID)
        {
            List<ReportDetailtForTotalMoneyType> listDataCompareMonth = new ReportBL().ColumnChartStackDetailtMTCompareMonthForAllPercent(year, month, reportTypeID, marketID);
            List<ReportDetailtForTotalMoneyType> listDataCompareMonthConvert = new List<ReportDetailtForTotalMoneyType>();

            foreach (ReportDetailtForTotalMoneyType item in listDataCompareMonth)
            {
                item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;
            }

            GradationCharColumn[] arrayGradation = null;

            if (listDataCompareMonth.Count > 0)
            {
                List<ReportDetailtForTotalMoneyType> listDataYear = listDataCompareMonth.Where(x => x.Month == month.ToString() && x.Year == year.ToString()).ToList();
                List<ReportDetailtForTotalMoneyType> listDataLastMonth = listDataCompareMonth.Where(x => x.Month == (month - 1).ToString() && x.Year == year.ToString()).ToList();
                List<ReportDetailtForTotalMoneyType> listDataLastYear = listDataCompareMonth.Where(x => x.Month == month.ToString() && x.Year == (year - 1).ToString()).ToList();
                // Trường hợp tháng 1
                if (month == 1)
                {
                    listDataLastMonth = listDataCompareMonth.Where(x => x.Month == "12" && x.Year == (year - 1).ToString()).ToList();
                }
                // Trường hợp đủ dữ liệu tháng hiện tại, tháng trước và cùng kì năm trước
                if (listDataYear.Count > 0 && listDataLastMonth.Count > 0 && listDataLastYear.Count > 0)
                {
                    // Tháng hiện tại
                    double sumVNDTotalYear = listDataYear.Sum(x => x.VND);
                    double sumUSDTotalYear = listDataYear.Sum(x => x.USD);
                    double sumEURTotalYear = listDataYear.Sum(x => x.EUR);
                    double sumCADTotalYear = listDataYear.Sum(x => x.CAD);
                    double sumAUDTotalYear = listDataYear.Sum(x => x.AUD);
                    double sumGBPTotalYear = listDataYear.Sum(x => x.GBP);

                    double sumTongDSTotalYear = listDataYear.Sum(x => x.TongDS);

                    // Tháng trước
                    double sumVNDTotalLastMonth = listDataLastMonth.Sum(x => x.VND);
                    double sumUSDTotalLastMonth = listDataLastMonth.Sum(x => x.USD);
                    double sumEURTotalLastMonth = listDataLastMonth.Sum(x => x.EUR);
                    double sumCADTotalLastMonth = listDataLastMonth.Sum(x => x.CAD);
                    double sumAUDTotalLastMonth = listDataLastMonth.Sum(x => x.AUD);
                    double sumGBPTotalLastMonth = listDataLastMonth.Sum(x => x.GBP);

                    double sumTongDSTotalLastMonth = listDataLastMonth.Sum(x => x.TongDS);

                    // Cùng kì năm trước
                    double sumVNDTotalLastYear = listDataLastYear.Sum(x => x.VND);
                    double sumUSDTotalLastYear = listDataLastYear.Sum(x => x.USD);
                    double sumEURTotalLastYear = listDataLastYear.Sum(x => x.EUR);
                    double sumCADTotalLastYear = listDataLastYear.Sum(x => x.CAD);
                    double sumAUDTotalLastYear = listDataLastYear.Sum(x => x.AUD);
                    double sumGBPTotalLastYear = listDataLastYear.Sum(x => x.GBP);

                    double sumTongDSTotalLastYear = listDataLastYear.Sum(x => x.TongDS);

                    foreach (ReportDetailtForTotalMoneyType item in listDataCompareMonth)
                    {
                        // Tháng hiện tại
                        if (item.Month == month.ToString() && item.Year == year.ToString())
                        {
                            listDataCompareMonthConvert.Add(
                                new ReportDetailtForTotalMoneyType()
                                {
                                    MarketCode = item.MarketCode,
                                    MarketName = item.MarketName,
                                    VND = sumVNDTotalYear == 0 ? 0 : Math.Round((item.VND / sumVNDTotalYear) * 100, 2, MidpointRounding.ToEven),
                                    USD = sumUSDTotalYear == 0 ? 0 : Math.Round((item.USD / sumUSDTotalYear) * 100, 2, MidpointRounding.ToEven),
                                    EUR = sumEURTotalYear == 0 ? 0 : Math.Round((item.EUR / sumEURTotalYear) * 100, 2, MidpointRounding.ToEven),
                                    CAD = sumCADTotalYear == 0 ? 0 : Math.Round((item.CAD / sumCADTotalYear) * 100, 2, MidpointRounding.ToEven),
                                    AUD = sumAUDTotalYear == 0 ? 0 : Math.Round((item.AUD / sumAUDTotalYear) * 100, 2, MidpointRounding.ToEven),
                                    GBP = sumGBPTotalYear == 0 ? 0 : Math.Round((item.GBP / sumGBPTotalYear) * 100, 2, MidpointRounding.ToEven),
                                    Month = month.ToString(),
                                    Year = year.ToString()
                                }
                            );
                        }

                        // Tháng trước
                        if (month == 1)
                        {
                            if (item.Month == "12" && item.Year == (year -1).ToString())
                            {
                                listDataCompareMonthConvert.Add(
                                    new ReportDetailtForTotalMoneyType()
                                    {
                                        MarketCode = item.MarketCode,
                                        MarketName = item.MarketName,
                                        VND = sumVNDTotalLastMonth == 0 ? 0 : Math.Round((item.VND / sumVNDTotalLastMonth) * 100, 2, MidpointRounding.ToEven),
                                        USD = sumUSDTotalLastMonth == 0 ? 0 : Math.Round((item.USD / sumUSDTotalLastMonth) * 100, 2, MidpointRounding.ToEven),
                                        EUR = sumEURTotalLastMonth == 0 ? 0 : Math.Round((item.EUR / sumEURTotalLastMonth) * 100, 2, MidpointRounding.ToEven),
                                        CAD = sumCADTotalLastMonth == 0 ? 0 : Math.Round((item.CAD / sumCADTotalLastMonth) * 100, 2, MidpointRounding.ToEven),
                                        AUD = sumAUDTotalLastMonth == 0 ? 0 : Math.Round((item.AUD / sumAUDTotalLastMonth) * 100, 2, MidpointRounding.ToEven),
                                        GBP = sumGBPTotalLastMonth == 0 ? 0 : Math.Round((item.GBP / sumGBPTotalLastMonth) * 100, 2, MidpointRounding.ToEven),
                                        Month = "12",
                                        Year = (year - 1).ToString()
                                    }
                                );
                            }
                        }
                        else
                        {
                            if (item.Month == (month - 1).ToString() && item.Year == year.ToString())
                            {
                                listDataCompareMonthConvert.Add(
                                    new ReportDetailtForTotalMoneyType()
                                    {
                                        MarketCode = item.MarketCode,
                                        MarketName = item.MarketName,
                                        VND = sumVNDTotalLastMonth == 0 ? 0 : Math.Round((item.VND / sumVNDTotalLastMonth) * 100, 2, MidpointRounding.ToEven),
                                        USD = sumUSDTotalLastMonth == 0 ? 0 : Math.Round((item.USD / sumUSDTotalLastMonth) * 100, 2, MidpointRounding.ToEven),
                                        EUR = sumEURTotalLastMonth == 0 ? 0 : Math.Round((item.EUR / sumEURTotalLastMonth) * 100, 2, MidpointRounding.ToEven),
                                        CAD = sumCADTotalLastMonth == 0 ? 0 : Math.Round((item.CAD / sumCADTotalLastMonth) * 100, 2, MidpointRounding.ToEven),
                                        AUD = sumAUDTotalLastMonth == 0 ? 0 : Math.Round((item.AUD / sumAUDTotalLastMonth) * 100, 2, MidpointRounding.ToEven),
                                        GBP = sumGBPTotalLastMonth == 0 ? 0 : Math.Round((item.GBP / sumGBPTotalLastMonth) * 100, 2, MidpointRounding.ToEven),
                                        Month = (month - 1).ToString(),
                                        Year = year.ToString()
                                    }
                                );
                            }
                        }
                        

                        // Cùng kì năm trước
                        if (item.Month == month.ToString() && item.Year == (year - 1).ToString())
                        {
                            listDataCompareMonthConvert.Add(
                                new ReportDetailtForTotalMoneyType()
                                {
                                    MarketCode = item.MarketCode,
                                    MarketName = item.MarketName,
                                    VND = sumVNDTotalLastYear == 0 ? 0 : Math.Round((item.VND / sumVNDTotalLastYear) * 100, 2, MidpointRounding.ToEven),
                                    USD = sumUSDTotalLastYear == 0 ? 0 : Math.Round((item.USD / sumUSDTotalLastYear) * 100, 2, MidpointRounding.ToEven),
                                    EUR = sumEURTotalLastYear == 0 ? 0 : Math.Round((item.EUR / sumEURTotalLastYear) * 100, 2, MidpointRounding.ToEven),
                                    CAD = sumCADTotalLastYear == 0 ? 0 : Math.Round((item.CAD / sumCADTotalLastYear) * 100, 2, MidpointRounding.ToEven),
                                    AUD = sumAUDTotalLastYear == 0 ? 0 : Math.Round((item.AUD / sumAUDTotalLastYear) * 100, 2, MidpointRounding.ToEven),
                                    GBP = sumGBPTotalLastYear == 0 ? 0 : Math.Round((item.GBP / sumGBPTotalLastYear) * 100, 2, MidpointRounding.ToEven),
                                    Month = month.ToString(),
                                    Year = (year - 1).ToString()
                                }
                            );
                        }
                    }
                    
                    // Số record của mảng
                    int countArray = 6;
                    arrayGradation = new GradationCharColumn[countArray * listDataCompareMonthConvert.Count];
                    int count = 0;
                    int tooltip = 1;

                    foreach (ReportDetailtForTotalMoneyType item in listDataCompareMonthConvert)
                    {
                        // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                        arrayGradation[count] = new GradationCharColumn()
                        {
                            Serie = item.MarketName,
                            Segmento = string.Format("VND Tháng {0}/{1}", item.Month, item.Year),
                            Valor1 = item.VND,
                            Tooltip = tooltip
                        };

                        count++;

                        arrayGradation[count] = new GradationCharColumn()
                        {
                            Serie = item.MarketName,
                            Segmento = string.Format("USD Tháng {0}/{1}", item.Month, item.Year),
                            Valor1 = item.USD,
                            Tooltip = tooltip
                        };

                        count++;

                        arrayGradation[count] = new GradationCharColumn()
                        {
                            Serie = item.MarketName,
                            Segmento = string.Format("EUR Tháng {0}/{1}", item.Month, item.Year),
                            Valor1 = item.EUR,
                            Tooltip = tooltip
                        };

                        count++;

                        arrayGradation[count] = new GradationCharColumn()
                        {
                            Serie = item.MarketName,
                            Segmento = string.Format("CAD Tháng {0}/{1}", item.Month, item.Year),
                            Valor1 = item.CAD,
                            Tooltip = tooltip
                        };

                        count++;

                        arrayGradation[count] = new GradationCharColumn()
                        {
                            Serie = item.MarketName,
                            Segmento = string.Format("AUD Tháng {0}/{1}", item.Month, item.Year),
                            Valor1 = item.AUD,
                            Tooltip = tooltip
                        };

                        count++;

                        arrayGradation[count] = new GradationCharColumn()
                        {
                            Serie = item.MarketName,
                            Segmento = string.Format("GBP Tháng {0}/{1}", item.Month, item.Year),
                            Valor1 = item.GBP,
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
        /// Get data cho việc vẽ biểu đồ cột cho so sánh giai đoạn - VND
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult SearchColumnsChartCompareMonthForAllVND([DataSourceRequest]DataSourceRequest request, int month, int year, string reportTypeID, string marketID)
        {
            List<ReportDetailtForTotalMoneyType> listDataCompareMonth = new ReportBL().ReportDetailtMTCompareMonthForAllConvert(year, month, reportTypeID, marketID);

            GradationCompare[] arrayGradation = null;
            if (listDataCompareMonth.Count > 0)
            {
                List<ReportDetailtForTotalMoneyType> listDataYear = listDataCompareMonth.Where(x => x.Month == month.ToString() && x.Year == year.ToString()).ToList();
                List<ReportDetailtForTotalMoneyType> listDataLastMonth = listDataCompareMonth.Where(x => x.Month == (month - 1).ToString() && x.Year == year.ToString()).ToList();
                List<ReportDetailtForTotalMoneyType> listDataLastYear = listDataCompareMonth.Where(x => x.Month == month.ToString() && x.Year == (year - 1).ToString()).ToList();
                // Trường hợp tháng 1
                if (month == 1)
                {
                    listDataLastMonth = listDataCompareMonth.Where(x => x.Month == "12" && x.Year == (year - 1).ToString()).ToList();
                }

                // Trường hợp đủ dữ liệu tháng hiện tại, tháng trước và cùng kì năm trước
                if (listDataYear.Count > 0 && listDataLastMonth.Count > 0 && listDataLastYear.Count > 0)
                {
                    // Số record của mảng
                    int countArray = 1;
                    arrayGradation = new GradationCompare[countArray * listDataCompareMonth.Count];
                    int count = 0;

                    foreach (ReportDetailtForTotalMoneyType item in listDataCompareMonth)
                    {
                        // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                        arrayGradation[count] = new GradationCompare()
                        {
                            NameGradationCompare = string.Format("VND Tháng {0}/{1}", item.Month, item.Year),
                            amount = item.VND,
                            NameType = item.MarketName
                        };

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
        /// Get data cho việc vẽ biểu đồ cột cho so sánh giai đoạn - USD
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult SearchColumnsChartCompareMonthForAllUSD([DataSourceRequest]DataSourceRequest request, int month, int year, string reportTypeID, string marketID)
        {
            List<ReportDetailtForTotalMoneyType> listDataCompareMonth = new ReportBL().ReportDetailtMTCompareMonthForAllConvert(year, month, reportTypeID, marketID);

            GradationCompare[] arrayGradation = null;
            if (listDataCompareMonth.Count > 0)
            {
                List<ReportDetailtForTotalMoneyType> listDataYear = listDataCompareMonth.Where(x => x.Month == month.ToString() && x.Year == year.ToString()).ToList();
                List<ReportDetailtForTotalMoneyType> listDataLastMonth = listDataCompareMonth.Where(x => x.Month == (month - 1).ToString() && x.Year == year.ToString()).ToList();
                List<ReportDetailtForTotalMoneyType> listDataLastYear = listDataCompareMonth.Where(x => x.Month == month.ToString() && x.Year == (year - 1).ToString()).ToList();
                // Trường hợp tháng 1
                if (month == 1)
                {
                    listDataLastMonth = listDataCompareMonth.Where(x => x.Month == "12" && x.Year == (year - 1).ToString()).ToList();
                }

                // Trường hợp đủ dữ liệu tháng hiện tại, tháng trước và cùng kì năm trước
                if (listDataYear.Count > 0 && listDataLastMonth.Count > 0 && listDataLastYear.Count > 0)
                {
                    // Số record của mảng
                    int countArray = 1;
                    arrayGradation = new GradationCompare[countArray * listDataCompareMonth.Count];
                    int count = 0;

                    foreach (ReportDetailtForTotalMoneyType item in listDataCompareMonth)
                    {
                        // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                        arrayGradation[count] = new GradationCompare()
                        {
                            NameGradationCompare = string.Format("USD Tháng {0}/{1}", item.Month, item.Year),
                            amount = item.USD,
                            NameType = item.MarketName
                        };

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
        /// Get data cho việc vẽ biểu đồ cột cho so sánh giai đoạn - EUR
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult SearchColumnsChartCompareMonthForAllEUR([DataSourceRequest]DataSourceRequest request, int month, int year, string reportTypeID, string marketID)
        {
            List<ReportDetailtForTotalMoneyType> listDataCompareMonth = new ReportBL().ReportDetailtMTCompareMonthForAllConvert(year, month, reportTypeID, marketID);

            GradationCompare[] arrayGradation = null;
            if (listDataCompareMonth.Count > 0)
            {
                List<ReportDetailtForTotalMoneyType> listDataYear = listDataCompareMonth.Where(x => x.Month == month.ToString() && x.Year == year.ToString()).ToList();
                List<ReportDetailtForTotalMoneyType> listDataLastMonth = listDataCompareMonth.Where(x => x.Month == (month - 1).ToString() && x.Year == year.ToString()).ToList();
                List<ReportDetailtForTotalMoneyType> listDataLastYear = listDataCompareMonth.Where(x => x.Month == month.ToString() && x.Year == (year - 1).ToString()).ToList();
                // Trường hợp tháng 1
                if (month == 1)
                {
                    listDataLastMonth = listDataCompareMonth.Where(x => x.Month == "12" && x.Year == (year - 1).ToString()).ToList();
                }

                // Trường hợp đủ dữ liệu tháng hiện tại, tháng trước và cùng kì năm trước
                if (listDataYear.Count > 0 && listDataLastMonth.Count > 0 && listDataLastYear.Count > 0)
                {
                    // Số record của mảng
                    int countArray = 1;
                    arrayGradation = new GradationCompare[countArray * listDataCompareMonth.Count];
                    int count = 0;

                    foreach (ReportDetailtForTotalMoneyType item in listDataCompareMonth)
                    {
                        // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                        arrayGradation[count] = new GradationCompare()
                        {
                            NameGradationCompare = string.Format("EUR Tháng {0}/{1}", item.Month, item.Year),
                            amount = item.EUR,
                            NameType = item.MarketName
                        };

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
        /// Get data cho việc vẽ biểu đồ cột cho so sánh giai đoạn - CAD
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult SearchColumnsChartCompareMonthForAllCAD([DataSourceRequest]DataSourceRequest request, int month, int year, string reportTypeID, string marketID)
        {
            List<ReportDetailtForTotalMoneyType> listDataCompareMonth = new ReportBL().ReportDetailtMTCompareMonthForAllConvert(year, month, reportTypeID, marketID);

            GradationCompare[] arrayGradation = null;
            if (listDataCompareMonth.Count > 0)
            {
                List<ReportDetailtForTotalMoneyType> listDataYear = listDataCompareMonth.Where(x => x.Month == month.ToString() && x.Year == year.ToString()).ToList();
                List<ReportDetailtForTotalMoneyType> listDataLastMonth = listDataCompareMonth.Where(x => x.Month == (month - 1).ToString() && x.Year == year.ToString()).ToList();
                List<ReportDetailtForTotalMoneyType> listDataLastYear = listDataCompareMonth.Where(x => x.Month == month.ToString() && x.Year == (year - 1).ToString()).ToList();
                // Trường hợp tháng 1
                if (month == 1)
                {
                    listDataLastMonth = listDataCompareMonth.Where(x => x.Month == "12" && x.Year == (year - 1).ToString()).ToList();
                }

                // Trường hợp đủ dữ liệu tháng hiện tại, tháng trước và cùng kì năm trước
                if (listDataYear.Count > 0 && listDataLastMonth.Count > 0 && listDataLastYear.Count > 0)
                {
                    // Số record của mảng
                    int countArray = 1;
                    arrayGradation = new GradationCompare[countArray * listDataCompareMonth.Count];
                    int count = 0;

                    foreach (ReportDetailtForTotalMoneyType item in listDataCompareMonth)
                    {
                        // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                        arrayGradation[count] = new GradationCompare()
                        {
                            NameGradationCompare = string.Format("CAD Tháng {0}/{1}", item.Month, item.Year),
                            amount = item.CAD,
                            NameType = item.MarketName
                        };

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
        /// Get data cho việc vẽ biểu đồ cột cho so sánh giai đoạn - AUD
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult SearchColumnsChartCompareMonthForAllAUD([DataSourceRequest]DataSourceRequest request, int month, int year, string reportTypeID, string marketID)
        {
            List<ReportDetailtForTotalMoneyType> listDataCompareMonth = new ReportBL().ReportDetailtMTCompareMonthForAllConvert(year, month, reportTypeID, marketID);

            GradationCompare[] arrayGradation = null;
            if (listDataCompareMonth.Count > 0)
            {
                List<ReportDetailtForTotalMoneyType> listDataYear = listDataCompareMonth.Where(x => x.Month == month.ToString() && x.Year == year.ToString()).ToList();
                List<ReportDetailtForTotalMoneyType> listDataLastMonth = listDataCompareMonth.Where(x => x.Month == (month - 1).ToString() && x.Year == year.ToString()).ToList();
                List<ReportDetailtForTotalMoneyType> listDataLastYear = listDataCompareMonth.Where(x => x.Month == month.ToString() && x.Year == (year - 1).ToString()).ToList();
                // Trường hợp tháng 1
                if (month == 1)
                {
                    listDataLastMonth = listDataCompareMonth.Where(x => x.Month == "12" && x.Year == (year - 1).ToString()).ToList();
                }

                // Trường hợp đủ dữ liệu tháng hiện tại, tháng trước và cùng kì năm trước
                if (listDataYear.Count > 0 && listDataLastMonth.Count > 0 && listDataLastYear.Count > 0)
                {
                    // Số record của mảng
                    int countArray = 1;
                    arrayGradation = new GradationCompare[countArray * listDataCompareMonth.Count];
                    int count = 0;

                    foreach (ReportDetailtForTotalMoneyType item in listDataCompareMonth)
                    {
                        // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                        arrayGradation[count] = new GradationCompare()
                        {
                            NameGradationCompare = string.Format("AUD Tháng {0}/{1}", item.Month, item.Year),
                            amount = item.AUD,
                            NameType = item.MarketName
                        };

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
        /// Get data cho việc vẽ biểu đồ cột cho so sánh giai đoạn - GBP
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult SearchColumnsChartCompareMonthForAllGBP([DataSourceRequest]DataSourceRequest request, int month, int year, string reportTypeID, string marketID)
        {
            List<ReportDetailtForTotalMoneyType> listDataCompareMonth = new ReportBL().ReportDetailtMTCompareMonthForAllConvert(year, month, reportTypeID, marketID);

            GradationCompare[] arrayGradation = null;
            if (listDataCompareMonth.Count > 0)
            {
                List<ReportDetailtForTotalMoneyType> listDataYear = listDataCompareMonth.Where(x => x.Month == month.ToString() && x.Year == year.ToString()).ToList();
                List<ReportDetailtForTotalMoneyType> listDataLastMonth = listDataCompareMonth.Where(x => x.Month == (month - 1).ToString() && x.Year == year.ToString()).ToList();
                List<ReportDetailtForTotalMoneyType> listDataLastYear = listDataCompareMonth.Where(x => x.Month == month.ToString() && x.Year == (year - 1).ToString()).ToList();
                // Trường hợp tháng 1
                if (month == 1)
                {
                    listDataLastMonth = listDataCompareMonth.Where(x => x.Month == "12" && x.Year == (year - 1).ToString()).ToList();
                }

                // Trường hợp đủ dữ liệu tháng hiện tại, tháng trước và cùng kì năm trước
                if (listDataYear.Count > 0 && listDataLastMonth.Count > 0 && listDataLastYear.Count > 0)
                {
                    // Số record của mảng
                    int countArray = 1;
                    arrayGradation = new GradationCompare[countArray * listDataCompareMonth.Count];
                    int count = 0;

                    foreach (ReportDetailtForTotalMoneyType item in listDataCompareMonth)
                    {
                        // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                        arrayGradation[count] = new GradationCompare()
                        {
                            NameGradationCompare = string.Format("GBP Tháng {0}/{1}", item.Month, item.Year),
                            amount = item.GBP,
                            NameType = item.MarketName
                        };

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
        public ActionResult SearchReportDetailtCompareMonthForOne([DataSourceRequest]DataSourceRequest request, int month, int year, string reportTypeID, string marketID)
        {
            List<ReportDetailtForTotalMoneyType> listDataCompareMonth = new ReportBL().ReportDetailtMTCompareMonthForOneConvert(year, month, reportTypeID, marketID);
            
            // Khởi tạo datatable
            DataTable table = new DataTable();
            // Tạo các cột cho datatable
            table.Columns.Add("PartnerName", typeof(String));
            table.Columns.Add("VND1", typeof(double));
            table.Columns.Add("VND2", typeof(double));
            table.Columns.Add("VND3", typeof(double));

            table.Columns.Add("USD1", typeof(double));
            table.Columns.Add("USD2", typeof(double));
            table.Columns.Add("USD3", typeof(double));

            table.Columns.Add("EUR1", typeof(double));
            table.Columns.Add("EUR2", typeof(double));
            table.Columns.Add("EUR3", typeof(double));

            table.Columns.Add("CAD1", typeof(double));
            table.Columns.Add("CAD2", typeof(double));
            table.Columns.Add("CAD3", typeof(double));

            table.Columns.Add("AUD1", typeof(double));
            table.Columns.Add("AUD2", typeof(double));
            table.Columns.Add("AUD3", typeof(double));

            table.Columns.Add("GBP1", typeof(double));
            table.Columns.Add("GBP2", typeof(double));
            table.Columns.Add("GBP3", typeof(double));

            table.Columns.Add("TDS1", typeof(double));
            table.Columns.Add("TDS2", typeof(double));
            table.Columns.Add("TDS3", typeof(double));

            table.Columns.Add("MarketName", typeof(String));
            
            List<string> listTemp = new List<string>();

            if (!marketID.Contains("005"))
            {
                foreach (ReportDetailtForTotalMoneyType item in listDataCompareMonth)
                {
                    item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;
                }

                foreach (ReportDetailtForTotalMoneyType item in listDataCompareMonth)
                {
                    // Cùng kì
                    ReportDetailtForTotalMoneyType dataItemLastYear = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == month.ToString() && x.Year == (year - 1).ToString());
                    ReportDetailtForTotalMoneyType dataItemYear = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == month.ToString() && x.Year == year.ToString());
                    ReportDetailtForTotalMoneyType dataItemLastMonth = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == (month - 1).ToString() && x.Year == year.ToString());
                    // Trường hợp tháng 1
                    if (month == 1)
                    {
                        dataItemLastMonth = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == "12" && x.Year == (year - 1).ToString());
                    }

                    if (!listTemp.Contains(item.PartnerCode))
                    {
                        // Trường hợp năm trước có đối tác không có
                        if (dataItemLastYear == null)
                        {
                            dataItemLastYear = new ReportDetailtForTotalMoneyType();
                            dataItemLastYear.MarketCode = item.MarketCode;
                            dataItemLastYear.MarketName = item.MarketName;
                            dataItemLastYear.PartnerName = item.PartnerName;
                            dataItemLastYear.Year = (year - 1).ToString();
                            dataItemLastYear.Month = month.ToString();
                        }

                        // Trường hợp năm có đối tác không có
                        if (dataItemYear == null)
                        {
                            dataItemYear = new ReportDetailtForTotalMoneyType();
                            dataItemYear.MarketCode = item.MarketCode;
                            dataItemYear.MarketName = item.MarketName;
                            dataItemYear.PartnerName = item.PartnerName;
                            dataItemYear.Year = year.ToString();
                            dataItemYear.Month = month.ToString();
                        }

                        // Trường hợp năm có tháng trước có đối tác không có
                        if (dataItemLastMonth == null)
                        {
                            dataItemLastMonth = new ReportDetailtForTotalMoneyType();
                            dataItemLastMonth.MarketCode = item.MarketCode;
                            dataItemLastMonth.MarketName = item.MarketName;
                            dataItemLastMonth.PartnerName = item.PartnerName;
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
                                , dataItemYear.VND, dataItemLastMonth.VND, dataItemLastYear.VND
                                , dataItemYear.USD, dataItemLastMonth.USD, dataItemLastYear.USD
                                , dataItemYear.EUR, dataItemLastMonth.EUR, dataItemLastYear.EUR
                                , dataItemYear.CAD, dataItemLastMonth.CAD, dataItemLastYear.CAD
                                , dataItemYear.AUD, dataItemLastMonth.AUD, dataItemLastYear.AUD
                                , dataItemYear.GBP, dataItemLastMonth.GBP, dataItemLastYear.GBP
                                , dataItemYear.TongDS, dataItemLastMonth.TongDS, dataItemLastYear.TongDS
                                , item.MarketName
                                );
                        }
                    }
                }
            }
            else
            {
                foreach (ReportDetailtForTotalMoneyType item in listDataCompareMonth)
                {
                    if(!listTemp.Contains(item.MarketName))
                    {
                        listTemp.Add(item.MarketName);
                    }
                }
                
                foreach (string item in listTemp)
                {
                    List<ReportDetailtForTotalMoneyType> dataItemLastYear = listDataCompareMonth.Where(x => x.MarketName == item && x.Month == month.ToString() && x.Year == (year - 1).ToString()).ToList();
                    List<ReportDetailtForTotalMoneyType> dataItemYear = listDataCompareMonth.Where(x => x.MarketName == item && x.Month == month.ToString() && x.Year == year.ToString()).ToList();
                    List<ReportDetailtForTotalMoneyType> dataItemLastMonth = listDataCompareMonth.Where(x => x.MarketName == item && x.Month == (month - 1).ToString() && x.Year == year.ToString()).ToList();
                    // Trường hợp tháng 1
                    if (month == 1)
                    {
                        dataItemLastMonth = listDataCompareMonth.Where(x => x.MarketName == item && x.Month == "12" && x.Year == (year - 1).ToString()).ToList();
                    }

                    // Cùng kì
                    if (dataItemLastYear.Count == 0)
                    {
                        dataItemLastYear = new List<ReportDetailtForTotalMoneyType>()
                        {
                            new ReportDetailtForTotalMoneyType()
                            {
                                MarketName = item,
                                Month = month.ToString(),
                                Year = (year - 1).ToString()
                            }
                        };
                    }

                    // Tháng hiện tại
                    if (dataItemYear.Count == 0)
                    {
                        dataItemYear = new List<ReportDetailtForTotalMoneyType>()
                        {
                            new ReportDetailtForTotalMoneyType()
                            {
                                MarketName = item,
                                Month = month.ToString(),
                                Year = year.ToString()
                            }
                        };
                    }

                    // Tháng trước
                    if (dataItemLastMonth.Count == 0)
                    {
                        if (month == 1)
                        {
                            dataItemLastMonth = new List<ReportDetailtForTotalMoneyType>()
                            {
                                new ReportDetailtForTotalMoneyType()
                                {
                                    MarketName = item,
                                    Month = "12",
                                    Year = (year - 1).ToString()
                                }
                            };
                        }
                        else
                        {

                            dataItemLastMonth = new List<ReportDetailtForTotalMoneyType>()
                            {
                                new ReportDetailtForTotalMoneyType()
                                {
                                    MarketName = item,
                                    Month = (month -1).ToString(),
                                    Year = year.ToString()
                                }
                            };
                        }
                    }

                    // Cùng kì năm trước
                    double sumVNDLastYear = dataItemLastYear.Sum(x => x.VND);
                    double sumUSDLastYear = dataItemLastYear.Sum(x => x.USD);
                    double sumEURLastYear = dataItemLastYear.Sum(x => x.EUR);
                    double sumCADLastYear = dataItemLastYear.Sum(x => x.CAD);
                    double sumAUDLastYear = dataItemLastYear.Sum(x => x.AUD);
                    double sumGBPLastYear = dataItemLastYear.Sum(x => x.GBP);

                    double sumTongDSLastYear = dataItemLastYear.Sum(x => x.TongDS);

                    // Tháng hiện tại
                    double sumVNDYear = dataItemYear.Sum(x => x.VND);
                    double sumUSDYear = dataItemYear.Sum(x => x.USD);
                    double sumEURYear = dataItemYear.Sum(x => x.EUR);
                    double sumCADYear = dataItemYear.Sum(x => x.CAD);
                    double sumAUDYear = dataItemYear.Sum(x => x.AUD);
                    double sumGBPYear = dataItemYear.Sum(x => x.GBP);

                    double sumTongDSYear = dataItemYear.Sum(x => x.TongDS);

                    // Tháng trước
                    double sumVNDLastMonth = dataItemLastMonth.Sum(x => x.VND);
                    double sumUSDLastMonth = dataItemLastMonth.Sum(x => x.USD);
                    double sumEURLastMonth = dataItemLastMonth.Sum(x => x.EUR);
                    double sumCADLastMonth = dataItemLastMonth.Sum(x => x.CAD);
                    double sumAUDLastMonth = dataItemLastMonth.Sum(x => x.AUD);
                    double sumGBPLastMonth = dataItemLastMonth.Sum(x => x.GBP);

                    double sumTongDSLastMonth = dataItemLastMonth.Sum(x => x.TongDS);

                    table.Rows.Add(
                        item
                        , sumVNDYear, sumVNDLastMonth, sumVNDLastYear
                        , sumUSDYear, sumUSDLastMonth, sumUSDLastYear
                        , sumEURYear, sumEURLastMonth, sumEURLastYear
                        , sumCADYear, sumCADLastMonth, sumCADLastYear
                        , sumAUDYear, sumAUDLastMonth, sumAUDLastYear
                        , sumGBPYear, sumGBPLastMonth, sumGBPLastYear
                        , sumTongDSYear, sumTongDSLastMonth, sumTongDSLastYear
                        , "Thị trường Châu Á"
                    );
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
            List<ReportDetailtForTotalMoneyType> listDataCompareMonth = new ReportBL().ReportDetailtMTCompareMonthForOneConvert(year, month, reportTypeID, marketID);

            foreach (ReportDetailtForTotalMoneyType item in listDataCompareMonth)
            {
                item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;
            }

            // Khởi tạo datatable
            DataTable table = new DataTable();
            // Tạo các cột cho datatable
            // So sánh với tháng trước
            table.Columns.Add("PartnerName", typeof(String));
            table.Columns.Add("VND1", typeof(double));
            table.Columns.Add("VND2", typeof(double));

            table.Columns.Add("USD1", typeof(double));
            table.Columns.Add("USD2", typeof(double));

            table.Columns.Add("EUR1", typeof(double));
            table.Columns.Add("EUR2", typeof(double));

            table.Columns.Add("CAD1", typeof(double));
            table.Columns.Add("CAD2", typeof(double));

            table.Columns.Add("AUD1", typeof(double));
            table.Columns.Add("AUD2", typeof(double));

            table.Columns.Add("GBP1", typeof(double));
            table.Columns.Add("GBP2", typeof(double));

            table.Columns.Add("TDS1", typeof(double));
            table.Columns.Add("TDS2", typeof(double));

            table.Columns.Add("MarketName", typeof(String));

            List<string> listTemp = new List<string>();

            if (!marketID.Contains("005"))
            {
                foreach (ReportDetailtForTotalMoneyType item in listDataCompareMonth)
                {
                    // Cùng kì
                    ReportDetailtForTotalMoneyType dataItemLastYear = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == month.ToString() && x.Year == (year - 1).ToString());
                    ReportDetailtForTotalMoneyType dataItemYear = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == month.ToString() && x.Year == year.ToString());
                    ReportDetailtForTotalMoneyType dataItemLastMonth = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == (month - 1).ToString() && x.Year == year.ToString());
                    // Trường hợp tháng 1
                    if (month == 1)
                    {
                        dataItemLastMonth = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == "12" && x.Year == (year - 1).ToString());
                    }

                    if (!listTemp.Contains(item.PartnerCode))
                    {
                        // Trường hợp năm trước không có đối tác
                        if (dataItemLastYear == null)
                        {
                            dataItemLastYear = new ReportDetailtForTotalMoneyType();
                            dataItemLastYear.MarketCode = item.MarketCode;
                            dataItemLastYear.MarketName = item.MarketName;
                            dataItemLastYear.PartnerName = item.PartnerName;
                            dataItemLastYear.Year = (year - 1).ToString();
                            dataItemLastYear.Month = month.ToString();
                        }

                        // Trường hợp năm hiện tại không có đối tác
                        if (dataItemYear == null)
                        {
                            dataItemYear = new ReportDetailtForTotalMoneyType();
                            dataItemYear.MarketCode = item.MarketCode;
                            dataItemYear.MarketName = item.MarketName;
                            dataItemYear.PartnerName = item.PartnerName;
                            dataItemYear.Year = year.ToString();
                            dataItemYear.Month = month.ToString();
                        }

                        // Trường hợp tháng trước không có
                        if (dataItemLastMonth == null)
                        {
                            dataItemLastMonth = new ReportDetailtForTotalMoneyType();
                            dataItemLastMonth.MarketCode = item.MarketCode;
                            dataItemLastMonth.MarketName = item.MarketName;
                            dataItemLastMonth.PartnerName = item.PartnerName;
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

                        // so với tháng trước
                        double sumVND = dataItemYear.VND - dataItemLastMonth.VND;
                        double sumUSD = dataItemYear.USD - dataItemLastMonth.USD;
                        double sumEUR = dataItemYear.EUR - dataItemLastMonth.EUR;
                        double sumCAD = dataItemYear.CAD - dataItemLastMonth.CAD;
                        double sumAUD = dataItemYear.AUD - dataItemLastMonth.AUD;
                        double sumGBP = dataItemYear.GBP - dataItemLastMonth.GBP;

                        double sumTongDS = dataItemYear.TongDS - dataItemLastMonth.TongDS;

                        // so với cùng kì năm trước
                        double sumVNDLastYear = dataItemYear.VND - dataItemLastYear.VND;
                        double sumUSDLastYear = dataItemYear.USD - dataItemLastYear.USD;
                        double sumEURLastYear = dataItemYear.EUR - dataItemLastYear.EUR;
                        double sumCADLastYear = dataItemYear.CAD - dataItemLastYear.CAD;
                        double sumAUDLastYear = dataItemYear.AUD - dataItemLastYear.AUD;
                        double sumGBPLastYear = dataItemYear.GBP - dataItemLastYear.GBP;

                        double sumTongDSLastYear = dataItemYear.TongDS - dataItemLastYear.TongDS;

                        // Check tồn tại của item
                        string value = string.Format("PartnerName='{0}'", item.PartnerName);
                        DataRow[] foundRows = table.Select(value);

                        if (dataItemLastYear != null && dataItemYear != null && dataItemLastMonth != null && foundRows.Count() == 0)
                        {
                            // add item vào table
                            table.Rows.Add(item.PartnerName
                                // So với tháng hiện tại
                                , Math.Round(sumVND, 2, MidpointRounding.ToEven)
                                , Math.Round(sumVNDLastYear, 2, MidpointRounding.ToEven)

                                , Math.Round(sumUSD, 2, MidpointRounding.ToEven)
                                , Math.Round(sumUSDLastYear, 2, MidpointRounding.ToEven)

                                , Math.Round(sumEUR, 2, MidpointRounding.ToEven)
                                , Math.Round(sumEURLastYear, 2, MidpointRounding.ToEven)

                                , Math.Round(sumCAD, 2, MidpointRounding.ToEven)
                                , Math.Round(sumCADLastYear, 2, MidpointRounding.ToEven)

                                , Math.Round(sumAUD, 2, MidpointRounding.ToEven)
                                , Math.Round(sumAUDLastYear, 2, MidpointRounding.ToEven)

                                , Math.Round(sumGBP, 2, MidpointRounding.ToEven)
                                , Math.Round(sumGBPLastYear, 2, MidpointRounding.ToEven)

                                , Math.Round(sumTongDS, 2, MidpointRounding.ToEven)
                                , Math.Round(sumTongDSLastYear, 2, MidpointRounding.ToEven)
                                , item.MarketName
                                );
                        }

                        // Add partnerCode để kiểm tra
                        listTemp.Add(item.PartnerCode);
                    }
                }
            }
            else
            {
                foreach (ReportDetailtForTotalMoneyType item in listDataCompareMonth)
                {
                    if (!listTemp.Contains(item.MarketName))
                    {
                        listTemp.Add(item.MarketName);
                    }
                }

                foreach (string item in listTemp)
                {
                    List<ReportDetailtForTotalMoneyType> dataItemLastYear = listDataCompareMonth.Where(x => x.MarketName == item && x.Month == month.ToString() && x.Year == (year - 1).ToString()).ToList();
                    List<ReportDetailtForTotalMoneyType> dataItemYear = listDataCompareMonth.Where(x => x.MarketName == item && x.Month == month.ToString() && x.Year == year.ToString()).ToList();
                    List<ReportDetailtForTotalMoneyType> dataItemLastMonth = listDataCompareMonth.Where(x => x.MarketName == item && x.Month == (month - 1).ToString() && x.Year == year.ToString()).ToList();
                    // Trường hợp tháng 1
                    if (month == 1)
                    {
                        dataItemLastMonth = listDataCompareMonth.Where(x => x.MarketName == item && x.Month == "12" && x.Year == (year - 1).ToString()).ToList();
                    }
                    // Cùng kì
                    if (dataItemLastYear.Count == 0)
                    {
                        dataItemLastYear = new List<ReportDetailtForTotalMoneyType>()
                        {
                            new ReportDetailtForTotalMoneyType()
                            {
                                MarketName = item,
                                Month = month.ToString(),
                                Year = (year - 1).ToString()
                            }
                        };
                    }

                    // Tháng hiện tại
                    if (dataItemYear.Count == 0)
                    {
                        dataItemYear = new List<ReportDetailtForTotalMoneyType>()
                        {
                            new ReportDetailtForTotalMoneyType()
                            {
                                MarketName = item,
                                Month = month.ToString(),
                                Year = year.ToString()
                            }
                        };
                    }

                    // Tháng trước
                    if (dataItemLastMonth.Count == 0)
                    {
                        // Trường hợp tháng 1
                        if (month == 1)
                        {
                            dataItemLastMonth = new List<ReportDetailtForTotalMoneyType>()
                            {
                                new ReportDetailtForTotalMoneyType()
                                {
                                    MarketName = item,
                                    Month = "12",
                                    Year = (year - 1).ToString()
                                }
                            };
                        }
                        else
                        {
                            dataItemLastMonth = new List<ReportDetailtForTotalMoneyType>()
                            {
                                new ReportDetailtForTotalMoneyType()
                                {
                                    MarketName = item,
                                    Month = (month -1).ToString(),
                                    Year = year.ToString()
                                }
                            };
                        }
                        
                    }

                    // Cùng kì năm trước
                    double sumVNDLastYear = dataItemLastYear.Sum(x => x.VND);
                    double sumUSDLastYear = dataItemLastYear.Sum(x => x.USD);
                    double sumEURLastYear = dataItemLastYear.Sum(x => x.EUR);
                    double sumCADLastYear = dataItemLastYear.Sum(x => x.CAD);
                    double sumAUDLastYear = dataItemLastYear.Sum(x => x.AUD);
                    double sumGBPLastYear = dataItemLastYear.Sum(x => x.GBP);

                    double sumTongDSLastYear = dataItemLastYear.Sum(x => x.TongDS);

                    // Tháng hiện tại
                    double sumVNDYear = dataItemYear.Sum(x => x.VND);
                    double sumUSDYear = dataItemYear.Sum(x => x.USD);
                    double sumEURYear = dataItemYear.Sum(x => x.EUR);
                    double sumCADYear = dataItemYear.Sum(x => x.CAD);
                    double sumAUDYear = dataItemYear.Sum(x => x.AUD);
                    double sumGBPYear = dataItemYear.Sum(x => x.GBP);

                    double sumTongDSYear = dataItemYear.Sum(x => x.TongDS);

                    // Tháng trước
                    double sumVNDLastMonth = dataItemLastMonth.Sum(x => x.VND);
                    double sumUSDLastMonth = dataItemLastMonth.Sum(x => x.USD);
                    double sumEURLastMonth = dataItemLastMonth.Sum(x => x.EUR);
                    double sumCADLastMonth = dataItemLastMonth.Sum(x => x.CAD);
                    double sumAUDLastMonth = dataItemLastMonth.Sum(x => x.AUD);
                    double sumGBPLastMonth = dataItemLastMonth.Sum(x => x.GBP);

                    double sumTongDSLastMonth = dataItemLastMonth.Sum(x => x.TongDS);

                    table.Rows.Add(
                        item
                        , Math.Round(sumVNDYear - sumVNDLastMonth, 2, MidpointRounding.ToEven), Math.Round(sumVNDYear - sumVNDLastYear, 2, MidpointRounding.ToEven)
                        , Math.Round(sumUSDYear - sumUSDLastMonth, 2, MidpointRounding.ToEven), Math.Round(sumUSDYear - sumUSDLastYear, 2, MidpointRounding.ToEven)
                        , Math.Round(sumEURYear - sumEURLastMonth, 2, MidpointRounding.ToEven), Math.Round(sumEURYear - sumEURLastYear, 2, MidpointRounding.ToEven)
                        , Math.Round(sumCADYear - sumCADLastMonth, 2, MidpointRounding.ToEven), Math.Round(sumCADYear - sumCADLastYear, 2, MidpointRounding.ToEven)
                        , Math.Round(sumAUDYear - sumAUDLastMonth, 2, MidpointRounding.ToEven), Math.Round(sumAUDYear - sumAUDLastYear, 2, MidpointRounding.ToEven)
                        , Math.Round(sumGBPYear - sumGBPLastMonth, 2, MidpointRounding.ToEven), Math.Round(sumGBPYear - sumGBPLastYear, 2, MidpointRounding.ToEven)
                        , Math.Round(sumTongDSYear - sumTongDSLastMonth, 2, MidpointRounding.ToEven), Math.Round(sumTongDSYear - sumTongDSLastYear, 2, MidpointRounding.ToEven)
                        , "Thị trường Châu Á"
                    );
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
            List<ReportDetailtForTotalMoneyType> listDataCompareMonth = new ReportBL().ReportDetailtMTCompareMonthForOneConvertPercent(year, month, reportTypeID, marketID);
            List<ReportDetailtForTotalMoneyType> listDataCompareMonthConvert = new List<ReportDetailtForTotalMoneyType>();
            List<ReportDetailtForTotalMoneyType> listDataCommpareMonthClone = new List<ReportDetailtForTotalMoneyType>(listDataCompareMonth);

            int arrayCount = 6;
            // Trường hợp là thị trường chi tiết của loại tiền theo tháng của Châu Á
            if (marketID.Contains("005"))
            {
                // Danh sách các thị trường con của Châu Á
                List<string> listMarket = new List<string>();
                foreach (ReportDetailtForTotalMoneyType item in listDataCompareMonth)
                {
                    if (!listMarket.Contains(item.MarketName))
                    {
                        listMarket.Add(item.MarketName);
                    }
                }

                foreach(string item in listMarket)
                {
                    List<ReportDetailtForTotalMoneyType> dataItemLastYear = listDataCompareMonth.Where(x => x.MarketName == item && x.Month == month.ToString() && x.Year == (year - 1).ToString()).ToList();
                    List<ReportDetailtForTotalMoneyType> dataItemYear = listDataCompareMonth.Where(x => x.MarketName == item && x.Month == month.ToString() && x.Year == year.ToString()).ToList();
                    List<ReportDetailtForTotalMoneyType> dataItemLastMonth = listDataCompareMonth.Where(x => x.MarketName == item && x.Month == (month - 1).ToString() && x.Year == year.ToString()).ToList();
                    // Trường hợp tháng 1
                    if (month == 1)
                    {
                        dataItemLastMonth = listDataCompareMonth.Where(x => x.MarketName == item && x.Month == "12" && x.Year == (year - 1).ToString()).ToList();
                    }
                    // Cùng kì
                    if (dataItemLastYear.Count == 0)
                    {
                        dataItemLastYear = new List<ReportDetailtForTotalMoneyType>()
                        {
                            new ReportDetailtForTotalMoneyType()
                            {
                                MarketName = item,
                                Month = month.ToString(),
                                Year = (year - 1).ToString()
                            }
                        };
                    }

                    // Tháng hiện tại
                    if (dataItemYear.Count == 0)
                    {
                        dataItemYear = new List<ReportDetailtForTotalMoneyType>()
                        {
                            new ReportDetailtForTotalMoneyType()
                            {
                                MarketName = item,
                                Month = month.ToString(),
                                Year = year.ToString()
                            }
                        };
                    }

                    // Tháng trước
                    if (dataItemLastMonth.Count == 0)
                    {// Trường hợp tháng 1
                        if (month == 1)
                        {
                            dataItemLastMonth = new List<ReportDetailtForTotalMoneyType>()
                            {
                                new ReportDetailtForTotalMoneyType()
                                {
                                    MarketName = item,
                                    Month = "12",
                                    Year = (year - 1).ToString()
                                }
                            };
                        }
                        else
                        {
                            dataItemLastMonth = new List<ReportDetailtForTotalMoneyType>()
                            {
                                new ReportDetailtForTotalMoneyType()
                                {
                                    MarketName = item,
                                    Month = (month -1).ToString(),
                                    Year = year.ToString()
                                }
                            };
                        }
                        
                    }

                    // Cùng kì năm trước
                    double sumVNDLastYear = dataItemLastYear.Sum(x => x.VND);
                    double sumUSDLastYear = dataItemLastYear.Sum(x => x.USD);
                    double sumEURLastYear = dataItemLastYear.Sum(x => x.EUR);
                    double sumCADLastYear = dataItemLastYear.Sum(x => x.CAD);
                    double sumAUDLastYear = dataItemLastYear.Sum(x => x.AUD);
                    double sumGBPLastYear = dataItemLastYear.Sum(x => x.GBP);

                    double sumTongDSLastYear = dataItemLastYear.Sum(x => x.TongDS);

                    // Tháng hiện tại
                    double sumVNDYear = dataItemYear.Sum(x => x.VND);
                    double sumUSDYear = dataItemYear.Sum(x => x.USD);
                    double sumEURYear = dataItemYear.Sum(x => x.EUR);
                    double sumCADYear = dataItemYear.Sum(x => x.CAD);
                    double sumAUDYear = dataItemYear.Sum(x => x.AUD);
                    double sumGBPYear = dataItemYear.Sum(x => x.GBP);

                    double sumTongDSYear = dataItemYear.Sum(x => x.TongDS);

                    // Tháng trước
                    double sumVNDLastMonth = dataItemLastMonth.Sum(x => x.VND);
                    double sumUSDLastMonth = dataItemLastMonth.Sum(x => x.USD);
                    double sumEURLastMonth = dataItemLastMonth.Sum(x => x.EUR);
                    double sumCADLastMonth = dataItemLastMonth.Sum(x => x.CAD);
                    double sumAUDLastMonth = dataItemLastMonth.Sum(x => x.AUD);
                    double sumGBPLastMonth = dataItemLastMonth.Sum(x => x.GBP);

                    double sumTongDSLastMonth = dataItemLastMonth.Sum(x => x.TongDS);

                    // Vì  get  dữ liệu theo partnerName nên thị trường gán cho partner
                    // Tháng hiện tại
                    listDataCompareMonthConvert.Add(
                        new ReportDetailtForTotalMoneyType()
                        {
                            PartnerName = item,
                            Month = month.ToString(),
                            Year = year.ToString(),
                            VND = sumVNDYear,
                            USD = sumUSDYear,
                            EUR = sumEURYear,
                            CAD = sumCADYear,
                            AUD = sumAUDYear,
                            GBP = sumGBPYear,
                        }
                    );

                    // Tháng Trước
                    // Trường hợp tháng 1
                    if (month == 1)
                    {
                        listDataCompareMonthConvert.Add(
                            new ReportDetailtForTotalMoneyType()
                            {
                                PartnerName = item,
                                Month = "12",
                                Year = (year - 1).ToString(),
                                VND = sumVNDLastMonth,
                                USD = sumUSDLastMonth,
                                EUR = sumEURLastMonth,
                                CAD = sumCADLastMonth,
                                AUD = sumAUDLastMonth,
                                GBP = sumGBPLastMonth,
                            }
                        );
                    }
                    else
                    {
                        listDataCompareMonthConvert.Add(
                            new ReportDetailtForTotalMoneyType()
                            {
                                PartnerName = item,
                                Month = (month - 1).ToString(),
                                Year = year.ToString(),
                                VND = sumVNDLastMonth,
                                USD = sumUSDLastMonth,
                                EUR = sumEURLastMonth,
                                CAD = sumCADLastMonth,
                                AUD = sumAUDLastMonth,
                                GBP = sumGBPLastMonth,
                            }
                        );
                    }
                    

                    // Cùng kì năm trước
                    listDataCompareMonthConvert.Add(
                        new ReportDetailtForTotalMoneyType()
                        {
                            PartnerName = item,
                            Month = month.ToString(),
                            Year = (year - 1).ToString(),
                            VND = sumVNDLastYear,
                            USD = sumUSDLastYear,
                            EUR = sumEURLastYear,
                            CAD = sumCADLastYear,
                            AUD = sumAUDLastYear,
                            GBP = sumGBPLastYear,
                        }
                    );
                }
            }
            else
            {
                listDataCompareMonthConvert = new List<ReportDetailtForTotalMoneyType>(listDataCompareMonth);
            }
            
            GradationCharColumn[] arrayGradation = new GradationCharColumn[listDataCompareMonthConvert.Count * arrayCount];
            int count = 0;

            foreach (ReportDetailtForTotalMoneyType item in listDataCompareMonthConvert)
            {
                // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                arrayGradation[count] = new GradationCharColumn()
                {
                    Serie = string.Format("VND {0}/{1}", item.Month, item.Year),
                    Segmento = item.PartnerName,
                    Valor1 = item.VND,
                    //Tooltip = tooltip
                };
                count++;

                arrayGradation[count] = new GradationCharColumn()
                {
                    Serie = string.Format("USD {0}/{1}", item.Month, item.Year),
                    Segmento = item.PartnerName,
                    Valor1 = item.USD,
                    //Tooltip = tooltip
                };
                count++;

                arrayGradation[count] = new GradationCharColumn()
                {
                    Serie = string.Format("EUR {0}/{1}", item.Month, item.Year),
                    Segmento = item.PartnerName,
                    Valor1 = item.EUR,
                    //Tooltip = tooltip
                };
                count++;

                arrayGradation[count] = new GradationCharColumn()
                {
                    Serie = string.Format("CAD {0}/{1}", item.Month, item.Year),
                    Segmento = item.PartnerName,
                    Valor1 = item.CAD,
                    //Tooltip = tooltip
                };
                count++;

                arrayGradation[count] = new GradationCharColumn()
                {
                    Serie = string.Format("AUD {0}/{1}", item.Month, item.Year),
                    Segmento = item.PartnerName,
                    Valor1 = item.AUD,
                    //Tooltip = tooltip
                };
                count++;

                arrayGradation[count] = new GradationCharColumn()
                {
                    Serie = string.Format("GBP {0}/{1}", item.Month, item.Year),
                    Segmento = item.PartnerName,
                    Valor1 = item.GBP,
                    //Tooltip = tooltip
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
        public ActionResult SearchColumnsChartCompareMonthForOneVND([DataSourceRequest]DataSourceRequest request, int month, int year, string reportTypeID, string marketID)
        {
            List<ReportDetailtForTotalMoneyType> listDataCompareMonth = new ReportBL().ReportDetailtMTCompareMonthForOneConvert(year, month, reportTypeID, marketID);
            
            List<ReportDetailtForTotalMoneyType> listDataCompareMonthConvert = new List<ReportDetailtForTotalMoneyType>();
            // Trường hợp là thị trường chi tiết của loại tiền theo tháng của Châu Á
            if (marketID.Contains("005"))
            {
                // Danh sách các thị trường con của Châu Á
                List<string> listMarket = new List<string>();
                foreach (ReportDetailtForTotalMoneyType item in listDataCompareMonth)
                {
                    if (!listMarket.Contains(item.MarketName))
                    {
                        listMarket.Add(item.MarketName);
                    }
                }

                foreach (string item in listMarket)
                {
                    List<ReportDetailtForTotalMoneyType> dataItemLastYear = listDataCompareMonth.Where(x => x.MarketName == item && x.Month == month.ToString() && x.Year == (year - 1).ToString()).ToList();
                    List<ReportDetailtForTotalMoneyType> dataItemYear = listDataCompareMonth.Where(x => x.MarketName == item && x.Month == month.ToString() && x.Year == year.ToString()).ToList();
                    List<ReportDetailtForTotalMoneyType> dataItemLastMonth = listDataCompareMonth.Where(x => x.MarketName == item && x.Month == (month - 1).ToString() && x.Year == year.ToString()).ToList();
                    // Trường hợp tháng 1
                    if (month == 1)
                    {
                        dataItemLastMonth = listDataCompareMonth.Where(x => x.MarketName == item && x.Month == "12" && x.Year == (year - 1).ToString()).ToList();
                    }
                    // Cùng kì
                    if (dataItemLastYear.Count == 0)
                    {
                        dataItemLastYear = new List<ReportDetailtForTotalMoneyType>()
                        {
                            new ReportDetailtForTotalMoneyType()
                            {
                                MarketName = item,
                                Month = month.ToString(),
                                Year = (year - 1).ToString()
                            }
                        };
                    }

                    // Tháng hiện tại
                    if (dataItemYear.Count == 0)
                    {
                        dataItemYear = new List<ReportDetailtForTotalMoneyType>()
                        {
                            new ReportDetailtForTotalMoneyType()
                            {
                                MarketName = item,
                                Month = month.ToString(),
                                Year = year.ToString()
                            }
                        };
                    }

                    // Tháng trước
                    if (dataItemLastMonth.Count == 0)
                    {
                        // Trường hợp tháng 1
                        if (month == 1)
                        {
                            dataItemLastMonth = new List<ReportDetailtForTotalMoneyType>()
                            {
                                new ReportDetailtForTotalMoneyType()
                                {
                                    MarketName = item,
                                    Month = "12",
                                    Year = (year -1).ToString()
                                }
                            };
                        }
                        else
                        {
                            dataItemLastMonth = new List<ReportDetailtForTotalMoneyType>()
                            {
                                new ReportDetailtForTotalMoneyType()
                                {
                                    MarketName = item,
                                    Month = (month -1).ToString(),
                                    Year = year.ToString()
                                }
                            };
                        }
                    }

                    // Cùng kì năm trước
                    double sumVNDLastYear = dataItemLastYear.Sum(x => x.VND);
                    double sumUSDLastYear = dataItemLastYear.Sum(x => x.USD);
                    double sumEURLastYear = dataItemLastYear.Sum(x => x.EUR);
                    double sumCADLastYear = dataItemLastYear.Sum(x => x.CAD);
                    double sumAUDLastYear = dataItemLastYear.Sum(x => x.AUD);
                    double sumGBPLastYear = dataItemLastYear.Sum(x => x.GBP);

                    double sumTongDSLastYear = dataItemLastYear.Sum(x => x.TongDS);

                    // Tháng hiện tại
                    double sumVNDYear = dataItemYear.Sum(x => x.VND);
                    double sumUSDYear = dataItemYear.Sum(x => x.USD);
                    double sumEURYear = dataItemYear.Sum(x => x.EUR);
                    double sumCADYear = dataItemYear.Sum(x => x.CAD);
                    double sumAUDYear = dataItemYear.Sum(x => x.AUD);
                    double sumGBPYear = dataItemYear.Sum(x => x.GBP);

                    double sumTongDSYear = dataItemYear.Sum(x => x.TongDS);

                    // Tháng trước
                    double sumVNDLastMonth = dataItemLastMonth.Sum(x => x.VND);
                    double sumUSDLastMonth = dataItemLastMonth.Sum(x => x.USD);
                    double sumEURLastMonth = dataItemLastMonth.Sum(x => x.EUR);
                    double sumCADLastMonth = dataItemLastMonth.Sum(x => x.CAD);
                    double sumAUDLastMonth = dataItemLastMonth.Sum(x => x.AUD);
                    double sumGBPLastMonth = dataItemLastMonth.Sum(x => x.GBP);

                    double sumTongDSLastMonth = dataItemLastMonth.Sum(x => x.TongDS);

                    // Vì  get  dữ liệu theo partnerName nên thị trường gán cho partner
                    // Tháng hiện tại
                    listDataCompareMonthConvert.Add(
                        new ReportDetailtForTotalMoneyType()
                        {
                            PartnerName = item,
                            Month = month.ToString(),
                            Year = year.ToString(),
                            VND = sumVNDYear,
                            USD = sumUSDYear,
                            EUR = sumEURYear,
                            CAD = sumCADYear,
                            AUD = sumAUDYear,
                            GBP = sumGBPYear,
                        }
                    );

                    // Tháng Trước// Trường hợp tháng 1
                    if (month == 1)
                    {
                        listDataCompareMonthConvert.Add(
                            new ReportDetailtForTotalMoneyType()
                            {
                                PartnerName = item,
                                Month = "12",
                                Year = (year -1).ToString(),
                                VND = sumVNDLastMonth,
                                USD = sumUSDLastMonth,
                                EUR = sumEURLastMonth,
                                CAD = sumCADLastMonth,
                                AUD = sumAUDLastMonth,
                                GBP = sumGBPLastMonth,
                            }
                        );
                    }
                    else
                    {
                        listDataCompareMonthConvert.Add(
                            new ReportDetailtForTotalMoneyType()
                            {
                                PartnerName = item,
                                Month = (month - 1).ToString(),
                                Year = year.ToString(),
                                VND = sumVNDLastMonth,
                                USD = sumUSDLastMonth,
                                EUR = sumEURLastMonth,
                                CAD = sumCADLastMonth,
                                AUD = sumAUDLastMonth,
                                GBP = sumGBPLastMonth,
                            }
                        );
                    }
                    

                    // Cùng kì năm trước
                    listDataCompareMonthConvert.Add(
                        new ReportDetailtForTotalMoneyType()
                        {
                            PartnerName = item,
                            Month = month.ToString(),
                            Year = (year - 1).ToString(),
                            VND = sumVNDLastYear,
                            USD = sumUSDLastYear,
                            EUR = sumEURLastYear,
                            CAD = sumCADLastYear,
                            AUD = sumAUDLastYear,
                            GBP = sumGBPLastYear,
                        }
                    );
                }
            }
            else
            {
                listDataCompareMonthConvert = new List<ReportDetailtForTotalMoneyType>(listDataCompareMonth);
            }

            // Số record của mảng
            int countArray = 1;
            GradationCompare[] arrayGradation = new GradationCompare[countArray * listDataCompareMonthConvert.Count];
            int count = 0;
            foreach (ReportDetailtForTotalMoneyType item in listDataCompareMonthConvert)
            {
                // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                arrayGradation[count] = new GradationCompare()
                {
                    NameGradationCompare = string.Format("VND Tháng {0}/{1}", item.Month, item.Year),
                    amount = item.VND,
                    NameType = item.PartnerName
                };

                count++;
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
        public ActionResult SearchColumnsChartCompareMonthForOneUSD([DataSourceRequest]DataSourceRequest request, int month, int year, string reportTypeID, string marketID)
        {
            List<ReportDetailtForTotalMoneyType> listDataCompareMonth = new ReportBL().ReportDetailtMTCompareMonthForOneConvert(year, month, reportTypeID, marketID);

            List<ReportDetailtForTotalMoneyType> listDataCompareMonthConvert = new List<ReportDetailtForTotalMoneyType>();
            // Trường hợp là thị trường chi tiết của loại tiền theo tháng của Châu Á
            if (marketID.Contains("005"))
            {
                // Danh sách các thị trường con của Châu Á
                List<string> listMarket = new List<string>();
                foreach (ReportDetailtForTotalMoneyType item in listDataCompareMonth)
                {
                    if (!listMarket.Contains(item.MarketName))
                    {
                        listMarket.Add(item.MarketName);
                    }
                }

                foreach (string item in listMarket)
                {
                    List<ReportDetailtForTotalMoneyType> dataItemLastYear = listDataCompareMonth.Where(x => x.MarketName == item && x.Month == month.ToString() && x.Year == (year - 1).ToString()).ToList();
                    List<ReportDetailtForTotalMoneyType> dataItemYear = listDataCompareMonth.Where(x => x.MarketName == item && x.Month == month.ToString() && x.Year == year.ToString()).ToList();
                    List<ReportDetailtForTotalMoneyType> dataItemLastMonth = listDataCompareMonth.Where(x => x.MarketName == item && x.Month == (month - 1).ToString() && x.Year == year.ToString()).ToList();
                    // Trường hợp tháng 1
                    if (month == 1)
                    {
                        dataItemLastMonth = listDataCompareMonth.Where(x => x.MarketName == item && x.Month == "12" && x.Year == (year - 1).ToString()).ToList();
                    }
                    // Cùng kì
                    if (dataItemLastYear.Count == 0)
                    {
                        dataItemLastYear = new List<ReportDetailtForTotalMoneyType>()
                        {
                            new ReportDetailtForTotalMoneyType()
                            {
                                MarketName = item,
                                Month = month.ToString(),
                                Year = (year - 1).ToString()
                            }
                        };
                    }

                    // Tháng hiện tại
                    if (dataItemYear.Count == 0)
                    {
                        dataItemYear = new List<ReportDetailtForTotalMoneyType>()
                        {
                            new ReportDetailtForTotalMoneyType()
                            {
                                MarketName = item,
                                Month = month.ToString(),
                                Year = year.ToString()
                            }
                        };
                    }

                    // Tháng trước
                    if (dataItemLastMonth.Count == 0)
                    {
                        // Trường hợp tháng 1
                        if (month == 1)
                        {
                            dataItemLastMonth = new List<ReportDetailtForTotalMoneyType>()
                            {
                                new ReportDetailtForTotalMoneyType()
                                {
                                    MarketName = item,
                                    Month = "12",
                                    Year = (year -1).ToString()
                                }
                            };
                        }
                        else
                        {
                            dataItemLastMonth = new List<ReportDetailtForTotalMoneyType>()
                            {
                                new ReportDetailtForTotalMoneyType()
                                {
                                    MarketName = item,
                                    Month = (month -1).ToString(),
                                    Year = year.ToString()
                                }
                            };
                        }
                    }

                    // Cùng kì năm trước
                    double sumVNDLastYear = dataItemLastYear.Sum(x => x.VND);
                    double sumUSDLastYear = dataItemLastYear.Sum(x => x.USD);
                    double sumEURLastYear = dataItemLastYear.Sum(x => x.EUR);
                    double sumCADLastYear = dataItemLastYear.Sum(x => x.CAD);
                    double sumAUDLastYear = dataItemLastYear.Sum(x => x.AUD);
                    double sumGBPLastYear = dataItemLastYear.Sum(x => x.GBP);

                    double sumTongDSLastYear = dataItemLastYear.Sum(x => x.TongDS);

                    // Tháng hiện tại
                    double sumVNDYear = dataItemYear.Sum(x => x.VND);
                    double sumUSDYear = dataItemYear.Sum(x => x.USD);
                    double sumEURYear = dataItemYear.Sum(x => x.EUR);
                    double sumCADYear = dataItemYear.Sum(x => x.CAD);
                    double sumAUDYear = dataItemYear.Sum(x => x.AUD);
                    double sumGBPYear = dataItemYear.Sum(x => x.GBP);

                    double sumTongDSYear = dataItemYear.Sum(x => x.TongDS);

                    // Tháng trước
                    double sumVNDLastMonth = dataItemLastMonth.Sum(x => x.VND);
                    double sumUSDLastMonth = dataItemLastMonth.Sum(x => x.USD);
                    double sumEURLastMonth = dataItemLastMonth.Sum(x => x.EUR);
                    double sumCADLastMonth = dataItemLastMonth.Sum(x => x.CAD);
                    double sumAUDLastMonth = dataItemLastMonth.Sum(x => x.AUD);
                    double sumGBPLastMonth = dataItemLastMonth.Sum(x => x.GBP);

                    double sumTongDSLastMonth = dataItemLastMonth.Sum(x => x.TongDS);

                    // Vì  get  dữ liệu theo partnerName nên thị trường gán cho partner
                    // Tháng hiện tại
                    listDataCompareMonthConvert.Add(
                        new ReportDetailtForTotalMoneyType()
                        {
                            PartnerName = item,
                            Month = month.ToString(),
                            Year = year.ToString(),
                            VND = sumVNDYear,
                            USD = sumUSDYear,
                            EUR = sumEURYear,
                            CAD = sumCADYear,
                            AUD = sumAUDYear,
                            GBP = sumGBPYear,
                        }
                    );

                    // Tháng Trước// Trường hợp tháng 1
                    if (month == 1)
                    {
                        listDataCompareMonthConvert.Add(
                            new ReportDetailtForTotalMoneyType()
                            {
                                PartnerName = item,
                                Month = "12",
                                Year = (year - 1).ToString(),
                                VND = sumVNDLastMonth,
                                USD = sumUSDLastMonth,
                                EUR = sumEURLastMonth,
                                CAD = sumCADLastMonth,
                                AUD = sumAUDLastMonth,
                                GBP = sumGBPLastMonth,
                            }
                        );
                    }
                    else
                    {
                        listDataCompareMonthConvert.Add(
                            new ReportDetailtForTotalMoneyType()
                            {
                                PartnerName = item,
                                Month = (month - 1).ToString(),
                                Year = year.ToString(),
                                VND = sumVNDLastMonth,
                                USD = sumUSDLastMonth,
                                EUR = sumEURLastMonth,
                                CAD = sumCADLastMonth,
                                AUD = sumAUDLastMonth,
                                GBP = sumGBPLastMonth,
                            }
                        );
                    }


                    // Cùng kì năm trước
                    listDataCompareMonthConvert.Add(
                        new ReportDetailtForTotalMoneyType()
                        {
                            PartnerName = item,
                            Month = month.ToString(),
                            Year = (year - 1).ToString(),
                            VND = sumVNDLastYear,
                            USD = sumUSDLastYear,
                            EUR = sumEURLastYear,
                            CAD = sumCADLastYear,
                            AUD = sumAUDLastYear,
                            GBP = sumGBPLastYear,
                        }
                    );
                }
            }
            else
            {
                listDataCompareMonthConvert = new List<ReportDetailtForTotalMoneyType>(listDataCompareMonth);
            }

            // Số record của mảng
            int countArray = 1;
            GradationCompare[] arrayGradation = new GradationCompare[countArray * listDataCompareMonthConvert.Count];
            int count = 0;
            foreach (ReportDetailtForTotalMoneyType item in listDataCompareMonthConvert)
            {
                // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                arrayGradation[count] = new GradationCompare()
                {
                    NameGradationCompare = string.Format("USD Tháng {0}/{1}", item.Month, item.Year),
                    amount = item.USD,
                    NameType = item.PartnerName
                };

                count++;
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
        public ActionResult SearchColumnsChartCompareMonthForOneEUR([DataSourceRequest]DataSourceRequest request, int month, int year, string reportTypeID, string marketID)
        {
            List<ReportDetailtForTotalMoneyType> listDataCompareMonth = new ReportBL().ReportDetailtMTCompareMonthForOneConvert(year, month, reportTypeID, marketID);

            List<ReportDetailtForTotalMoneyType> listDataCompareMonthConvert = new List<ReportDetailtForTotalMoneyType>();
            // Trường hợp là thị trường chi tiết của loại tiền theo tháng của Châu Á
            if (marketID.Contains("005"))
            {
                // Danh sách các thị trường con của Châu Á
                List<string> listMarket = new List<string>();
                foreach (ReportDetailtForTotalMoneyType item in listDataCompareMonth)
                {
                    if (!listMarket.Contains(item.MarketName))
                    {
                        listMarket.Add(item.MarketName);
                    }
                }

                foreach (string item in listMarket)
                {
                    List<ReportDetailtForTotalMoneyType> dataItemLastYear = listDataCompareMonth.Where(x => x.MarketName == item && x.Month == month.ToString() && x.Year == (year - 1).ToString()).ToList();
                    List<ReportDetailtForTotalMoneyType> dataItemYear = listDataCompareMonth.Where(x => x.MarketName == item && x.Month == month.ToString() && x.Year == year.ToString()).ToList();
                    List<ReportDetailtForTotalMoneyType> dataItemLastMonth = listDataCompareMonth.Where(x => x.MarketName == item && x.Month == (month - 1).ToString() && x.Year == year.ToString()).ToList();
                    // Trường hợp tháng 1
                    if (month == 1)
                    {
                        dataItemLastMonth = listDataCompareMonth.Where(x => x.MarketName == item && x.Month == "12" && x.Year == (year - 1).ToString()).ToList();
                    }
                    // Cùng kì
                    if (dataItemLastYear.Count == 0)
                    {
                        dataItemLastYear = new List<ReportDetailtForTotalMoneyType>()
                        {
                            new ReportDetailtForTotalMoneyType()
                            {
                                MarketName = item,
                                Month = month.ToString(),
                                Year = (year - 1).ToString()
                            }
                        };
                    }

                    // Tháng hiện tại
                    if (dataItemYear.Count == 0)
                    {
                        dataItemYear = new List<ReportDetailtForTotalMoneyType>()
                        {
                            new ReportDetailtForTotalMoneyType()
                            {
                                MarketName = item,
                                Month = month.ToString(),
                                Year = year.ToString()
                            }
                        };
                    }

                    // Tháng trước
                    if (dataItemLastMonth.Count == 0)
                    {
                        // Trường hợp tháng 1
                        if (month == 1)
                        {
                            dataItemLastMonth = new List<ReportDetailtForTotalMoneyType>()
                            {
                                new ReportDetailtForTotalMoneyType()
                                {
                                    MarketName = item,
                                    Month = "12",
                                    Year = (year -1).ToString()
                                }
                            };
                        }
                        else
                        {
                            dataItemLastMonth = new List<ReportDetailtForTotalMoneyType>()
                            {
                                new ReportDetailtForTotalMoneyType()
                                {
                                    MarketName = item,
                                    Month = (month -1).ToString(),
                                    Year = year.ToString()
                                }
                            };
                        }
                    }

                    // Cùng kì năm trước
                    double sumVNDLastYear = dataItemLastYear.Sum(x => x.VND);
                    double sumUSDLastYear = dataItemLastYear.Sum(x => x.USD);
                    double sumEURLastYear = dataItemLastYear.Sum(x => x.EUR);
                    double sumCADLastYear = dataItemLastYear.Sum(x => x.CAD);
                    double sumAUDLastYear = dataItemLastYear.Sum(x => x.AUD);
                    double sumGBPLastYear = dataItemLastYear.Sum(x => x.GBP);

                    double sumTongDSLastYear = dataItemLastYear.Sum(x => x.TongDS);

                    // Tháng hiện tại
                    double sumVNDYear = dataItemYear.Sum(x => x.VND);
                    double sumUSDYear = dataItemYear.Sum(x => x.USD);
                    double sumEURYear = dataItemYear.Sum(x => x.EUR);
                    double sumCADYear = dataItemYear.Sum(x => x.CAD);
                    double sumAUDYear = dataItemYear.Sum(x => x.AUD);
                    double sumGBPYear = dataItemYear.Sum(x => x.GBP);

                    double sumTongDSYear = dataItemYear.Sum(x => x.TongDS);

                    // Tháng trước
                    double sumVNDLastMonth = dataItemLastMonth.Sum(x => x.VND);
                    double sumUSDLastMonth = dataItemLastMonth.Sum(x => x.USD);
                    double sumEURLastMonth = dataItemLastMonth.Sum(x => x.EUR);
                    double sumCADLastMonth = dataItemLastMonth.Sum(x => x.CAD);
                    double sumAUDLastMonth = dataItemLastMonth.Sum(x => x.AUD);
                    double sumGBPLastMonth = dataItemLastMonth.Sum(x => x.GBP);

                    double sumTongDSLastMonth = dataItemLastMonth.Sum(x => x.TongDS);

                    // Vì  get  dữ liệu theo partnerName nên thị trường gán cho partner
                    // Tháng hiện tại
                    listDataCompareMonthConvert.Add(
                        new ReportDetailtForTotalMoneyType()
                        {
                            PartnerName = item,
                            Month = month.ToString(),
                            Year = year.ToString(),
                            VND = sumVNDYear,
                            USD = sumUSDYear,
                            EUR = sumEURYear,
                            CAD = sumCADYear,
                            AUD = sumAUDYear,
                            GBP = sumGBPYear,
                        }
                    );

                    // Tháng Trước// Trường hợp tháng 1
                    if (month == 1)
                    {
                        listDataCompareMonthConvert.Add(
                            new ReportDetailtForTotalMoneyType()
                            {
                                PartnerName = item,
                                Month = "12",
                                Year = (year - 1).ToString(),
                                VND = sumVNDLastMonth,
                                USD = sumUSDLastMonth,
                                EUR = sumEURLastMonth,
                                CAD = sumCADLastMonth,
                                AUD = sumAUDLastMonth,
                                GBP = sumGBPLastMonth,
                            }
                        );
                    }
                    else
                    {
                        listDataCompareMonthConvert.Add(
                            new ReportDetailtForTotalMoneyType()
                            {
                                PartnerName = item,
                                Month = (month - 1).ToString(),
                                Year = year.ToString(),
                                VND = sumVNDLastMonth,
                                USD = sumUSDLastMonth,
                                EUR = sumEURLastMonth,
                                CAD = sumCADLastMonth,
                                AUD = sumAUDLastMonth,
                                GBP = sumGBPLastMonth,
                            }
                        );
                    }


                    // Cùng kì năm trước
                    listDataCompareMonthConvert.Add(
                        new ReportDetailtForTotalMoneyType()
                        {
                            PartnerName = item,
                            Month = month.ToString(),
                            Year = (year - 1).ToString(),
                            VND = sumVNDLastYear,
                            USD = sumUSDLastYear,
                            EUR = sumEURLastYear,
                            CAD = sumCADLastYear,
                            AUD = sumAUDLastYear,
                            GBP = sumGBPLastYear,
                        }
                    );
                }
            }
            else
            {
                listDataCompareMonthConvert = new List<ReportDetailtForTotalMoneyType>(listDataCompareMonth);
            }

            // Số record của mảng
            int countArray = 1;
            GradationCompare[] arrayGradation = new GradationCompare[countArray * listDataCompareMonthConvert.Count];
            int count = 0;
            foreach (ReportDetailtForTotalMoneyType item in listDataCompareMonthConvert)
            {
                // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                arrayGradation[count] = new GradationCompare()
                {
                    NameGradationCompare = string.Format("EUR Tháng {0}/{1}", item.Month, item.Year),
                    amount = item.EUR,
                    NameType = item.PartnerName
                };

                count++;
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
        public ActionResult SearchColumnsChartCompareMonthForOneCAD([DataSourceRequest]DataSourceRequest request, int month, int year, string reportTypeID, string marketID)
        {
            List<ReportDetailtForTotalMoneyType> listDataCompareMonth = new ReportBL().ReportDetailtMTCompareMonthForOneConvert(year, month, reportTypeID, marketID);
            
            List<ReportDetailtForTotalMoneyType> listDataCompareMonthConvert = new List<ReportDetailtForTotalMoneyType>();
            // Trường hợp là thị trường chi tiết của loại tiền theo tháng của Châu Á
            if (marketID.Contains("005"))
            {
                // Danh sách các thị trường con của Châu Á
                List<string> listMarket = new List<string>();
                foreach (ReportDetailtForTotalMoneyType item in listDataCompareMonth)
                {
                    if (!listMarket.Contains(item.MarketName))
                    {
                        listMarket.Add(item.MarketName);
                    }
                }

                foreach (string item in listMarket)
                {
                    List<ReportDetailtForTotalMoneyType> dataItemLastYear = listDataCompareMonth.Where(x => x.MarketName == item && x.Month == month.ToString() && x.Year == (year - 1).ToString()).ToList();
                    List<ReportDetailtForTotalMoneyType> dataItemYear = listDataCompareMonth.Where(x => x.MarketName == item && x.Month == month.ToString() && x.Year == year.ToString()).ToList();
                    List<ReportDetailtForTotalMoneyType> dataItemLastMonth = listDataCompareMonth.Where(x => x.MarketName == item && x.Month == (month - 1).ToString() && x.Year == year.ToString()).ToList();
                    // Trường hợp tháng 1
                    if (month == 1)
                    {
                        dataItemLastMonth = listDataCompareMonth.Where(x => x.MarketName == item && x.Month == "12" && x.Year == (year - 1).ToString()).ToList();
                    }
                    // Cùng kì
                    if (dataItemLastYear.Count == 0)
                    {
                        dataItemLastYear = new List<ReportDetailtForTotalMoneyType>()
                        {
                            new ReportDetailtForTotalMoneyType()
                            {
                                MarketName = item,
                                Month = month.ToString(),
                                Year = (year - 1).ToString()
                            }
                        };
                    }

                    // Tháng hiện tại
                    if (dataItemYear.Count == 0)
                    {
                        dataItemYear = new List<ReportDetailtForTotalMoneyType>()
                        {
                            new ReportDetailtForTotalMoneyType()
                            {
                                MarketName = item,
                                Month = month.ToString(),
                                Year = year.ToString()
                            }
                        };
                    }

                    // Tháng trước
                    if (dataItemLastMonth.Count == 0)
                    {
                        // Trường hợp tháng 1
                        if (month == 1)
                        {
                            dataItemLastMonth = new List<ReportDetailtForTotalMoneyType>()
                            {
                                new ReportDetailtForTotalMoneyType()
                                {
                                    MarketName = item,
                                    Month = "12",
                                    Year = (year -1).ToString()
                                }
                            };
                        }
                        else
                        {
                            dataItemLastMonth = new List<ReportDetailtForTotalMoneyType>()
                            {
                                new ReportDetailtForTotalMoneyType()
                                {
                                    MarketName = item,
                                    Month = (month -1).ToString(),
                                    Year = year.ToString()
                                }
                            };
                        }
                    }

                    // Cùng kì năm trước
                    double sumVNDLastYear = dataItemLastYear.Sum(x => x.VND);
                    double sumUSDLastYear = dataItemLastYear.Sum(x => x.USD);
                    double sumEURLastYear = dataItemLastYear.Sum(x => x.EUR);
                    double sumCADLastYear = dataItemLastYear.Sum(x => x.CAD);
                    double sumAUDLastYear = dataItemLastYear.Sum(x => x.AUD);
                    double sumGBPLastYear = dataItemLastYear.Sum(x => x.GBP);

                    double sumTongDSLastYear = dataItemLastYear.Sum(x => x.TongDS);

                    // Tháng hiện tại
                    double sumVNDYear = dataItemYear.Sum(x => x.VND);
                    double sumUSDYear = dataItemYear.Sum(x => x.USD);
                    double sumEURYear = dataItemYear.Sum(x => x.EUR);
                    double sumCADYear = dataItemYear.Sum(x => x.CAD);
                    double sumAUDYear = dataItemYear.Sum(x => x.AUD);
                    double sumGBPYear = dataItemYear.Sum(x => x.GBP);

                    double sumTongDSYear = dataItemYear.Sum(x => x.TongDS);

                    // Tháng trước
                    double sumVNDLastMonth = dataItemLastMonth.Sum(x => x.VND);
                    double sumUSDLastMonth = dataItemLastMonth.Sum(x => x.USD);
                    double sumEURLastMonth = dataItemLastMonth.Sum(x => x.EUR);
                    double sumCADLastMonth = dataItemLastMonth.Sum(x => x.CAD);
                    double sumAUDLastMonth = dataItemLastMonth.Sum(x => x.AUD);
                    double sumGBPLastMonth = dataItemLastMonth.Sum(x => x.GBP);

                    double sumTongDSLastMonth = dataItemLastMonth.Sum(x => x.TongDS);

                    // Vì  get  dữ liệu theo partnerName nên thị trường gán cho partner
                    // Tháng hiện tại
                    listDataCompareMonthConvert.Add(
                        new ReportDetailtForTotalMoneyType()
                        {
                            PartnerName = item,
                            Month = month.ToString(),
                            Year = year.ToString(),
                            VND = sumVNDYear,
                            USD = sumUSDYear,
                            EUR = sumEURYear,
                            CAD = sumCADYear,
                            AUD = sumAUDYear,
                            GBP = sumGBPYear,
                        }
                    );

                    // Tháng Trước// Trường hợp tháng 1
                    if (month == 1)
                    {
                        listDataCompareMonthConvert.Add(
                            new ReportDetailtForTotalMoneyType()
                            {
                                PartnerName = item,
                                Month = "12",
                                Year = (year - 1).ToString(),
                                VND = sumVNDLastMonth,
                                USD = sumUSDLastMonth,
                                EUR = sumEURLastMonth,
                                CAD = sumCADLastMonth,
                                AUD = sumAUDLastMonth,
                                GBP = sumGBPLastMonth,
                            }
                        );
                    }
                    else
                    {
                        listDataCompareMonthConvert.Add(
                            new ReportDetailtForTotalMoneyType()
                            {
                                PartnerName = item,
                                Month = (month - 1).ToString(),
                                Year = year.ToString(),
                                VND = sumVNDLastMonth,
                                USD = sumUSDLastMonth,
                                EUR = sumEURLastMonth,
                                CAD = sumCADLastMonth,
                                AUD = sumAUDLastMonth,
                                GBP = sumGBPLastMonth,
                            }
                        );
                    }


                    // Cùng kì năm trước
                    listDataCompareMonthConvert.Add(
                        new ReportDetailtForTotalMoneyType()
                        {
                            PartnerName = item,
                            Month = month.ToString(),
                            Year = (year - 1).ToString(),
                            VND = sumVNDLastYear,
                            USD = sumUSDLastYear,
                            EUR = sumEURLastYear,
                            CAD = sumCADLastYear,
                            AUD = sumAUDLastYear,
                            GBP = sumGBPLastYear,
                        }
                    );
                }
            }
            else
            {
                listDataCompareMonthConvert = new List<ReportDetailtForTotalMoneyType>(listDataCompareMonth);
            }

            // Số record của mảng
            int countArray = 1;
            GradationCompare[] arrayGradation = new GradationCompare[countArray * listDataCompareMonthConvert.Count];
            int count = 0;
            foreach (ReportDetailtForTotalMoneyType item in listDataCompareMonthConvert)
            {
                // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                arrayGradation[count] = new GradationCompare()
                {
                    NameGradationCompare = string.Format("CAD Tháng {0}/{1}", item.Month, item.Year),
                    amount = item.CAD,
                    NameType = item.PartnerName
                };

                count++;
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
        public ActionResult SearchColumnsChartCompareMonthForOneAUD([DataSourceRequest]DataSourceRequest request, int month, int year, string reportTypeID, string marketID)
        {
            List<ReportDetailtForTotalMoneyType> listDataCompareMonth = new ReportBL().ReportDetailtMTCompareMonthForOneConvert(year, month, reportTypeID, marketID);
            
            List<ReportDetailtForTotalMoneyType> listDataCompareMonthConvert = new List<ReportDetailtForTotalMoneyType>();
            // Trường hợp là thị trường chi tiết của loại tiền theo tháng của Châu Á
            if (marketID.Contains("005"))
            {
                // Danh sách các thị trường con của Châu Á
                List<string> listMarket = new List<string>();
                foreach (ReportDetailtForTotalMoneyType item in listDataCompareMonth)
                {
                    if (!listMarket.Contains(item.MarketName))
                    {
                        listMarket.Add(item.MarketName);
                    }
                }

                foreach (string item in listMarket)
                {
                    List<ReportDetailtForTotalMoneyType> dataItemLastYear = listDataCompareMonth.Where(x => x.MarketName == item && x.Month == month.ToString() && x.Year == (year - 1).ToString()).ToList();
                    List<ReportDetailtForTotalMoneyType> dataItemYear = listDataCompareMonth.Where(x => x.MarketName == item && x.Month == month.ToString() && x.Year == year.ToString()).ToList();
                    List<ReportDetailtForTotalMoneyType> dataItemLastMonth = listDataCompareMonth.Where(x => x.MarketName == item && x.Month == (month - 1).ToString() && x.Year == year.ToString()).ToList();
                    // Trường hợp tháng 1
                    if (month == 1)
                    {
                        dataItemLastMonth = listDataCompareMonth.Where(x => x.MarketName == item && x.Month == "12" && x.Year == (year - 1).ToString()).ToList();
                    }
                    // Cùng kì
                    if (dataItemLastYear.Count == 0)
                    {
                        dataItemLastYear = new List<ReportDetailtForTotalMoneyType>()
                        {
                            new ReportDetailtForTotalMoneyType()
                            {
                                MarketName = item,
                                Month = month.ToString(),
                                Year = (year - 1).ToString()
                            }
                        };
                    }

                    // Tháng hiện tại
                    if (dataItemYear.Count == 0)
                    {
                        dataItemYear = new List<ReportDetailtForTotalMoneyType>()
                        {
                            new ReportDetailtForTotalMoneyType()
                            {
                                MarketName = item,
                                Month = month.ToString(),
                                Year = year.ToString()
                            }
                        };
                    }

                    // Tháng trước
                    if (dataItemLastMonth.Count == 0)
                    {
                        // Trường hợp tháng 1
                        if (month == 1)
                        {
                            dataItemLastMonth = new List<ReportDetailtForTotalMoneyType>()
                            {
                                new ReportDetailtForTotalMoneyType()
                                {
                                    MarketName = item,
                                    Month = "12",
                                    Year = (year -1).ToString()
                                }
                            };
                        }
                        else
                        {
                            dataItemLastMonth = new List<ReportDetailtForTotalMoneyType>()
                            {
                                new ReportDetailtForTotalMoneyType()
                                {
                                    MarketName = item,
                                    Month = (month -1).ToString(),
                                    Year = year.ToString()
                                }
                            };
                        }
                    }

                    // Cùng kì năm trước
                    double sumVNDLastYear = dataItemLastYear.Sum(x => x.VND);
                    double sumUSDLastYear = dataItemLastYear.Sum(x => x.USD);
                    double sumEURLastYear = dataItemLastYear.Sum(x => x.EUR);
                    double sumCADLastYear = dataItemLastYear.Sum(x => x.CAD);
                    double sumAUDLastYear = dataItemLastYear.Sum(x => x.AUD);
                    double sumGBPLastYear = dataItemLastYear.Sum(x => x.GBP);

                    double sumTongDSLastYear = dataItemLastYear.Sum(x => x.TongDS);

                    // Tháng hiện tại
                    double sumVNDYear = dataItemYear.Sum(x => x.VND);
                    double sumUSDYear = dataItemYear.Sum(x => x.USD);
                    double sumEURYear = dataItemYear.Sum(x => x.EUR);
                    double sumCADYear = dataItemYear.Sum(x => x.CAD);
                    double sumAUDYear = dataItemYear.Sum(x => x.AUD);
                    double sumGBPYear = dataItemYear.Sum(x => x.GBP);

                    double sumTongDSYear = dataItemYear.Sum(x => x.TongDS);

                    // Tháng trước
                    double sumVNDLastMonth = dataItemLastMonth.Sum(x => x.VND);
                    double sumUSDLastMonth = dataItemLastMonth.Sum(x => x.USD);
                    double sumEURLastMonth = dataItemLastMonth.Sum(x => x.EUR);
                    double sumCADLastMonth = dataItemLastMonth.Sum(x => x.CAD);
                    double sumAUDLastMonth = dataItemLastMonth.Sum(x => x.AUD);
                    double sumGBPLastMonth = dataItemLastMonth.Sum(x => x.GBP);

                    double sumTongDSLastMonth = dataItemLastMonth.Sum(x => x.TongDS);

                    // Vì  get  dữ liệu theo partnerName nên thị trường gán cho partner
                    // Tháng hiện tại
                    listDataCompareMonthConvert.Add(
                        new ReportDetailtForTotalMoneyType()
                        {
                            PartnerName = item,
                            Month = month.ToString(),
                            Year = year.ToString(),
                            VND = sumVNDYear,
                            USD = sumUSDYear,
                            EUR = sumEURYear,
                            CAD = sumCADYear,
                            AUD = sumAUDYear,
                            GBP = sumGBPYear,
                        }
                    );

                    // Tháng Trước// Trường hợp tháng 1
                    if (month == 1)
                    {
                        listDataCompareMonthConvert.Add(
                            new ReportDetailtForTotalMoneyType()
                            {
                                PartnerName = item,
                                Month = "12",
                                Year = (year - 1).ToString(),
                                VND = sumVNDLastMonth,
                                USD = sumUSDLastMonth,
                                EUR = sumEURLastMonth,
                                CAD = sumCADLastMonth,
                                AUD = sumAUDLastMonth,
                                GBP = sumGBPLastMonth,
                            }
                        );
                    }
                    else
                    {
                        listDataCompareMonthConvert.Add(
                            new ReportDetailtForTotalMoneyType()
                            {
                                PartnerName = item,
                                Month = (month - 1).ToString(),
                                Year = year.ToString(),
                                VND = sumVNDLastMonth,
                                USD = sumUSDLastMonth,
                                EUR = sumEURLastMonth,
                                CAD = sumCADLastMonth,
                                AUD = sumAUDLastMonth,
                                GBP = sumGBPLastMonth,
                            }
                        );
                    }


                    // Cùng kì năm trước
                    listDataCompareMonthConvert.Add(
                        new ReportDetailtForTotalMoneyType()
                        {
                            PartnerName = item,
                            Month = month.ToString(),
                            Year = (year - 1).ToString(),
                            VND = sumVNDLastYear,
                            USD = sumUSDLastYear,
                            EUR = sumEURLastYear,
                            CAD = sumCADLastYear,
                            AUD = sumAUDLastYear,
                            GBP = sumGBPLastYear,
                        }
                    );
                }
            }
            else
            {
                listDataCompareMonthConvert = new List<ReportDetailtForTotalMoneyType>(listDataCompareMonth);
            }

            // Số record của mảng
            int countArray = 1;
            GradationCompare[] arrayGradation = new GradationCompare[countArray * listDataCompareMonthConvert.Count];
            int count = 0;
            foreach (ReportDetailtForTotalMoneyType item in listDataCompareMonthConvert)
            {
                // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                arrayGradation[count] = new GradationCompare()
                {
                    NameGradationCompare = string.Format("AUD Tháng {0}/{1}", item.Month, item.Year),
                    amount = item.AUD,
                    NameType = item.PartnerName
                };

                count++;
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
        public ActionResult SearchColumnsChartCompareMonthForOneGBP([DataSourceRequest]DataSourceRequest request, int month, int year, string reportTypeID, string marketID)
        {
            List<ReportDetailtForTotalMoneyType> listDataCompareMonth = new ReportBL().ReportDetailtMTCompareMonthForOneConvert(year, month, reportTypeID, marketID);

            List<ReportDetailtForTotalMoneyType> listDataCompareMonthConvert = new List<ReportDetailtForTotalMoneyType>();
            // Trường hợp là thị trường chi tiết của loại tiền theo tháng của Châu Á
            if (marketID.Contains("005"))
            {
                // Danh sách các thị trường con của Châu Á
                List<string> listMarket = new List<string>();
                foreach (ReportDetailtForTotalMoneyType item in listDataCompareMonth)
                {
                    if (!listMarket.Contains(item.MarketName))
                    {
                        listMarket.Add(item.MarketName);
                    }
                }

                foreach (string item in listMarket)
                {
                    List<ReportDetailtForTotalMoneyType> dataItemLastYear = listDataCompareMonth.Where(x => x.MarketName == item && x.Month == month.ToString() && x.Year == (year - 1).ToString()).ToList();
                    List<ReportDetailtForTotalMoneyType> dataItemYear = listDataCompareMonth.Where(x => x.MarketName == item && x.Month == month.ToString() && x.Year == year.ToString()).ToList();
                    List<ReportDetailtForTotalMoneyType> dataItemLastMonth = listDataCompareMonth.Where(x => x.MarketName == item && x.Month == (month - 1).ToString() && x.Year == year.ToString()).ToList();
                    // Trường hợp tháng 1
                    if (month == 1)
                    {
                        dataItemLastMonth = listDataCompareMonth.Where(x => x.MarketName == item && x.Month == "12" && x.Year == (year - 1).ToString()).ToList();
                    }
                    // Cùng kì
                    if (dataItemLastYear.Count == 0)
                    {
                        dataItemLastYear = new List<ReportDetailtForTotalMoneyType>()
                        {
                            new ReportDetailtForTotalMoneyType()
                            {
                                MarketName = item,
                                Month = month.ToString(),
                                Year = (year - 1).ToString()
                            }
                        };
                    }

                    // Tháng hiện tại
                    if (dataItemYear.Count == 0)
                    {
                        dataItemYear = new List<ReportDetailtForTotalMoneyType>()
                        {
                            new ReportDetailtForTotalMoneyType()
                            {
                                MarketName = item,
                                Month = month.ToString(),
                                Year = year.ToString()
                            }
                        };
                    }

                    // Tháng trước
                    if (dataItemLastMonth.Count == 0)
                    {
                        // Trường hợp tháng 1
                        if (month == 1)
                        {
                            dataItemLastMonth = new List<ReportDetailtForTotalMoneyType>()
                            {
                                new ReportDetailtForTotalMoneyType()
                                {
                                    MarketName = item,
                                    Month = "12",
                                    Year = (year -1).ToString()
                                }
                            };
                        }
                        else
                        {
                            dataItemLastMonth = new List<ReportDetailtForTotalMoneyType>()
                            {
                                new ReportDetailtForTotalMoneyType()
                                {
                                    MarketName = item,
                                    Month = (month -1).ToString(),
                                    Year = year.ToString()
                                }
                            };
                        }
                    }

                    // Cùng kì năm trước
                    double sumVNDLastYear = dataItemLastYear.Sum(x => x.VND);
                    double sumUSDLastYear = dataItemLastYear.Sum(x => x.USD);
                    double sumEURLastYear = dataItemLastYear.Sum(x => x.EUR);
                    double sumCADLastYear = dataItemLastYear.Sum(x => x.CAD);
                    double sumAUDLastYear = dataItemLastYear.Sum(x => x.AUD);
                    double sumGBPLastYear = dataItemLastYear.Sum(x => x.GBP);

                    double sumTongDSLastYear = dataItemLastYear.Sum(x => x.TongDS);

                    // Tháng hiện tại
                    double sumVNDYear = dataItemYear.Sum(x => x.VND);
                    double sumUSDYear = dataItemYear.Sum(x => x.USD);
                    double sumEURYear = dataItemYear.Sum(x => x.EUR);
                    double sumCADYear = dataItemYear.Sum(x => x.CAD);
                    double sumAUDYear = dataItemYear.Sum(x => x.AUD);
                    double sumGBPYear = dataItemYear.Sum(x => x.GBP);

                    double sumTongDSYear = dataItemYear.Sum(x => x.TongDS);

                    // Tháng trước
                    double sumVNDLastMonth = dataItemLastMonth.Sum(x => x.VND);
                    double sumUSDLastMonth = dataItemLastMonth.Sum(x => x.USD);
                    double sumEURLastMonth = dataItemLastMonth.Sum(x => x.EUR);
                    double sumCADLastMonth = dataItemLastMonth.Sum(x => x.CAD);
                    double sumAUDLastMonth = dataItemLastMonth.Sum(x => x.AUD);
                    double sumGBPLastMonth = dataItemLastMonth.Sum(x => x.GBP);

                    double sumTongDSLastMonth = dataItemLastMonth.Sum(x => x.TongDS);

                    // Vì  get  dữ liệu theo partnerName nên thị trường gán cho partner
                    // Tháng hiện tại
                    listDataCompareMonthConvert.Add(
                        new ReportDetailtForTotalMoneyType()
                        {
                            PartnerName = item,
                            Month = month.ToString(),
                            Year = year.ToString(),
                            VND = sumVNDYear,
                            USD = sumUSDYear,
                            EUR = sumEURYear,
                            CAD = sumCADYear,
                            AUD = sumAUDYear,
                            GBP = sumGBPYear,
                        }
                    );

                    // Tháng Trước// Trường hợp tháng 1
                    if (month == 1)
                    {
                        listDataCompareMonthConvert.Add(
                            new ReportDetailtForTotalMoneyType()
                            {
                                PartnerName = item,
                                Month = "12",
                                Year = (year - 1).ToString(),
                                VND = sumVNDLastMonth,
                                USD = sumUSDLastMonth,
                                EUR = sumEURLastMonth,
                                CAD = sumCADLastMonth,
                                AUD = sumAUDLastMonth,
                                GBP = sumGBPLastMonth,
                            }
                        );
                    }
                    else
                    {
                        listDataCompareMonthConvert.Add(
                            new ReportDetailtForTotalMoneyType()
                            {
                                PartnerName = item,
                                Month = (month - 1).ToString(),
                                Year = year.ToString(),
                                VND = sumVNDLastMonth,
                                USD = sumUSDLastMonth,
                                EUR = sumEURLastMonth,
                                CAD = sumCADLastMonth,
                                AUD = sumAUDLastMonth,
                                GBP = sumGBPLastMonth,
                            }
                        );
                    }


                    // Cùng kì năm trước
                    listDataCompareMonthConvert.Add(
                        new ReportDetailtForTotalMoneyType()
                        {
                            PartnerName = item,
                            Month = month.ToString(),
                            Year = (year - 1).ToString(),
                            VND = sumVNDLastYear,
                            USD = sumUSDLastYear,
                            EUR = sumEURLastYear,
                            CAD = sumCADLastYear,
                            AUD = sumAUDLastYear,
                            GBP = sumGBPLastYear,
                        }
                    );
                }
            }
            else
            {
                listDataCompareMonthConvert = new List<ReportDetailtForTotalMoneyType>(listDataCompareMonth);
            }

            // Số record của mảng
            int countArray = 1;
            GradationCompare[] arrayGradation = new GradationCompare[countArray * listDataCompareMonthConvert.Count];
            int count = 0;
            foreach (ReportDetailtForTotalMoneyType item in listDataCompareMonthConvert)
            {
                // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                arrayGradation[count] = new GradationCompare()
                {
                    NameGradationCompare = string.Format("GBP Tháng {0}/{1}", item.Month, item.Year),
                    amount = item.GBP,
                    NameType = item.PartnerName
                };

                count++;
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
    }
}
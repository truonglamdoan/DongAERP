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
        public ActionResult MarketForTotal()
        {
            string nameUrl = "Báo cáo chi tiết/Theo doanh số chi trả/Theo thị trường/Loại tiền chi trả/Tất cả";
            ViewBag.NameURL = nameUrl;
            return View();
        }

        /// <summary>
        /// Báo cáo thị trường theo loại tiền chi trả từng thị trường
        /// </summary>
        /// <returns></returns>
        public ActionResult MarketForOne()
        {
            string nameUrl = "Báo cáo chi tiết/Theo doanh số chi trả/Theo thị trường/Loại tiền chi trả/Từng thị trường";
            ViewBag.NameURL = nameUrl;
            return View();
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
            List<ReportDetailtForTotalMoneyType> listData = new ReportBL().ListReportDetailtMarketForAll(reportTypeID);
            int count = 1;
            foreach (ReportDetailtForTotalMoneyType item in listData)
            {
                item.STT = (count++).ToString();
                item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;
            }

            ReportDetailtForTotalMoneyType dataItem = new ReportDetailtForTotalMoneyType()
            {
                MarketName = "Tổng",
                VND = listData.Sum(x => x.VND),
                USD = listData.Sum(x => x.USD),
                EUR = listData.Sum(x => x.EUR),
                CAD = listData.Sum(x => x.CAD),
                AUD = listData.Sum(x => x.AUD),
                GBP = listData.Sum(x => x.GBP),
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
            List<ReportDetailtForTotalMoneyType> listData = new ReportBL().SearchReportDetailtMarketForAll(fromDay, toDay, reportTypeID, marketID);

            foreach (ReportDetailtForTotalMoneyType item in listData)
            {
                item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;
            }

            ReportDetailtForTotalMoneyType dataItem = new ReportDetailtForTotalMoneyType()
            {
                MarketName = "Tổng",
                VND = listData.Sum(x => x.VND),
                USD = listData.Sum(x => x.USD),
                EUR = listData.Sum(x => x.EUR),
                CAD = listData.Sum(x => x.CAD),
                AUD = listData.Sum(x => x.AUD),
                GBP = listData.Sum(x => x.GBP),
                TongDS = listData.Sum(x => x.TongDS)
            };

            listData.Add(dataItem);

            return Json(listData.ToDataSourceResult(request), JsonRequestBehavior.AllowGet);
        }


        /// <summary>
        /// Bảng hiển thị thông tin các giao dịch qua các ngày
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult MarketForTotalConvert([DataSourceRequest]DataSourceRequest request, string reportTypeID)
        {
            List<ReportDetailtForTotalMoneyType> listData = new ReportBL().ListReportDetailtMarketForAllConvert(reportTypeID);
            int count = 1;
            foreach (ReportDetailtForTotalMoneyType item in listData)
            {
                item.STT = (count++).ToString();
                item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;
            }

            ReportDetailtForTotalMoneyType dataItem = new ReportDetailtForTotalMoneyType()
            {
                MarketName = "Tổng",
                VND = listData.Sum(x => x.VND),
                USD = listData.Sum(x => x.USD),
                EUR = listData.Sum(x => x.EUR),
                CAD = listData.Sum(x => x.CAD),
                AUD = listData.Sum(x => x.AUD),
                GBP = listData.Sum(x => x.GBP),
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
        public ActionResult SearchMarketForTotalConvert([DataSourceRequest]DataSourceRequest request, DateTime fromDay, DateTime toDay, string reportTypeID, string marketID)
        {
            List<ReportDetailtForTotalMoneyType> listData = new ReportBL().SearchReportDetailtMarketForAllConvert(fromDay, toDay, reportTypeID, marketID);

            foreach (ReportDetailtForTotalMoneyType item in listData)
            {
                item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;
            }

            ReportDetailtForTotalMoneyType dataItem = new ReportDetailtForTotalMoneyType()
            {
                MarketName = "Tổng",
                VND = listData.Sum(x => x.VND),
                USD = listData.Sum(x => x.USD),
                EUR = listData.Sum(x => x.EUR),
                CAD = listData.Sum(x => x.CAD),
                AUD = listData.Sum(x => x.AUD),
                GBP = listData.Sum(x => x.GBP),
                TongDS = listData.Sum(x => x.TongDS)
            };

            listData.Add(dataItem);

            return Json(listData.ToDataSourceResult(request), JsonRequestBehavior.AllowGet);
        }
    }
}
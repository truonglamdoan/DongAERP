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
    }
}
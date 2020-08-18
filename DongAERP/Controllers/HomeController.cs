using DongA.Bussiness;
using DongAERP.Models;
using Kendo.Mvc.Extensions;
using Kendo.Mvc.UI;
using Models;
using Models.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace DongAERP.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            //return View();
            return RedirectToAction("Index", "Login", new { Area = "Admin" });
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }

        public ActionResult Select([DataSourceRequest]DataSourceRequest request)
        {
            List<MarKetDetail> listData = new MarKetBL().GetAll();

            return Json(listData.ToDataSourceResult(request), JsonRequestBehavior.AllowGet);
        }
    }
}
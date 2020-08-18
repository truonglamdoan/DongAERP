﻿using DongA.Bussiness;
using DongA.Entities;
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

namespace DongAERP.Areas.Admin.Controllers
{
    [Authorize]
    public class HomeController : Controller
    {
        // GET: Admin/Home
        public ActionResult Index()
        {
            return View();
        }

        /// <summary>
        /// Get thông tin đối tác
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public PartialViewResult LeftMenu()
        {
            //List<Partner> listPartners = new PartnerBL().GetListPartner();
            List<Partner> listPartners = new List<Partner>() {
                new Partner {PartnerID = "PN00", PartnerName = "Tổng hợp", CreateDate = DateTime.Now, CreateUserID = "Admin" },
                new Partner {PartnerID = "PN01", PartnerName = "DAB", CreateDate = DateTime.Now, CreateUserID = "Admin" },
                new Partner {PartnerID = "PN02", PartnerName = "DAM", CreateDate = DateTime.Now, CreateUserID = "Admin" }
            };
            ViewBag.listPartners = listPartners;
            return PartialView("~/Areas/Admin/Views/Shared/_LeftMenu.cshtml");
        }

        /// <summary>
        /// Theo Doanh số tổng hợp
        /// </summary>
        /// <returns></returns>
        public JsonResult Categories()
        {
            List<Partner> list = new List<Partner>();
            list.Add(new Partner()
            {
                PartnerID = "item1",
                PartnerName = "Tổng doanh số chi trả"
            });

            list.Add(new Partner()
            {
                PartnerID = "item2",
                PartnerName = "Theo loại hình dịch vụ"
            });

            list.Add(new Partner()
            {
                PartnerID = "item3",
                PartnerName = "Theo thị trường"
            });

            list.Add(new Partner()
            {
                PartnerID = "item4",
                PartnerName = "Theo loại tiền chi trả"
            });
            return Json(list, JsonRequestBehavior.AllowGet);
        }

        public JsonResult CategoryDetailt()
        {
            List<Partner> list = new List<Partner>();
            list.Add(new Partner()
            {
                PartnerID = "item1",
                PartnerName = "Chi tiết"
            });

            list.Add(new Partner()
            {
                PartnerID = "item2",
                PartnerName = "So Sánh"
            });

            return Json(list, JsonRequestBehavior.AllowGet);
        }

        public JsonResult DetailtReport(string reportID)
        {
            List<Partner> list = new List<Partner>();
            if (reportID.Equals("item1"))
            {
                list.Add(new Partner()
                {
                    PartnerID = "item1",
                    PartnerName = "Theo ngày"
                });

                list.Add(new Partner()
                {
                    PartnerID = "item2",
                    PartnerName = "Theo tháng"
                });

                list.Add(new Partner()
                {
                    PartnerID = "item3",
                    PartnerName = "Theo năm"
                });
            }
            else
            {
                list.Add(new Partner()
                {
                    PartnerID = "item4",
                    PartnerName = "Theo giai đoạn"
                });

                list.Add(new Partner()
                {
                    PartnerID = "item5",
                    PartnerName = "Theo tháng"
                });
            }

            return Json(list, JsonRequestBehavior.AllowGet);
        }
        public JsonResult ReportType()
        {
            List<Partner> list = new List<Partner>();
            list.Add(new Partner()
            {
                PartnerID = "2",
                PartnerName = "Tổng hợp"
            });

            list.Add(new Partner()
            {
                PartnerID = "0",
                PartnerName = "DAB"
            });

            list.Add(new Partner()
            {
                PartnerID = "1",
                PartnerName = "DAMTC"
            });

            return Json(list, JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// Theo Doanh số chi tiết
        /// </summary>
        /// <returns></returns>
        public JsonResult CategoriesDetailt()
        {
            List<Partner> list = new List<Partner>();
            list.Add(new Partner()
            {
                PartnerID = "item1",
                PartnerName = "Thị trường"
            });

            list.Add(new Partner()
            {
                PartnerID = "item2",
                PartnerName = "Đối tác"
            });
            return Json(list, JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// Loại hình cho doanh số chi tiết
        /// </summary>
        /// <returns></returns>
        public JsonResult CategoryMenuDetailt()
        {
            List<Partner> list = new List<Partner>();
            list.Add(new Partner()
            {
                PartnerID = "item1",
                PartnerName = "Chi trả - Tất cả"
            });

            list.Add(new Partner()
            {
                PartnerID = "item2",
                PartnerName = "Chi trả - Từng thị trường"
            });

            list.Add(new Partner()
            {
                PartnerID = "item3",
                PartnerName = "Chi trả - So Sánh"
            });

            return Json(list, JsonRequestBehavior.AllowGet);
        }

        public JsonResult CompareDetailtReport(string reportID)
        {
            List<Partner> list = new List<Partner>();
            if (reportID.Equals("item3"))
            {
                list.Add(new Partner()
                {
                    PartnerID = "item1",
                    PartnerName = "Theo giai đoạn - Tất cả"
                });

                list.Add(new Partner()
                {
                    PartnerID = "item2",
                    PartnerName = "Theo giai đoạn - Từng thị trường"
                });

                list.Add(new Partner()
                {
                    PartnerID = "item3",
                    PartnerName = "Theo tháng - Tất cả"
                });

                list.Add(new Partner()
                {
                    PartnerID = "item4",
                    PartnerName = "Theo tháng - Từng thị trường"
                });
            }
            return Json(list, JsonRequestBehavior.AllowGet);
        }

        public ActionResult CreateDataMarket()
        {
            List<ReportForMaket> result = new List<ReportForMaket>();
            result = new ReportBL().CreateDataMarket();
            return Json(result, JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// Theo Doanh số tổng hợp
        /// </summary>
        /// <returns></returns>
        public JsonResult TypeOfSale()
        {
            List<Partner> list = new List<Partner>();
            list.Add(new Partner()
            {
                PartnerID = "item1",
                PartnerName = "Doanh số tổng hợp"
            });

            list.Add(new Partner()
            {
                PartnerID = "item2",
                PartnerName = "Doanh số chi tiết"
            });

            return Json(list, JsonRequestBehavior.AllowGet);
        }
    }
}
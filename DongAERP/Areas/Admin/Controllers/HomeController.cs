using DongA.Bussiness;
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

        public ActionResult ReportDetailt()
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
        public JsonResult CategoryLevelOne(string levelZeroID)
        {
            List<Partner> list = new List<Partner>();
            if (levelZeroID.Equals("level0_item1"))
            {
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
            }
            else
            {
                list.Add(new Partner()
                {
                    PartnerID = "item1",
                    PartnerName = "Thị trường - Loại hình"
                });

                list.Add(new Partner()
                {
                    PartnerID = "item2",
                    PartnerName = "Thị trường - Loại tiền"
                });

                list.Add(new Partner()
                {
                    PartnerID = "item3",
                    PartnerName = "Đối tác - Loại hình"
                });

                list.Add(new Partner()
                {
                    PartnerID = "item3",
                    PartnerName = "Đối tác - Loại tiền"
                });
            }
            
            return Json(list, JsonRequestBehavior.AllowGet);
        }

        public JsonResult CategoryLevelTwo(string levelOneID)
        {
            List<Partner> list = new List<Partner>();
            if (levelOneID.Equals("level0_item1"))
            {
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
            }
            else
            {
                list.Add(new Partner()
                {
                    PartnerID = "item1",
                    PartnerName = "Tất cả"
                });

                list.Add(new Partner()
                {
                    PartnerID = "item2",
                    PartnerName = "Từng"
                });

                list.Add(new Partner()
                {
                    PartnerID = "item3",
                    PartnerName = "So sánh"
                });
            }
            
            return Json(list, JsonRequestBehavior.AllowGet);
        }

        public JsonResult CategoryLevelThree(string levelZeroID, string levelTwoID)
        {
            List<Partner> list = new List<Partner>();
            // Tổng hợp
            if (levelZeroID.Equals("level0_item1"))
            {
                if(levelTwoID == "item1")
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
                        PartnerID = "item1",
                        PartnerName ="Giai đoạn"
                    });

                    list.Add(new Partner()
                    {
                        PartnerID = "item2",
                        PartnerName = "Theo tháng"
                    });
                }
            }

            // Chi tiết
            if (levelZeroID.Equals("level0_item2"))
            {
                if (levelTwoID == "item1" || levelTwoID == "item2")
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

                if (levelTwoID == "item3")
                {
                    list.Add(new Partner()
                    {
                        PartnerID = "item1",
                        PartnerName = "Giai đoạn - Tất cả"
                    });

                    list.Add(new Partner()
                    {
                        PartnerID = "item2",
                        PartnerName = "Giai đoạn - Từng"
                    });

                    list.Add(new Partner()
                    {
                        PartnerID = "item3",
                        PartnerName = "Tháng - Tất cả"
                    });

                    list.Add(new Partner()
                    {
                        PartnerID = "item4",
                        PartnerName = "Tháng - Từng"
                    });
                }
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
                PartnerName = "Thị trường - Loại hình"
            });

            list.Add(new Partner()
            {
                PartnerID = "item2",
                PartnerName = "Thị trường - Loại tiền"
            });

            list.Add(new Partner()
            {
                PartnerID = "item3",
                PartnerName = "Đối tác - Loại hình"
            });

            list.Add(new Partner()
            {
                PartnerID = "item3",
                PartnerName = "Đối tác - Loại tiền"
            });
            return Json(list, JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// Loại hình cho doanh số chi tiết
        /// </summary>
        /// <returns></returns>
        public JsonResult CategoryMenuDetailt(string categoriesDetailt)
        {
            List<Partner> list = new List<Partner>();
            string text = string.Empty;
            switch(categoriesDetailt)
            {
                case "item1":
                    text = "Loại hình";
                    break;

                case "item2":
                    text = "Loại tiền";
                    break;
                default:

                    break;
            }

            list.Add(new Partner()
            {
                PartnerID = "item1",
                PartnerName = string.Format("{0} - Tất cả", text)
            });

            list.Add(new Partner()
            {
                PartnerID = "item2",
                PartnerName = string.Format("{0} - Từng thị trường", text)
            });

            list.Add(new Partner()
            {
                PartnerID = "item3",
                PartnerName = string.Format("{0} - So Sánh", text)
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
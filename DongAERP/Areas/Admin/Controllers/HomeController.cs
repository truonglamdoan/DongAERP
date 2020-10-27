using DongA.Bussiness;
using DongA.Entities;
using DongAERP.Areas.Admin.Models;
using DongAERP.Models;
using Kendo.Mvc.Extensions;
using Kendo.Mvc.UI;
using Models;
using Models.Framework;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.Configuration;
using System.Web.Mvc;

namespace DongAERP.Areas.Admin.Controllers
{
    [Authorize]
    public class HomeController : Controller
    {
        // GET: Admin/Home
        public ActionResult Index()
        {
            CountryModel objProductModel = new CountryModel();
            objProductModel.CountrytData = new Country();
            objProductModel.CountrytData = GetChartData();
            objProductModel.CountryTitle = "Country";
            objProductModel.PopulationTitle = "Population";

            ViewBag.GetDataMartket = GetChartDataWorld();

            return View(objProductModel);
            //return View();
        }

        [HttpPost]
        public ActionResult Index(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            // Get dữ liệu chỉ tiêu của hệ thống thiết lập
            string targetPoint = WebConfigurationManager.AppSettings["targetPoint"];
            string[] dataTarget = targetPoint.Split('-');
            string year = fromDate.Year.ToString();
            string targetValue = string.Empty;

            foreach(string item in dataTarget)
            {
                string[] dataArray = item.Split('_');

                if(dataArray[0] == year)
                {
                    targetValue = dataArray[1];
                }
            }

            CountryModel objProductModel = new CountryModel();
            objProductModel.CountrytData = new Country();
            // Get biểu đồ của Việt Nam
            objProductModel.CountrytData = GetChartData();
            objProductModel.CountryTitle = "Country";
            objProductModel.PopulationTitle = "Population";

            if (fromDate != null && toDate != null && reportTypeID != null)
            {
                List<string> listData = new List<string>()
                {
                    fromDate.ToString("MM/dd/yyyy"),
                    toDate.ToString("MM/dd/yyyy"),
                    reportTypeID
                };

                ViewData["listData"] = listData;
            }

            // Get dữ liệu của Biểu đồ thế giới
            ViewBag.GetDataMartket = GetChartDataWorld();

            //// Get dữ liệu của doanh số
            //List<Report> listDataDoanhSo = GetDataDoanhSo(fromDate, toDate, reportTypeID);

            //foreach(Report item in listDataDoanhSo)
            //{
            //    item.TongDS = item.DSChiQuay + item.DSChiNha + item.DSCK;
            //}

            //if(listDataDoanhSo.Count > 0)
            //{
            //    ViewBag.TongDS = listDataDoanhSo.Sum(x => x.TongDS);
            //    ViewBag.DSChiNha = listDataDoanhSo.Sum(x => x.DSChiNha);
            //}

            return View(objProductModel);
            //return View();
        }

        public Country GetChartData()
        {
            string listRegion = "An Giang, Bà Rịa–Vũng Tàu, Bắc Giang, Bắc Kạn, Bạc Liêu, Bắc Ninh, Bến Tre, Bình Định, Bình Dương, Bình Phước, Bình Thuận, Cà Mau, Cao Bằng, " +
                "Đắk Lắk, Đắk Nông, Điện Biên, Đồng Nai, Đồng Tháp, Gia Lai, Hà Giang, Hà Nam, Hà Tĩnh, Hải Dương, Hậu Giang, Hòa Bình, Hưng Yên, Khánh Hòa, Kiến Giang, Kon Tum, " +
                "Lai Châu, Lâm Đồng, Lạng Sơn, Lào Cai, Long An, Nam Định, Nghệ An, Ninh Bình, Ninh Thuận, Phú Thọ, Phú Yên, Quảng Bình, Quảng Nam, Quảng Ngãi, Quảng Ninh, Quảng Trị, " +
                "Sóc Trăng, Sơn La, Tây Ninh, Thái Bình, Thái Nguyên, Thanh Hóa, Thừa Thiên–Huế, Tiền Giang, Trà Vinh, Tuyên Quang, Vĩnh Long, Vĩnh Phúc, Yên Bái, Cần Thơ, Đà Nẵng, Hà Nội, Hai Phong, Hồ Chí Minh City";
            Country objproduct = new Country();
            /*Get the data from databse and prepare the chart record data in string form.*/
            objproduct.CountryName = listRegion;
            int count = 0;
            int i = 1;
            string[] str = objproduct.CountryName.Split(',');
            string[] listArr = new string[str.Count()];
            foreach(string item in str)
            {
                listArr[count] = (i++).ToString();
                count++;
            }
            objproduct.Population = string.Join(", ", listArr);
            return objproduct;
        }

        public Country GetChartDataWorld()
        {
            string listRegion = "Germany, United States, Brazil, Canada, France, RU";
            Country objproduct = new Country();
            /*Get the data from databse and prepare the chart record data in string form.*/
            objproduct.CountryName = listRegion;

            objproduct.Population = "200, 300, 400, 500, 600, 700";
            return objproduct;
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
        public JsonResult CategoryLevelZero(string levelID)
        {
            List<Partner> list = new List<Partner>();
            if (levelID.Equals("1"))
            {
                list.Add(new Partner()
                {
                    PartnerID = "level0_item1",
                    PartnerName = "Tổng hợp"
                });

                list.Add(new Partner()
                {
                    PartnerID = "level0_item2",
                    PartnerName = "Chi tiết"
                });
            }
            else
            {
                list.Add(new Partner()
                {
                    PartnerID = "level0_item3",
                    PartnerName = "Tổng hợp"
                });
                list.Add(new Partner()
                {
                    PartnerID = "level0_item4",
                    PartnerName = "Chi tiết"
                });
            }

            return Json(list, JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// Theo Doanh số tổng hợp
        /// </summary>
        /// <returns></returns>
        public JsonResult CategoryLevelOne(string levelZeroID)
        {
            List<Partner> list = new List<Partner>();
            switch (levelZeroID)
            {
                case "level0_item1":
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

                    break;
                case "level0_item2":

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
                        PartnerID = "item4",
                        PartnerName = "Đối tác - Loại tiền"
                    });
                    break;

                case "level0_item3":

                    list.Add(new Partner()
                    {
                        PartnerID = "item1",
                        PartnerName = "HS-TH Loại hình"
                    });

                    list.Add(new Partner()
                    {
                        PartnerID = "item2",
                        PartnerName = "HS-TH Loại tiền"
                    });

                    list.Add(new Partner()
                    {
                        PartnerID = "item3",
                        PartnerName = "HS-TH Thị trường"
                    });

                    list.Add(new Partner()
                    {
                        PartnerID = "item4",
                        PartnerName = "HS-TH Tổng hồ sơ"
                    });
                    break;
                default:

                    list.Add(new Partner()
                    {
                        PartnerID = "item1",
                        PartnerName = "HS- Thị trường - Loại hình"
                    });

                    list.Add(new Partner()
                    {
                        PartnerID = "item2",
                        PartnerName = "HS- Thị trường - Loại tiền"
                    });

                    list.Add(new Partner()
                    {
                        PartnerID = "item3",
                        PartnerName = "HS- Đối tác - Loại hình"
                    });

                    list.Add(new Partner()
                    {
                        PartnerID = "item4",
                        PartnerName = "HS- Đối tác - Loại tiền"
                    });
                    break;
            }
            
            return Json(list, JsonRequestBehavior.AllowGet);
        }

        public JsonResult CategoryLevelTwo(string levelOneID)
        {
            List<Partner> list = new List<Partner>();
            // 05/10/2020
            // Thêm vào Hồ sơ
            if (levelOneID.Equals("level0_item1") || levelOneID.Equals("level0_item3"))
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


            if (levelOneID.Equals("level0_item2") || levelOneID.Equals("level0_item4"))
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
            // 05/10/2020 Thêm vào Hồ sơ với levelZeroID.Equals("level0_item1")
            if (levelZeroID.Equals("level0_item1") || levelZeroID.Equals("level0_item3"))
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
            if (levelZeroID.Equals("level0_item2") || levelZeroID.Equals("level0_item4"))
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

        /// <summary>
        /// Danh sách các thị trường theo đối tác
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Trường Lãm]   Created [10/08/2020]
        /// </history>
        public JsonResult ListMarketForPartner()
        {
            List<Partner> list = new ReportBL().ListPartner();
            return Json(list, JsonRequestBehavior.AllowGet);
        }

        #region Get dữ liệu cho Dashoard


        public JsonResult GetDataDoanhSo(DateTime? fromDate, DateTime? toDate, string reportTypeID)
        {
            JsonResult result = new JsonResult();
            
            List<Report> listData = new ReportBL().SearchMonth(fromDate.Value, toDate.Value, reportTypeID);
            foreach(Report item in listData)
            {
                item.TongDS = item.DSChiQuay + item.DSChiNha + item.DSCK;
            }

            double TongDS = 0;
            double DSChiNha = 0;
            Dictionary<string, double> listStr = new Dictionary<string, double>();

            if(listData.Count > 0)
            {
                TongDS = listData.Sum(x => x.TongDS);
                DSChiNha = listData.Sum(x => x.DSChiNha);
                listStr.Add("TongDS", TongDS);
                listStr.Add("DSChiNha", DSChiNha);
            }

            result = this.Json(JsonConvert.SerializeObject(listStr), JsonRequestBehavior.AllowGet);  

            return result;
        }

        #endregion
    }
}
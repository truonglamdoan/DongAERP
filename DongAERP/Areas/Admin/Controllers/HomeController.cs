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
using System.Data;
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
            //CountryModel objProductModel = new CountryModel();
            //objProductModel.CountrytData = new Country();
            //objProductModel.CountrytData = GetChartData();
            //objProductModel.CountryTitle = "Country";
            //objProductModel.PopulationTitle = "Population";

            //ViewBag.GetDataMartket = GetChartDataWorld();

            return View();
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
            objProductModel.CountrytData = GetChartData(fromDate, toDate, reportTypeID);
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
            return View(objProductModel);
        }

        public Country GetChartData(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            List<City> listData = new ReportBL().SearchCity(fromDate, toDate, reportTypeID);

            Country objproduct = new Country();

            string listRegion = string.Empty;
            List<string> listcountryCode = new List<string>();
            List<string> listcountryName = new List<string>();

            List<double> listRECNUM = new List<double>();
            List<double> listAMOUNT = new List<double>();

            
            if (listData.Count > 0)
            {
                string countryCode = string.Empty;
                // Sửa lại các tên bị sai so với google Map
                foreach (City item in listData)
                {
                    switch(item.CityName)
                    {
                        case "HUẾ":
                            countryCode = "VN-26";
                            break;
                        case "HỒ CHÍ MINH":
                            countryCode = "VN-SG";
                            break;
                        case "ĐỒNG NAI":
                            countryCode = "VN-39";
                            break;
                        case "VŨNG TÀU":
                            countryCode = "VN-43";
                            break;
                        case "TIỀN GIANG":
                            countryCode = "VN-46";
                            break;
                        case "KHÁNH HÒA":
                            countryCode = "VN-34";
                            break;
                        case "CẦN THƠ":
                            countryCode = "VN-CT";
                            break;
                        case "ĐÀ NẲNG" :
                            countryCode = "VN-DN";
                            break;
                        case "KIÊN GIANG":
                            countryCode = "VN-47";
                            break;
                        case "NGHỆ AN":
                            countryCode = "VN-22";
                            break;
                        case "SÓC TRĂNG":
                            countryCode = "VN-52";
                            break;
                        case "AN GIANG":
                            countryCode = "VN-44";
                            break;
                        case "VĨNH LONG":
                            countryCode = "VN-49";
                            break;
                        case "HÀ NỘI":
                            countryCode = "VN-HN";
                            break;
                        case "TRÀ VINH":
                            countryCode = "VN-51";
                            break;
                        case "HẢI PHÒNG":
                            countryCode = "VN-HP";
                            break;
                        case "BẾN TRE":
                            countryCode = "VN-50";
                            break;
                        case "HƯNG YÊN":
                            countryCode = "VN-66";
                            break;
                        case "BẠC LIÊU":
                            countryCode = "VN-55";
                            break;
                        case "BÌNH THUẬN":
                            countryCode = "VN-40";
                            break;
                        case "BÌNH DƯƠNG":
                            countryCode = "VN-57";
                            break;
                        case "TÂY NINH":
                            countryCode = "VN-37";
                            break;
                        case "ĐỒNG THÁP":
                            countryCode = "VN-45";
                            break;
                        case "LÂM ĐỒNG":
                            countryCode = "VN-35";
                            break;
                        case "BẮC GIANG":
                            countryCode = "VN-54";
                            break;
                        case "QUẢNG BÌNH":
                            countryCode = "VN-24";
                            break;
                        case "BÌNH ĐỊNH":
                            countryCode = "VN-31";
                            break;
                        case "CÀ MAU":
                            countryCode = "VN-59";
                            break;
                        case "PHÚ YÊN":
                            countryCode = "VN-32";
                            break;
                        case "ĐẮKLẮK":
                            countryCode = "VN-33";
                            break;
                        case "THANH HOÁ":
                            countryCode = "VN-21";
                            break;
                        case "NINH THUẬN":
                            countryCode = "VN-36";
                            break;
                        case "THÁI BÌNH":
                            countryCode = "VN-20";
                            break;
                        case "NAM ĐỊNH":
                            countryCode = "VN-67";
                            break;
                        case "QUẢNG NGÃI":
                            countryCode = "VN-29";
                            break;
                        case "QUẢNG NINH":
                            countryCode = "VN-13";
                            break;
                        case "VĨNH PHÚC":
                            countryCode = "VN-70";
                            break;
                        case "QUẢNG NAM":
                            countryCode = "VN-27";
                            break;
                        case "BÌNH PHƯỚC":
                            countryCode = "VN-58";
                            break;
                        case "GIA LAI":
                            countryCode = "VN-30";
                            break;
                        case "THÁI NGUYÊN":
                            countryCode = "VN-69";
                            break;
                        case "BẮC NINH":
                            countryCode = "VN-56";
                            break;
                        case "HÀ NAM":
                            countryCode = "VN-63";
                            break;
                        case "LONG AN":
                            countryCode = "VN-41";
                            break;
                        case "KON TUM":
                            countryCode = "VN-28";
                            break;
                        case "ĐẮK NÔNG":
                            countryCode = "VN-72";
                            break;

                        default:
                            countryCode = "";
                            break;
                    }
                    listcountryCode.Add(countryCode);
                    listcountryName.Add(item.CityName);
                    listRECNUM.Add(item.RECNUM);
                    listAMOUNT.Add(item.AMOUNT);
                }
            }

            if (listcountryCode.Count > 0)
            {
                objproduct.CountryCode = string.Join(",", listcountryCode);
            }

            /*Get the data from databse and prepare the chart record data in string form.*/
            objproduct.CountryName = string.Join(",", listcountryName);

            if (listData.Count > 0)
            {
                objproduct.Population = string.Join(",", listRECNUM);
            }

            if (listData.Count > 0)
            {
                objproduct.Area = string.Join(",", listAMOUNT);
            }

            return objproduct;
        }

        public Country GetChartDataWorld(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            // Doanh số
            List<ReportForMaket> listDataDS = new ReportBL().SearchReportMaketForMonth(fromDate, toDate, reportTypeID);

            // Hồ sơ
            List<ReportForMaket> listDataHS = new HSReportBL().SearchReportMaketForMonth(fromDate, toDate, reportTypeID);

            string[] listTypeMoney = { "American", "Asia", "Global", "Europe", "Canada", "Australia" };
            string[] listTypeMoneyVN = { "United States", "Asia", "Toàn Cầu", "Europe", "Canada", "Australia" };

            DataTable table = new DataTable();
            // Khởi tạo datatable
            table = new DataTable();
            // Tạo các cột cho datatable
            table.Columns.Add("PartnerCode", typeof(String));
            table.Columns.Add("PartnerName", typeof(String));

            table.Columns.Add("COL1", typeof(double));
            table.Columns.Add("COL2", typeof(double));

            
            List<string> listCountryCode = new List<string>();
            List<string> listCountryName = new List<string>();
            List<string> listPopulation = new List<string>();
            List<string> listArea = new List<string>();

            if (listDataDS.Count > 0 && listDataHS.Count > 0)
            {
                // Doanh so
                ReportForMaket dataIemSumDS = new ReportForMaket()
                {
                    American = listDataDS.Sum(x => x.American),
                    Asia = listDataDS.Sum(x => x.Asia),
                    Global = listDataDS.Sum(x => x.Global),
                    Europe = listDataDS.Sum(x => x.Europe),
                    Canada = listDataDS.Sum(x => x.Canada),
                    Australia = listDataDS.Sum(x => x.Australia),
                };

                // Ho so
                ReportForMaket dataIemSumHS = new ReportForMaket()
                {
                    American = listDataHS.Sum(x => x.American),
                    Asia = listDataHS.Sum(x => x.Asia),
                    Global = listDataHS.Sum(x => x.Global),
                    Europe = listDataHS.Sum(x => x.Europe),
                    Canada = listDataHS.Sum(x => x.Canada),
                    Australia = listDataHS.Sum(x => x.Australia),
                };


                int count = 0;
                string marketCode = string.Empty;

                foreach (string item in listTypeMoney)
                {
                    // Trường hợp thuộc châu á
                    if (item.Contains("Asia"))
                    {
                        string marketID = "005";
                        List<ReportDetailtForTotalMoneyType> listDataAsiaDS = new ReportBL().SearchReportDetailtMTForOneForMonthConvert(fromDate, toDate, reportTypeID, marketID);
                        List<string> listMartket = new List<string>();
                        foreach(ReportDetailtForTotalMoneyType dataItem in listDataAsiaDS)
                        {
                            dataItem.TongDS = dataItem.VND + dataItem.USD + dataItem.EUR + dataItem.CAD + dataItem.AUD + dataItem.GBP;

                            if (!listMartket.Contains(dataItem.MarketName))
                            {
                                listMartket.Add(dataItem.MarketName);
                            }
                        }

                        List<ReportDetailtForTotalMoneyType> listDataAsiaHS = new HSReportBL().SearchReportDetailtMTForOneForMonth(fromDate, toDate, reportTypeID, marketID);

                        foreach (ReportDetailtForTotalMoneyType dataItem in listDataAsiaHS)
                        {
                            dataItem.TongDS = dataItem.VND + dataItem.USD + dataItem.EUR + dataItem.CAD + dataItem.AUD + dataItem.GBP;
                        }

                        if (listDataAsiaDS.Count > 0 && listDataAsiaHS.Count > 0)
                        {
                            

                            string[] strArray = { "VND", "USD", "EUR", "CAD", "AUD", "GBP" };

                            foreach(string item1 in listMartket)
                            {
                                marketCode = item1;
                                switch (item1)
                                {
                                    case "Hàn Quốc":
                                        marketCode = "KP";
                                        break;
                                    case "Nhật Bản":
                                        marketCode = "JP";
                                        break;
                                    case "Đài Loan":
                                        marketCode = "Taiwan";
                                        break;
                                    case "Malaysia":
                                        marketCode = "MY";
                                        break;
                                    case "Singapore":
                                        marketCode = "SG";
                                        break;
                                    default:
                                        marketCode = "";
                                        break;
                                }

                                table.Rows.Add(
                                    marketCode, item1, listDataAsiaHS.Where(x=>x.MarketName == item1).Sum(x=>x.TongDS), listDataAsiaDS.Where(x => x.MarketName == item1).Sum(x => x.TongDS)
                                );
                            }
                        }
                    }
                    else
                    {
                        string martketName = string.Empty; 
                        switch (item)
                        {
                            case "American":
                                marketCode = "US";
                                martketName = "United States (Mỹ)";
                                break;
                            case "Global":
                                marketCode = "Global";
                                martketName = "Global";
                                break;
                            case "Europe":
                                marketCode = "FR";
                                martketName = "France (Châu Âu)";
                                break;
                            case "Canada":
                                marketCode = "CA";
                                martketName = "Canada";
                                break;
                            default:
                                marketCode = "AU";
                                martketName = "Australia (Úc)";
                                break;
                        }
                        //  Doanh so
                        var propertyInfo = dataIemSumDS.GetType().GetProperty(item);
                        var valueData = propertyInfo.GetValue(dataIemSumDS, null);

                        // Ho so
                        var propertyInfoConvert = dataIemSumHS.GetType().GetProperty(item);
                        var valueDataConvert = propertyInfoConvert.GetValue(dataIemSumHS, null);

                        table.Rows.Add(
                            marketCode, martketName, Convert.ToDouble(valueDataConvert), Convert.ToDouble(valueData)
                        );
                    }
                    
                }

                if(table.Rows.Count > 0)
                {
                    foreach (DataRow dtRow in table.Rows)
                    {
                        // On all tables' columns
                        foreach (DataColumn dc in table.Columns)
                        {
                            if (dc.ColumnName == "PartnerCode")
                            {
                                listCountryCode.Add(dtRow[dc].ToString());
                            }

                            if (dc.ColumnName == "PartnerName")
                            {
                                listCountryName.Add(dtRow[dc].ToString());
                            }

                            if (dc.ColumnName == "COL1")
                            {
                                listPopulation.Add(dtRow[dc].ToString());
                            }

                            if (dc.ColumnName == "COL2")
                            {
                                listArea.Add(dtRow[dc].ToString());
                            }
                        }
                    }
                }
            }

            Country objproduct = new Country();

            if (listCountryCode.Count > 0)
            {
                objproduct.CountryCode = string.Join(",", listCountryCode);
            }

            if (listCountryName.Count > 0)
            {
                objproduct.CountryName = string.Join(",", listCountryName);
            }

            if (listPopulation.Count > 0)
            {
                objproduct.Population = string.Join(",", listPopulation);
            }

            if (listArea.Count > 0)
            {
                objproduct.Area = string.Join(",", listArea);
            }

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
            // Doanh số
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

            // Hồ sơ
            if(levelID.Equals("2"))
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

            // TCKT
            if (levelID.Equals("3"))
            {
                list.Add(new Partner()
                {
                    PartnerID = "level0_item5",
                    PartnerName = "Tổng hợp"
                });
                list.Add(new Partner()
                {
                    PartnerID = "level0_item6",
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
                // Tổng hợp DS
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
                // Chi tiết DS
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
               // Tổng hợp HS
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

                // Chi tiết HS
                case "level0_item4":

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
                // Tổng hợp TCKT
                case "level0_item5":

                    list.Add(new Partner()
                    {
                        PartnerID = "item1",
                        PartnerName = "Doanh số theo thị trường"
                    });
                    break;

                // Chi tiết TCKT
                default:
                    list.Add(new Partner()
                    {
                        PartnerID = "item1",
                        PartnerName = "Chi tiết DS theo thị trường"
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

            // DS tài chính kế toán
            if (levelOneID.Equals("level0_item5") || levelOneID.Equals("level0_item6"))
            {
                list.Add(new Partner()
                {
                    PartnerID = "item1",
                    PartnerName = "Chi tiết"
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

            // DS TCKT
            if (levelZeroID.Equals("level0_item5") || levelZeroID.Equals("level0_item6"))
            {
                if (levelTwoID == "item1")
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
        
        public JsonResult GetDataDoanhSo(DateTime fromDate, DateTime toDate, string reportTypeID)
        {

            Dictionary<string, double> listStr = new Dictionary<string, double>();

            // Lấy giá trị mục tiêu của năm
            string targetPoint = WebConfigurationManager.AppSettings["targetPoint"];
            string[] dataTarget = targetPoint.Split('-');
            string year = fromDate.Year.ToString();
            string targetValue = string.Empty;

            foreach (string item in dataTarget)
            {
                string[] dataArray = item.Split('_');

                if (dataArray[0] == year)
                {
                    targetValue = dataArray[1];
                }
            }

            listStr.Add("targetPoint", Convert.ToDouble(targetValue));

            // Lấy giá trị mục tiêu Doanh số chi nhà của năm
            string targetDSChiNhaPoint = WebConfigurationManager.AppSettings["targetDSChiNhaPoint"];
            dataTarget = targetDSChiNhaPoint.Split('-');
            string targetDSChiNhaValue = string.Empty;

            foreach (string item in dataTarget)
            {
                string[] dataArray = item.Split('_');

                if (dataArray[0] == year)
                {
                    targetDSChiNhaValue = dataArray[1];
                }
            }
            
            listStr.Add("targetDSChiNha", Convert.ToDouble(targetDSChiNhaValue));

            //JsonResult result = new JsonResult();
            
            List<Report> listData = new ReportBL().SearchMonth(fromDate, toDate, reportTypeID);
            foreach(Report item in listData)
            {
                item.TongDS = item.DSChiQuay + item.DSChiNha + item.DSCK;
            }

            double TongDS = 0;
            double DSChiNha = 0;
            
            if (listData.Count > 0)
            {
                TongDS = listData.Sum(x => x.TongDS);
                DSChiNha = listData.Sum(x => x.DSChiNha);
                listStr.Add("TongDS", TongDS);
                listStr.Add("DSChiNha", DSChiNha);
            }

            // 03/11/2020 - [Trường Lãm] - Begin Add
            // Lấy giá trị mục tiêu Doanh số chi nhà của năm
            string targetRose = WebConfigurationManager.AppSettings["targetRose"];
            dataTarget = targetRose.Split('-');
            string targetRoseValue = string.Empty;

            foreach (string item in dataTarget)
            {
                string[] dataArray = item.Split('_');

                if (dataArray[0] == year)
                {
                    targetRoseValue = dataArray[1];
                }
            }

            listStr.Add("targetRoseValue", Convert.ToDouble(targetRoseValue));
            

            //List<Report> listData = new ReportBL().SearchMonth(fromDate, toDate, reportTypeID);
            //foreach (Report item in listData)
            //{
            //    item.TongDS = item.DSChiQuay + item.DSChiNha + item.DSCK;
            //}
            
            // Giá trị Hoa hồng
            double roseValueItem = 2000000000;

            listStr.Add("roseValueItem", roseValueItem);

            //if (listData.Count > 0)
            //{
            //    TongDS = listData.Sum(x => x.TongDS);
            //    DSChiNha = listData.Sum(x => x.DSChiNha);
            //    listStr.Add("TongDS", TongDS);
            //    listStr.Add("roseValueItem", DSChiNha);
            //}


            // Lấy giá trị mục tiêu Doanh số chi nhà của năm
            string targetNgoaiTe = WebConfigurationManager.AppSettings["targetNgoaiTe"];
            dataTarget = targetRose.Split('-');
            string targetNgoaiTeValue = string.Empty;

            foreach (string item in dataTarget)
            {
                string[] dataArray = item.Split('_');

                if (dataArray[0] == year)
                {
                    targetNgoaiTeValue = dataArray[1];
                }
            }

            listStr.Add("targetNgoaiTeValue", Convert.ToDouble(targetNgoaiTeValue));


            //List<Report> listData = new ReportBL().SearchMonth(fromDate, toDate, reportTypeID);
            //foreach (Report item in listData)
            //{
            //    item.TongDS = item.DSChiQuay + item.DSChiNha + item.DSCK;
            //}

            // Giá trị Hoa hồng
            double ngoaiTeValueItem = 3500000000;

            listStr.Add("ngoaiTeValueItem", ngoaiTeValueItem);
            // 03/11/2020 - [Trường Lãm] - End Add

            // Get số hồ sơ
            List<Report> listDataHS = new HSReportBL().SearchMonth(fromDate, toDate, reportTypeID);
            foreach (Report item in listDataHS)
            {
                item.TongDS = item.DSChiQuay + item.DSChiNha + item.DSCK;
            }

            double TongHS = 0;
            double HSChiNha = 0;
            double HSChiQuay = 0;
            double HSChuyenKhoan = 0;

            if (listDataHS.Count > 0)
            {
                TongHS = listDataHS.Sum(x => x.TongDS);

                HSChiNha = listDataHS.Sum(x => x.DSChiNha);
                HSChiQuay = listDataHS.Sum(x => x.DSChiQuay);
                HSChuyenKhoan = listDataHS.Sum(x => x.DSCK);

                listStr.Add("TongHS", TongHS);
                listStr.Add("HSChiNha", HSChiNha);
                listStr.Add("HSChiQuay", HSChiQuay);
                listStr.Add("HSChuyenKhoan", HSChuyenKhoan);
            }

            // Get chart Data
            Country dataChartVN = GetChartData(fromDate, toDate, reportTypeID);

            Country dataChartWorld = GetChartDataWorld(fromDate, toDate, reportTypeID);
            
            var result = this.Json(new { data = JsonConvert.SerializeObject(listStr), dataVN = dataChartVN, dataWorld = dataChartWorld }, JsonRequestBehavior.AllowGet);  

            return result;
        }

        /// <summary>
        /// search data cho biểu đồ của tháng hiện tại
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult SearchPieMonthDS([DataSourceRequest]DataSourceRequest request, DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            List<Report> listData = new ReportBL().SearchMonth(fromDate, toDate, reportTypeID);
            Report listDataGradationMonth = new Report();
            foreach(Report item in listData)
            {
                item.TongDS = item.DSChiQuay + item.DSChiNha + item.DSCK;
            }

            if (listData.Count > 0)
            {
                double sumDSChiQuay = listData.Sum(x => x.DSChiQuay);
                double sumDSChiNha = listData.Sum(x => x.DSChiNha);
                double sumDSChuyenKhoan = listData.Sum(x => x.DSCK);
                double sumTongDS = listData.Sum(x => x.TongDS);

                listDataGradationMonth = new Report()
                {
                    DSChiQuay = sumTongDS == 0 ? 0 : Math.Round((sumDSChiQuay / sumTongDS) * 100, 2, MidpointRounding.ToEven),
                    DSChiNha = sumTongDS == 0 ? 0 : Math.Round((sumDSChiNha / sumTongDS) * 100, 2, MidpointRounding.ToEven),
                    DSCK = sumTongDS == 0 ? 0 : Math.Round((sumDSChuyenKhoan / sumTongDS) * 100, 2, MidpointRounding.ToEven)
                };
            }
            
            // # dòng record
            GradationChartPie[] arrayGradation = null;

            int count = 0;
            if (listDataGradationMonth != null)
            {
                arrayGradation = new GradationChartPie[3];
                // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                arrayGradation[count] = new GradationChartPie()
                {
                    category = "Chi quầy",
                    value = listDataGradationMonth.DSChiQuay,
                    //color = "#FFBF00"
                };

                count++;
                arrayGradation[count] = new GradationChartPie()
                {
                    category = "Chi nhà",
                    value = listDataGradationMonth.DSChiNha,
                    //color = "#40FF00"
                };

                count++;
                arrayGradation[count] = new GradationChartPie()
                {
                    category = "CK",
                    value = listDataGradationMonth.DSCK,
                    //color = "#2ECCFA"
                };
            }

            if (arrayGradation == null)
            {
                arrayGradation = new GradationChartPie[1];
                arrayGradation[0] = new GradationChartPie()
                {
                    category = "1",
                    value = 0,
                    color = "#9de219"

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
        public ActionResult GridDataMoneyType([DataSourceRequest]DataSourceRequest request, DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            List<ReportForTotalMoneyType> listData = new ReportBL().SearchReportTMTForMonth(fromDate, toDate, reportTypeID);
            List<ReportForTotalMoneyType> listDataConvert = new ReportBL().SearchReportTMTForMonthConvert(fromDate, toDate, reportTypeID);


            DataTable table = new DataTable();
            // Khởi tạo datatable
            table = new DataTable();
            // Tạo các cột cho datatable
            table.Columns.Add("PartnerName", typeof(String));

            table.Columns.Add("COL1", typeof(double));
            table.Columns.Add("COL2", typeof(double));

            if (listData.Count > 0 && listDataConvert.Count > 0)
            {
                // Nguyên tệ
                ReportForTotalMoneyType dataIemSum = new ReportForTotalMoneyType()
                {
                    VND = listData.Sum(x => x.VND),
                    USD = listData.Sum(x => x.USD),
                    EUR = listData.Sum(x => x.EUR),
                    CAD = listData.Sum(x => x.CAD),
                    AUD = listData.Sum(x => x.AUD),
                    GBP = listData.Sum(x => x.GBP),
                };

                // Quy USD
                ReportForTotalMoneyType dataIemSumConvert = new ReportForTotalMoneyType()
                {
                    VND = listDataConvert.Sum(x => x.VND),
                    USD = listDataConvert.Sum(x => x.USD),
                    EUR = listDataConvert.Sum(x => x.EUR),
                    CAD = listDataConvert.Sum(x => x.CAD),
                    AUD = listDataConvert.Sum(x => x.AUD),
                    GBP = listDataConvert.Sum(x => x.GBP),
                };


                string[] listTypeMoney = { "VND", "USD", "EUR", "CAD", "AUD", "GBP" };

                foreach (string item in listTypeMoney)
                {
                    // Nguyên tệ
                    var propertyInfo = dataIemSum.GetType().GetProperty(item);
                    var valueData = propertyInfo.GetValue(dataIemSum, null);

                    // Quy USD
                    var propertyInfoConvert = dataIemSumConvert.GetType().GetProperty(item);
                    var valueDataConvert = propertyInfoConvert.GetValue(dataIemSumConvert, null);

                    table.Rows.Add(
                        item, Convert.ToDouble(valueData), Convert.ToDouble(valueDataConvert)
                    );
                }
            }
            
            DataRow row = table.NewRow();
            row["PartnerName"] = "Tổng";
            row["COL1"] = 0;
            row["COL2"] = table.Compute("Sum(COL2)", "");
            table.Rows.Add(row);

            return Json(table.ToDataSourceResult(request), JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// search data cho biểu đồ của tháng hiện tại
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult SearchPieMonthHS([DataSourceRequest]DataSourceRequest request, DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            List<ReportForTotalMoneyType> listData = new ReportBL().SearchReportTMTForMonthConvert(fromDate, toDate, reportTypeID);
            ReportForTotalMoneyType listDataGradationMonth = new ReportForTotalMoneyType();
            foreach (ReportForTotalMoneyType item in listData)
            {
                item.TongDS = item.VND + item.USD + item.EUR +  item.CAD + item.AUD + item.GBP;
            }

            if (listData.Count > 0)
            {
                double sumVND = listData.Sum(x => x.VND);
                double sumUSD = listData.Sum(x => x.USD);
                double sumEUR = listData.Sum(x => x.EUR);
                double sumCAD = listData.Sum(x => x.CAD);
                double sumAUD = listData.Sum(x => x.AUD);
                double sumGBP = listData.Sum(x => x.GBP);

                double sumTongDS = listData.Sum(x => x.TongDS);

                listDataGradationMonth = new ReportForTotalMoneyType()
                {
                    VND = sumTongDS == 0 ? 0 : Math.Round((sumVND / sumTongDS) * 100, 2, MidpointRounding.ToEven),
                    USD = sumTongDS == 0 ? 0 : Math.Round((sumUSD / sumTongDS) * 100, 2, MidpointRounding.ToEven),
                    EUR = sumTongDS == 0 ? 0 : Math.Round((sumEUR / sumTongDS) * 100, 2, MidpointRounding.ToEven),
                    CAD = sumTongDS == 0 ? 0 : Math.Round((sumCAD / sumTongDS) * 100, 2, MidpointRounding.ToEven),
                    AUD = sumTongDS == 0 ? 0 : Math.Round((sumAUD / sumTongDS) * 100, 2, MidpointRounding.ToEven),
                    GBP = sumTongDS == 0 ? 0 : Math.Round((sumGBP / sumTongDS) * 100, 2, MidpointRounding.ToEven),
                };
            }

            // # dòng record
            GradationChartPie[] arrayGradation = null;

            int count = 0;
            if (listDataGradationMonth != null)
            {
                arrayGradation = new GradationChartPie[6];
                // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                arrayGradation[count] = new GradationChartPie()
                {
                    category = "VND",
                    value = listDataGradationMonth.VND,
                    //color = "#FFBF00"
                };

                count++;
                arrayGradation[count] = new GradationChartPie()
                {
                    category = "USD",
                    value = listDataGradationMonth.USD,
                    //color = "#40FF00"
                };

                count++;
                arrayGradation[count] = new GradationChartPie()
                {
                    category = "EUR",
                    value = listDataGradationMonth.EUR,
                    //color = "#2ECCFA"
                };
                count++;
                arrayGradation[count] = new GradationChartPie()
                {
                    category = "CAD",
                    value = listDataGradationMonth.CAD,
                    //color = "#2ECCFA"
                };
                count++;
                arrayGradation[count] = new GradationChartPie()
                {
                    category = "AUD",
                    value = listDataGradationMonth.AUD,
                    //color = "#2ECCFA"
                };
                count++;
                arrayGradation[count] = new GradationChartPie()
                {
                    category = "GBP",
                    value = listDataGradationMonth.GBP,
                    //color = "#2ECCFA"
                };
            }

            if (arrayGradation == null)
            {
                arrayGradation = new GradationChartPie[1];
                arrayGradation[0] = new GradationChartPie()
                {
                    category = "1",
                    value = 0,
                    color = "#9de219"

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
        public ActionResult GridDataMarket([DataSourceRequest]DataSourceRequest request, DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            // Doanh số
            List<ReportForMaket> listDataDS = new ReportBL().SearchReportMaketForMonth(fromDate, toDate, reportTypeID);

            // Hồ sơ
            List<ReportForMaket> listDataHS = new HSReportBL().SearchReportMaketForMonth(fromDate, toDate, reportTypeID);

            string[] listTypeMoney = { "American", "Asia", "Global", "Europe", "Canada", "Australia" };
            string[] listTypeMoneyVN = { "Mỹ", "Châu Á", "Toàn Cầu", "Châu Âu", "Canada", "Úc" };

            DataTable table = new DataTable();
            // Khởi tạo datatable
            table = new DataTable();
            // Tạo các cột cho datatable
            table.Columns.Add("PartnerName", typeof(String));

            table.Columns.Add("COL1", typeof(double));
            table.Columns.Add("COL2", typeof(double));

            if (listDataDS.Count > 0 && listDataHS.Count > 0)
            {
                // Doanh so
                ReportForMaket dataIemSumDS = new ReportForMaket()
                {
                    American = listDataDS.Sum(x => x.American),
                    Asia = listDataDS.Sum(x => x.Asia),
                    Global = listDataDS.Sum(x => x.Global),
                    Europe = listDataDS.Sum(x => x.Europe),
                    Canada = listDataDS.Sum(x => x.Canada),
                    Australia = listDataDS.Sum(x => x.Australia),
                };

                // Ho so
                ReportForMaket dataIemSumHS = new ReportForMaket()
                {
                    American = listDataHS.Sum(x => x.American),
                    Asia = listDataHS.Sum(x => x.Asia),
                    Global = listDataHS.Sum(x => x.Global),
                    Europe = listDataHS.Sum(x => x.Europe),
                    Canada = listDataHS.Sum(x => x.Canada),
                    Australia = listDataHS.Sum(x => x.Australia),
                };


                int count = 0;
                foreach (string item in listTypeMoney)
                {
                    //  Doanh so
                    var propertyInfo = dataIemSumDS.GetType().GetProperty(item);
                    var valueData = propertyInfo.GetValue(dataIemSumDS, null);

                    // Ho so
                    var propertyInfoConvert = dataIemSumHS.GetType().GetProperty(item);
                    var valueDataConvert = propertyInfoConvert.GetValue(dataIemSumHS, null);

                    table.Rows.Add(
                        listTypeMoneyVN[count++], Convert.ToDouble(valueDataConvert), Convert.ToDouble(valueData)
                    );
                }

                // Sắp xếp
                DataView dv = table.DefaultView;
                dv.Sort = "COL2 desc";
                table = dv.ToTable();
            }

            DataRow row = table.NewRow();
            row["PartnerName"] = "Tổng";
            row["COL1"] = table.Compute("Sum(COL1)", "");
            row["COL2"] = table.Compute("Sum(COL2)", "");
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
        public ActionResult GridDataPartner([DataSourceRequest]DataSourceRequest request, DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            // Doanh số
            List<ReportDetailtForPartner> listDataDS = new ReportBL().SearchPartnerForTotalForMonth(fromDate, toDate, reportTypeID);
            

            foreach (ReportDetailtForPartner item in listDataDS)
            {
                item.TongDS = item.DSChiQuay + item.DSChiNha + item.DSCK;
            }
            // Hồ sơ
            List<ReportDetailtForPartner> listDataHS = new HSReportBL().SearchReportDetailtPartnerForMonth(fromDate, toDate, reportTypeID);

            foreach (ReportDetailtForPartner item in listDataHS)
            {
                item.TongDS = item.DSChiQuay + item.DSChiNha + item.DSCK;
            }
            

            string[] listTypeMoney = { "American", "Asia", "Global", "Europe", "Canada", "Australia" };

            DataTable table = new DataTable();
            // Khởi tạo datatable
            table = new DataTable();
            // Tạo các cột cho datatable
            table.Columns.Add("PartnerName", typeof(String));

            table.Columns.Add("COL1", typeof(double));
            table.Columns.Add("COL2", typeof(double));

            if (listDataDS.Count > 0 && listDataHS.Count > 0)
            {
                foreach (ReportDetailtForPartner item in listDataDS)
                {
                    ReportDetailtForPartner dataHS = listDataHS.Find(x => x.PartnerName == item.PartnerName);
                    
                    if(dataHS == null)
                    {
                        dataHS = new ReportDetailtForPartner()
                        {
                            PartnerName = item.PartnerName
                        };
                    }

                    table.Rows.Add(item.PartnerName, dataHS.TongDS, item.TongDS);
                }

                // Sắp xếp
                DataView dv = table.DefaultView;
                dv.Sort = "COL2 desc";
                table = dv.ToTable();
            }

            DataRow row = table.NewRow();
            row["PartnerName"] = "Tổng";
            row["COL1"] = table.Compute("Sum(COL1)", "");
            row["COL2"] = table.Compute("Sum(COL2)", "");
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
        public ActionResult GridDataCity([DataSourceRequest]DataSourceRequest request, DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            // Doanh số
            List<City> listData = new ReportBL().SearchCity(fromDate, toDate, reportTypeID);
            // Sắp xếp
            if (listData.Count > 0)
            {
                listData = listData.OrderByDescending(x => x.AMOUNT).ToList();
            }
            DataTable table = new DataTable();
            // Khởi tạo datatable
            table = new DataTable();
            // Tạo các cột cho datatable
            table.Columns.Add("PartnerName", typeof(String));

            table.Columns.Add("COL1", typeof(double));
            table.Columns.Add("COL2", typeof(double));

            foreach(City item in listData)
            {
                table.Rows.Add(item.CityName, item.RECNUM, item.AMOUNT);
            }

            DataRow row = table.NewRow();
            row["PartnerName"] = "Tổng";
            row["COL1"] = table.Compute("Sum(COL1)", "");
            row["COL2"] = table.Compute("Sum(COL2)", "");
            table.Rows.Add(row);

            return Json(table.ToDataSourceResult(request), JsonRequestBehavior.AllowGet);
        }
        #endregion

        #region Tạo form nhập chỉ tiêu

        public ActionResult FormTarget()
        {

            return View();
        }

        public ActionResult TargetRead([DataSourceRequest] DataSourceRequest request)
        {
            List<FormTarget> listTarget = new TargetBL().GetListTarget();
            
            //return Json(productService.Read().ToDataSourceResult(request));
            return Json(listTarget.ToDataSourceResult(request), JsonRequestBehavior.AllowGet);
        }

        [HttpPost]
        [AcceptVerbs(HttpVerbs.Post)]
        public ActionResult EditingCustom_Create([DataSourceRequest] DataSourceRequest request,
            [Bind(Prefix = "models")]IEnumerable<FormTarget> formTarget)
        {
            var results = new List<FormTarget>();

            if (formTarget != null && ModelState.IsValid)
            {
                foreach (var item in formTarget)
                {
                    if (string.IsNullOrEmpty(item.ObjectID))
                    {
                        item.ObjectID = Guid.NewGuid().ToString();
                    }
                    bool check = new TargetBL().Insert(item);

                    if(check)
                    {
                        results.Add(item);
                    }
                }
            }

            return Json(results.ToDataSourceResult(request, ModelState));
        }

        [HttpPost]
        [AcceptVerbs(HttpVerbs.Post)]
        public ActionResult EditingCustom_Update([DataSourceRequest] DataSourceRequest request,
            [Bind(Prefix = "models")]IEnumerable<FormTarget> formTarget)
        {
            var results = new List<FormTarget>();

            if (formTarget != null && ModelState.IsValid)
            {
                foreach (var item in formTarget)
                {
                    if (string.IsNullOrEmpty(item.ObjectID))
                    {
                        item.ObjectID = Guid.NewGuid().ToString();
                    }
                    bool check = new TargetBL().UpdateDataTarget(item);

                    if (check)
                    {
                        results.Add(item);
                    }
                }
            }

            return Json(formTarget.ToDataSourceResult(request, ModelState));
        }

        [HttpPost]
        [AcceptVerbs(HttpVerbs.Post)]
        public ActionResult EditingCustom_Destroy([DataSourceRequest] DataSourceRequest request,
            [Bind(Prefix = "models")]IEnumerable<FormTarget> formTarget)
        {
            var results = new List<FormTarget>();

            if (formTarget != null && ModelState.IsValid)
            {
                foreach (var item in formTarget)
                {
                    if (string.IsNullOrEmpty(item.ObjectID))
                    {
                        item.ObjectID = Guid.NewGuid().ToString();
                    }
                    bool check = new TargetBL().DeleteDataTarget(item);

                    if (check)
                    {
                        results.Add(item);
                    }
                }
            }

            return Json(formTarget.ToDataSourceResult(request, ModelState));
        }
        #endregion
    }
}
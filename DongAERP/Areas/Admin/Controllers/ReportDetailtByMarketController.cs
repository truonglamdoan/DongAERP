﻿using DongA.Bussiness;
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
            string nameUrl = "Báo cáo chi tiết/Theo doanh số chi trả/Theo thị trường/Loại hình dịch vụ/Từng thị trường";
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
            string nameUrl = "Báo cáo chi tiết/Theo doanh số chi trả/Theo thị trường/Loại hình dịch vụ/Từng thị trường";
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

            table.Columns.Add("ReportID", typeof(String));

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
                    foreach (ReportDetailtSTMarket item in listData)
                    {
                        item.ReportID = item.PartnerName;
                        item.TongDS = item.DSChiNha + item.DSChiQuay + item.DSCK;
                    }
                }
            }
            
            //ReportDetailtSTMarket dataItem = new ReportDetailtSTMarket()
            //{
            //    MarketName = "Tổng",
            //    DSChiQuay = listData.Sum(x => x.DSChiQuay),
            //    DSChiNha = listData.Sum(x => x.DSChiNha),
            //    DSCK = listData.Sum(x => x.DSCK),
            //    TongDS = listData.Sum(x => x.TongDS)
            //};

            //listData.Add(dataItem);

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
                    foreach (ReportDetailtSTMarket item in listData)
                    {
                        item.ReportID = item.PartnerName;
                        item.TongDS = item.DSChiNha + item.DSChiQuay + item.DSCK;
                    }
                }
            }

            //ReportDetailtSTMarket dataItem = new ReportDetailtSTMarket()
            //{
            //    MarketName = "Tổng",
            //    DSChiQuay = listData.Sum(x => x.DSChiQuay),
            //    DSChiNha = listData.Sum(x => x.DSChiNha),
            //    DSCK = listData.Sum(x => x.DSCK),
            //    TongDS = listData.Sum(x => x.TongDS)
            //};

            //listData.Add(dataItem);

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
                    foreach (ReportDetailtSTMarket item in listData)
                    {
                        item.ReportID = item.PartnerName;
                        item.TongDS = item.DSChiNha + item.DSChiQuay + item.DSCK;
                    }
                }
            }

            //ReportDetailtSTMarket dataItem = new ReportDetailtSTMarket()
            //{
            //    MarketName = "Tổng",
            //    DSChiQuay = listData.Sum(x => x.DSChiQuay),
            //    DSChiNha = listData.Sum(x => x.DSChiNha),
            //    DSCK = listData.Sum(x => x.DSCK),
            //    TongDS = listData.Sum(x => x.TongDS)
            //};

            //listData.Add(dataItem);

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

            if (!marketID.Equals("005"))
            {
                foreach (ReportDetailtServiceType item in listDataGradation)
                {
                    item.TongDS = item.DSChiQuay + item.DSChiNha + item.DSCK;
                }

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

            DataTable table = new DataTable();
            // Khởi tạo datatable
            table = new DataTable();
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

            table.Columns.Add("ReportID", typeof(String));

            // Trường hợp với tất cả thị trường
            if (marketID.Equals("0"))
            {
                foreach (ReportDetailtSTMarket item in listDataCompareMonth)
                {
                    item.ReportID = item.MarketName;
                    item.TongDS = item.DSChiQuay + item.DSChiNha + item.DSCK;
                    item.MarketName = "Tất cả thị trường";
                }
                
                // Danh sách mã thị trường của Tất cả
                List<string> listMarket = new List<string>() { "003", "005", "001", "002", "014", "004" };
                
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
                            , dataItemLastYear.DSChiQuay, dataItemLastYear.DSChiNha, dataItemLastYear.DSCK, dataItemLastYear.TongDS
                            , dataItemYear.ReportID);
                    }
                }

            }
            else
            {
                foreach (ReportDetailtSTMarket item in listDataCompareMonth)
                {
                    item.ReportID = item.PartnerName;
                    item.TongDS = item.DSChiQuay + item.DSChiNha + item.DSCK;
                }

                foreach (ReportDetailtSTMarket item in listDataCompareMonth)
                {
                    // Cùng kì
                    ReportDetailtSTMarket dataItemLastYear = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == month.ToString() && x.Year == (year - 1).ToString());
                    ReportDetailtSTMarket dataItemYear = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == month.ToString() && x.Year == year.ToString());
                    ReportDetailtSTMarket dataItemLastMonth = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == (month - 1).ToString() && x.Year == year.ToString());

                    // Cùng kì năm trước
                    if (dataItemLastYear == null)
                    {
                        dataItemLastYear = new ReportDetailtSTMarket();
                    }

                    // Hiện tại
                    if (dataItemYear == null)
                    {
                        dataItemYear = new ReportDetailtSTMarket();
                    }

                    // tháng trước
                    if (dataItemYear == null)
                    {
                        dataItemLastMonth = new ReportDetailtSTMarket();
                        
                    }

                    // Check tồn tại của item
                    string value = string.Format("ReportID='{0}'", item.ReportID);

                    DataRow[] foundRows = table.Select(value);

                    if (dataItemLastYear != null && dataItemYear != null && dataItemLastMonth != null && foundRows.Count() == 0)
                    {
                        // add item vào table
                        table.Rows.Add(item.MarketName, dataItemYear.DSChiQuay, dataItemYear.DSChiNha, dataItemYear.DSCK, dataItemYear.TongDS
                            , dataItemLastMonth.DSChiQuay, dataItemLastMonth.DSChiNha, dataItemLastMonth.DSCK, dataItemLastMonth.TongDS
                            , dataItemLastYear.DSChiQuay, dataItemLastYear.DSChiNha, dataItemLastYear.DSCK, dataItemLastYear.TongDS
                            , item.ReportID);
                    }

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

            // trường hợp tất cả thị trường
            if (marketID.Equals("0"))
            {
                foreach (ReportDetailtSTMarket item in listDataCompareMonth)
                {
                    item.ReportID = item.MarketName;
                    item.TongDS = item.DSChiQuay + item.DSChiNha + item.DSCK;
                    item.MarketName = "Tất cả thị trường";
                }
            }
            else
            {
                // Trường hợp các thị trường con của thị trường châu Á
                foreach (ReportDetailtSTMarket item in listDataCompareMonth)
                {
                    item.ReportID = item.PartnerName;
                    item.TongDS = item.DSChiQuay + item.DSChiNha + item.DSCK;
                }
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

            table.Columns.Add("ReportID", typeof(String));

            foreach (ReportDetailtSTMarket item in listDataCompareMonth)
            {
                // Cùng kì
                ReportDetailtSTMarket dataItemLastYear = listDataCompareMonth.Find(x => x.MarketCode == item.MarketCode && x.Month == item.Month && x.Year == (year - 1).ToString());
                ReportDetailtSTMarket dataItemMonth = listDataCompareMonth.Find(x => x.MarketCode == item.MarketCode && x.Month == item.Month && x.Year == year.ToString());
                ReportDetailtSTMarket dataItemLastMonth = listDataCompareMonth.Find(x => x.MarketCode == item.MarketCode && x.Month == (Int32.Parse(item.Month) - 1).ToString() && x.Year == year.ToString());

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

            //DataRow row = table.NewRow();
            //row["MarketName"] = "Tổng";
            //// So sánh với tháng trước
            //row["CQ1"] = table.Compute("Sum(CQ1)", "");
            //row["CN1"] = table.Compute("Sum(CN1)", "");
            //row["CK1"] = table.Compute("Sum(CK1)", "");
            //row["TDS1"] = table.Compute("Sum(TDS1)", "");

            //// So sánh với cùng kì năm trươc
            //row["CQ2"] = table.Compute("Sum(CQ2)", "");
            //row["CN2"] = table.Compute("Sum(CN2)", "");
            //row["CK2"] = table.Compute("Sum(CK2)", "");
            //row["TDS2"] = table.Compute("Sum(TDS2)", "");
            //table.Rows.Add(row);

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
            List<ReportDetailtSTMarket> listDataCompareMonth = new ReportBL().ColumnChartStackCompareMonthForAllPercent(year, month, reportTypeID, marketID);

            GradationCharColumn[] arrayGradation = null;

            // Trường hợp tất cả thị trường
            if (marketID.Equals("0"))
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
            else
            {
                // List các thị trường của châu Á
                List<string> listAsianItem = new List<string>();
                foreach (ReportDetailtSTMarket item in listDataCompareMonth)
                {
                    if (!listAsianItem.Contains(item.MarketName))
                    {
                        listAsianItem.Add(item.MarketName);
                    }
                }

                // Số record của mảng
                int countArray = 9;
                arrayGradation = new GradationCharColumn[countArray * listAsianItem.Count];
                int count = 0;
                int tooltip = 1;

                List<ReportDetailtSTMarket> listDataYear = new List<ReportDetailtSTMarket>();
                List<ReportDetailtSTMarket> listDataLastMonth = new List<ReportDetailtSTMarket>();
                List<ReportDetailtSTMarket> listDataLastYear = new List<ReportDetailtSTMarket>();

                foreach (string item in listAsianItem)
                {
                    listDataYear = listDataCompareMonth.Where(x => x.MarketName == item &&  x.Month == month.ToString() && x.Year == year.ToString()).ToList();
                    listDataLastMonth = listDataCompareMonth.Where(x => x.MarketName == item && x.Month == (month - 1).ToString() && x.Year == year.ToString()).ToList();
                    listDataLastYear = listDataCompareMonth.Where(x => x.MarketName == item && x.Month == month.ToString() && x.Year == (year - 1).ToString()).ToList();
                    
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
                    arrayGradation[count] = new GradationCharColumn()
                    {
                        Serie = item,
                        Segmento = string.Format("Chi Quầy {0}/{1}", month, year),
                        Valor1 = sumDSChiQuayYear,
                        Tooltip = tooltip
                    };

                    count++;
                    arrayGradation[count] = new GradationCharColumn()
                    {
                        Serie = item,
                        Segmento = string.Format("Chi Nhà {0}/{1}", month, year),
                        Valor1 = sumDSChiNhaYear,
                        Tooltip = tooltip
                    };

                    count++;
                    arrayGradation[count] = new GradationCharColumn()
                    {
                        Serie = item,
                        Segmento = string.Format("DSCK {0}/{1}", month, year),
                        Valor1 = sumDSCKYear,
                        Tooltip = tooltip
                    };
                    count++;

                    // Tháng trước
                    arrayGradation[count] = new GradationCharColumn()
                    {
                        Serie = item,
                        Segmento = string.Format("Chi Quầy {0}/{1}", month - 1, year),
                        Valor1 = sumDSChiQuayLastMonth,
                        Tooltip = tooltip
                    };

                    count++;
                    arrayGradation[count] = new GradationCharColumn()
                    {
                        Serie = item,
                        Segmento = string.Format("Chi Nhà {0}/{1}", month - 1, year),
                        Valor1 = sumDSChiNhaLastMonth,
                        Tooltip = tooltip
                    };

                    count++;
                    arrayGradation[count] = new GradationCharColumn()
                    {
                        Serie = item,
                        Segmento = string.Format("DSCK {0}/{1}", month - 1, year),
                        Valor1 = sumDSCKLastMonth,
                        Tooltip = tooltip
                    };
                    count++;

                    // Cùng kì năm trước
                    arrayGradation[count] = new GradationCharColumn()
                    {
                        Serie = item,
                        Segmento = string.Format("Chi Quầy {0}/{1}", month, year - 1),
                        Valor1 = sumDSChiQuayLastYear,
                        Tooltip = tooltip
                    };

                    count++;
                    arrayGradation[count] = new GradationCharColumn()
                    {
                        Serie = item,
                        Segmento = string.Format("Chi Nhà {0}/{1}", month, year - 1),
                        Valor1 = sumDSChiNhaLastYear,
                        Tooltip = tooltip
                    };

                    count++;
                    arrayGradation[count] = new GradationCharColumn()
                    {
                        Serie = item,
                        Segmento = string.Format("DSCK {0}/{1}", month, year - 1),
                        Valor1 = sumDSCKLastYear,
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
        /// Get data cho việc vẽ biểu đồ cột cho so sánh giai đoạn
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult SearchColumnsChartCompareMonthForAll([DataSourceRequest]DataSourceRequest request, int month, int year, string reportTypeID, string marketID)
        {
            List<ReportDetailtSTMarket> listDataCompareMonth = new ReportBL().ReportDetailtCompareMonthForAll(year, month, reportTypeID, marketID);

            GradationCompare[] arrayGradation = null;

            // Trường hợp tất cả thị trường
            if (marketID.Equals("0"))
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
                int countArray = 9;
                arrayGradation = new GradationCompare[countArray * listAsianItem.Count];
                int count = 0;

                List<ReportDetailtSTMarket> listDataYear = null;
                List<ReportDetailtSTMarket> listDataLastMonth = null;
                List<ReportDetailtSTMarket> listDataLastYear = null;

                foreach (string item in listAsianItem)
                {
                    listDataYear = listDataCompareMonth.Where(x => x.MarketName == item && x.Month == month.ToString() && x.Year == year.ToString()).ToList();
                    listDataLastMonth = listDataCompareMonth.Where(x => x.MarketName == item && x.Month == (month - 1).ToString() && x.Year == year.ToString()).ToList();
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
                        NameGradationCompare = item,
                        amount = sumDSChiQuayYear,
                        NameType = string.Format("Chi Quầy {0}/{1}", month, year),
                    };

                    count++;
                    arrayGradation[count] = new GradationCompare()
                    {
                        NameGradationCompare = item,
                        amount = sumDSChiNhaYear,
                        NameType = string.Format("Chi Nhà {0}/{1}", month, year),
                    };

                    count++;
                    arrayGradation[count] = new GradationCompare()
                    {
                        NameGradationCompare = item,
                        amount = sumDSCKYear,
                        NameType = string.Format("DSCK {0}/{1}", month, year)
                    };
                    count++;

                    // Tháng trước
                    arrayGradation[count] = new GradationCompare()
                    {
                        NameGradationCompare = item,
                        amount = sumDSChiQuayLastMonth,
                        NameType = string.Format("Chi Quầy {0}/{1}", month - 1, year),
                    };

                    count++;
                    arrayGradation[count] = new GradationCompare()
                    {
                        NameGradationCompare = item,
                        amount = sumDSChiNhaLastMonth,
                        NameType = string.Format("Chi Nhà {0}/{1}", month - 1, year),
                    };

                    count++;
                    arrayGradation[count] = new GradationCompare()
                    {
                        NameGradationCompare = item,
                        amount = sumDSCKLastMonth,
                        NameType = string.Format("DSCK {0}/{1}", month - 1, year),
                    };
                    count++;

                    // Cùng kì năm trước
                    arrayGradation[count] = new GradationCompare()
                    {
                        NameGradationCompare = item,
                        amount = sumDSChiQuayLastYear,
                        NameType = string.Format("Chi Quầy {0}/{1}", month, year - 1),
                    };

                    count++;
                    arrayGradation[count] = new GradationCompare()
                    {
                        NameGradationCompare = item,
                        amount = sumDSChiNhaLastYear,
                        NameType = string.Format("Chi Nhà {0}/{1}", month, year - 1),
                    };

                    count++;
                    arrayGradation[count] = new GradationCompare()
                    {
                        NameGradationCompare = item,
                        amount = sumDSCKLastYear,
                        NameType = string.Format("DSCK {0}/{1}", month, year - 1),
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
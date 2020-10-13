// #################################################################
// # Copyright (C) 2010-2011, ASoft JSC.  All Rights Reserved.
// #
// # History：
// #	Date Time		Updated			Content
// #	04/09/2020		Truong Lam		Create New
// ##################################################################

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
    public class ReportDetailtPartnerLTController : Controller
    {
        // GET: Admin/ReportDetailtPartnerLT
        public ActionResult Index()
        {
            return View();
        }
        
        /// <summary>
        /// Màn hình báo cáo cho ngày
        /// </summary>
        /// <returns></returns>
        public ActionResult PartnerForTotal(DateTime? fromDay, DateTime? toDay, string reportTypeID)
        {
            string nameUrl = "Báo cáo chi tiết/Theo doanh số theo đối tác/Loại tiền/Tất cả/Theo ngày";
            ViewBag.NameURL = nameUrl;

            if (fromDay != null && toDay != null && reportTypeID != null)
            {
                List<string> listData = new List<string>()
                {
                    fromDay.Value.ToString("MM/dd/yyyy"),
                    toDay.Value.ToString("MM/dd/yyyy"),
                    reportTypeID
                };

                ViewData["listData"] = listData;
            }
            return View();
        }

        /// <summary>
        /// Màn hình báo cáo cho ngày
        /// </summary>
        /// <returns></returns>
        public ActionResult PartnerForTotalForMonth(DateTime? fromDate, DateTime? toDate, string reportTypeID)
        {
            string nameUrl = "Báo cáo chi tiết/Theo doanh số theo đối tác/Loại tiền/Tất cả/Theo tháng";
            ViewBag.NameURL = nameUrl;

            if (fromDate != null && toDate != null && reportTypeID != null)
            {
                List<string> listData = new List<string>()
                {
                    fromDate.Value.ToString("MM/dd/yyyy"),
                    toDate.Value.ToString("MM/dd/yyyy"),
                    reportTypeID
                };

                ViewData["listData"] = listData;
            }
            return View();
        }

        /// <summary>
        /// Màn hình báo cáo cho ngày
        /// </summary>
        /// <returns></returns>
        public ActionResult PartnerForTotalForYear(DateTime? fromDate, DateTime? toDate, string reportTypeID)
        {
            string nameUrl = "Báo cáo chi tiết/Theo doanh số theo đối tác/Loại tiền/Tất cả/Theo năm";
            ViewBag.NameURL = nameUrl;

            if (fromDate != null && toDate != null && reportTypeID != null)
            {
                List<string> listData = new List<string>()
                {
                    fromDate.Value.ToString("MM/dd/yyyy"),
                    toDate.Value.ToString("MM/dd/yyyy"),
                    reportTypeID
                };

                ViewData["listData"] = listData;
            }
            return View();
        }

        /// <summary>
        /// Màn hình báo cáo chi tiết theo ngày theo từng đối tác
        /// </summary>
        /// <returns></returns>
        public ActionResult PartnerForOne(DateTime? fromDay, DateTime? toDay, string reportTypeID, string partnerID)
        {
            string nameUrl = "Báo cáo chi tiết/Theo doanh số chi trả/Theo loại tiền/Loại hình dịch vụ/Từng thị trường";
            ViewBag.NameURL = nameUrl;

            if (fromDay != null && toDay != null && reportTypeID != null && partnerID != null)
            {
                List<string> listData = new List<string>()
                {
                    fromDay.Value.ToString("MM/dd/yyyy"),
                    toDay.Value.ToString("MM/dd/yyyy"),
                    reportTypeID,
                    partnerID
                };

                ViewData["listData"] = listData;
            }

            return View();
        }

        /// <summary>
        /// Màn hình báo cáo chi tiết theo ngày theo từng đối tác
        /// </summary>
        /// <returns></returns>
        public ActionResult PartnerForOneForMonth(DateTime? fromDate, DateTime? toDate, string reportTypeID, string partnerID)
        {
            string nameUrl = "Báo cáo chi tiết/Theo doanh số chi trả/Theo thị trường/Loại hình dịch vụ/Từng thị trường";
            ViewBag.NameURL = nameUrl;

            if (fromDate != null && toDate != null && reportTypeID != null && partnerID != null)
            {
                List<string> listData = new List<string>()
                {
                    fromDate.Value.ToString("MM/dd/yyyy"),
                    toDate.Value.ToString("MM/dd/yyyy"),
                    reportTypeID,
                    partnerID
                };

                ViewData["listData"] = listData;
            }

            return View();
        }

        /// <summary>
        /// Màn hình báo cáo chi tiết theo ngày theo từng đối tác
        /// </summary>
        /// <returns></returns>
        public ActionResult PartnerForOneForYear(DateTime? fromDate, DateTime? toDate, string reportTypeID, string partnerID)
        {
            string nameUrl = "Báo cáo chi tiết/Theo doanh số chi trả/Theo thị trường/Loại hình dịch vụ/Từng thị trường";
            ViewBag.NameURL = nameUrl;

            if (fromDate != null && toDate != null && reportTypeID != null && partnerID != null)
            {
                List<string> listData = new List<string>()
                {
                    fromDate.Value.ToString("MM/dd/yyyy"),
                    toDate.Value.ToString("MM/dd/yyyy"),
                    reportTypeID,
                    partnerID
                };

                ViewData["listData"] = listData;
            }

            return View();
        }

        public ActionResult ReportDetailtGradationCompare(string gradation, int? year, string reportTypeID)
        {
            string nameUrl = "Báo cáo chi tiết/Theo đối tác/Loại tiền chi trả/So sánh - Giai đoạn - Tất cả";
            ViewBag.NameURL = nameUrl;

            if (!string.IsNullOrEmpty(gradation))
            {
                if (int.Parse(gradation) > 0 && year > 0 && reportTypeID != null)
                {
                    List<string> listData = new List<string>()
                {
                    gradation,
                    year.ToString(),
                    reportTypeID
                };

                    ViewData["listData"] = listData;
                }
            }
            return View();
        }
        
        public ActionResult ReportDetailtGradationCompareForOne(string gradation, int? year, string reportTypeID, string partnerID)
        {
            string nameUrl = "Báo cáo chi tiết/Theo đối tác/Loại tiền chi trả/So sánh - Giai đoạn - Từng đối tác";
            ViewBag.NameURL = nameUrl;

            if (!string.IsNullOrEmpty(gradation))
            {
                if (int.Parse(gradation) > 0 && year > 0 && reportTypeID != null && partnerID != null)
                {
                    List<string> listData = new List<string>()
                {
                    gradation,
                    year.ToString(),
                    reportTypeID,
                    partnerID
                };

                    ViewData["listData"] = listData;
                }
            }
            return View();
        }

        public ActionResult ReportDetailtCompareMonthForAll(int? month, int? year, string reportTypeID)
        {
            string nameUrl = "Báo cáo chi tiết/Theo Đối tác/So sánh tháng/Tất cả";
            ViewBag.NameURL = nameUrl;

            if (month != null && year != null)
            {
                if (month.Value > 0 && year.Value > 0 && reportTypeID != null)
                {
                    List<string> listData = new List<string>()
                    {
                        month.ToString(),
                        year.ToString(),
                        reportTypeID
                    };

                    ViewData["listData"] = listData;
                }
            }

            return View();
        }

        public ActionResult ReportDetailtCompareMonthForOne(int? month, int? year, string reportTypeID, string partnerID)
        {
            string nameUrl = "Báo cáo chi tiết/Theo Đối tác/So sánh tháng/Từng đối tác";
            ViewBag.NameURL = nameUrl;

            if (month != null && year != null)
            {
                if (month.Value > 0 && year.Value > 0 && reportTypeID != null && partnerID != null)
                {
                    List<string> listData = new List<string>()
                    {
                        month.ToString(),
                        year.ToString(),
                        reportTypeID,
                        partnerID
                    };

                    ViewData["listData"] = listData;
                }
            }

            return View();
        }

        /// <summary>
        /// Search report day theo ngày nhập vào
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public ActionResult SearchPartnerForTotal([DataSourceRequest]DataSourceRequest request, DateTime fromDay, DateTime toDay, string reportTypeID)
        {
            List<ReportDetailtForTotalMoneyType> listData = new ReportBL().SearchReportDetailtPartnerLTForDay(fromDay, toDay, reportTypeID);
            List<ReportDetailtForTotalMoneyType> listDataConvert = new ReportBL().SearchReportDetailtPartnerLTForDayConvert(fromDay, toDay, reportTypeID);

            List<ReportDetailtForTotalMoneyType> listDataTotal = new List<ReportDetailtForTotalMoneyType>();
            
            foreach(ReportDetailtForTotalMoneyType item in listData)
            {
                listDataTotal.Add(
                    new ReportDetailtForTotalMoneyType()
                    {
                        PartnerCode = item.PartnerCode,
                        PartnerName = item.PartnerName,
                        VND = item.VND,
                        USD = item.USD,
                        EUR = item.EUR,
                        CAD = item.CAD,
                        AUD = item.AUD,
                        GBP = item.GBP,
                        typeID = 0
                    }
                );
            }

            foreach (ReportDetailtForTotalMoneyType item in listDataConvert)
            {
                listDataTotal.Add(
                    new ReportDetailtForTotalMoneyType()
                    {
                        PartnerCode = item.PartnerCode,
                        PartnerName = item.PartnerName,
                        VND = item.VND,
                        USD = item.USD,
                        EUR = item.EUR,
                        CAD = item.CAD,
                        AUD = item.AUD,
                        GBP = item.GBP,
                        typeID = 1
                    }
                );
            }

            // Khởi tạo datatable
            DataTable table = new DataTable();
            // Tạo các cột cho datatable
            table.Columns.Add("STT", typeof(String));
            table.Columns.Add("PartnerName", typeof(String));

            table.Columns.Add("VND1", typeof(double));
            table.Columns.Add("USD1", typeof(double));
            table.Columns.Add("EUR1", typeof(double));
            table.Columns.Add("CAD1", typeof(double));
            table.Columns.Add("AUD1", typeof(double));
            table.Columns.Add("GBP1", typeof(double));


            table.Columns.Add("VND2", typeof(double));
            table.Columns.Add("USD2", typeof(double));
            table.Columns.Add("EUR2", typeof(double));
            table.Columns.Add("CAD2", typeof(double));
            table.Columns.Add("AUD2", typeof(double));
            table.Columns.Add("GBP2", typeof(double));

            table.Columns.Add("TDS2", typeof(double));

            List<string> listPartner = new List<string>();
            int count = 1;

            foreach(ReportDetailtForTotalMoneyType item in listDataTotal)
            {
                if(listPartner.Contains(item.PartnerCode))
                {
                    continue;
                }

                listPartner.Add(item.PartnerCode);

                // Nguyên tệ
                ReportDetailtForTotalMoneyType dataItem = listDataTotal.Find(x => x.PartnerCode == item.PartnerCode && x.typeID == 0);
                // Quy USD
                ReportDetailtForTotalMoneyType dataItemConvert = listDataTotal.Find(x => x.PartnerCode == item.PartnerCode && x.typeID == 1);
                
                if (dataItem == null)
                {
                    dataItem = new ReportDetailtForTotalMoneyType()
                    {
                        PartnerCode = item.PartnerCode,
                        PartnerName = item.PartnerName
                    };
                }

                if (dataItemConvert == null)
                {
                    dataItemConvert = new ReportDetailtForTotalMoneyType()
                    {
                        PartnerCode = item.PartnerCode,
                        PartnerName = item.PartnerName
                    };
                }
                else
                {
                    dataItemConvert.TongDS = dataItemConvert.VND + dataItemConvert.USD + dataItemConvert.EUR + dataItemConvert.CAD + dataItemConvert.AUD + dataItemConvert.GBP;
                }

                table.Rows.Add(
                    count++
                    , item.PartnerName
                    , dataItem.VND, dataItem.USD, dataItem.EUR, dataItem.CAD, dataItem.AUD, dataItem.GBP
                    , dataItemConvert.VND, dataItemConvert.USD, dataItemConvert.EUR, dataItemConvert.CAD, dataItemConvert.AUD, dataItemConvert.GBP, dataItemConvert.TongDS
                );
            }
            
            DataRow row = table.NewRow();
            row["STT"] = "";
            row["PartnerName"] = "Tổng";
            row["VND1"] = table.Compute("Sum(VND1)", "");
            row["USD1"] = table.Compute("Sum(USD1)", "");
            row["EUR1"] = table.Compute("Sum(EUR1)", "");
            row["CAD1"] = table.Compute("Sum(CAD1)", "");
            row["AUD1"] = table.Compute("Sum(AUD1)", "");
            row["GBP1"] = table.Compute("Sum(GBP1)", "");


            row["VND2"] = table.Compute("Sum(VND2)", "");
            row["USD2"] = table.Compute("Sum(USD2)", "");
            row["EUR2"] = table.Compute("Sum(EUR2)", "");
            row["CAD2"] = table.Compute("Sum(CAD2)", "");
            row["AUD2"] = table.Compute("Sum(AUD2)", "");
            row["GBP2"] = table.Compute("Sum(GBP2)", "");

            row["TDS2"] = table.Compute("Sum(TDS2)", "");
            table.Rows.Add(row);


            return Json(table.ToDataSourceResult(request), JsonRequestBehavior.AllowGet);
        }


        /// <summary>
        /// Search report day theo ngày nhập vào
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public ActionResult SearchPartnerForTotalForMonth([DataSourceRequest]DataSourceRequest request, DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            List<ReportDetailtForTotalMoneyType> listData = new ReportBL().SearchReportDetailtPartnerLTForMonth(fromDate, toDate, reportTypeID);
            List<ReportDetailtForTotalMoneyType> listDataConvert = new ReportBL().SearchReportDetailtPartnerLTForMonthConvert(fromDate, toDate, reportTypeID);

            List<ReportDetailtForTotalMoneyType> listDataTotal = new List<ReportDetailtForTotalMoneyType>();

            foreach (ReportDetailtForTotalMoneyType item in listData)
            {
                listDataTotal.Add(
                    new ReportDetailtForTotalMoneyType()
                    {
                        PartnerCode = item.PartnerCode,
                        PartnerName = item.PartnerName,
                        VND = item.VND,
                        USD = item.USD,
                        EUR = item.EUR,
                        CAD = item.CAD,
                        AUD = item.AUD,
                        GBP = item.GBP,
                        typeID = 0
                    }
                );
            }

            foreach (ReportDetailtForTotalMoneyType item in listDataConvert)
            {
                listDataTotal.Add(
                    new ReportDetailtForTotalMoneyType()
                    {
                        PartnerCode = item.PartnerCode,
                        PartnerName = item.PartnerName,
                        VND = item.VND,
                        USD = item.USD,
                        EUR = item.EUR,
                        CAD = item.CAD,
                        AUD = item.AUD,
                        GBP = item.GBP,
                        typeID = 1
                    }
                );
            }

            // Khởi tạo datatable
            DataTable table = new DataTable();
            // Tạo các cột cho datatable
            table.Columns.Add("STT", typeof(String));
            table.Columns.Add("PartnerName", typeof(String));

            table.Columns.Add("VND1", typeof(double));
            table.Columns.Add("USD1", typeof(double));
            table.Columns.Add("EUR1", typeof(double));
            table.Columns.Add("CAD1", typeof(double));
            table.Columns.Add("AUD1", typeof(double));
            table.Columns.Add("GBP1", typeof(double));


            table.Columns.Add("VND2", typeof(double));
            table.Columns.Add("USD2", typeof(double));
            table.Columns.Add("EUR2", typeof(double));
            table.Columns.Add("CAD2", typeof(double));
            table.Columns.Add("AUD2", typeof(double));
            table.Columns.Add("GBP2", typeof(double));

            table.Columns.Add("TDS2", typeof(double));

            List<string> listPartner = new List<string>();
            int count = 1;

            foreach (ReportDetailtForTotalMoneyType item in listDataTotal)
            {
                if (listPartner.Contains(item.PartnerCode))
                {
                    continue;
                }

                listPartner.Add(item.PartnerCode);

                // Nguyên tệ
                ReportDetailtForTotalMoneyType dataItem = listDataTotal.Find(x => x.PartnerCode == item.PartnerCode && x.typeID == 0);
                // Quy USD
                ReportDetailtForTotalMoneyType dataItemConvert = listDataTotal.Find(x => x.PartnerCode == item.PartnerCode && x.typeID == 1);

                if (dataItem == null)
                {
                    dataItem = new ReportDetailtForTotalMoneyType()
                    {
                        PartnerCode = item.PartnerCode,
                        PartnerName = item.PartnerName
                    };
                }

                if (dataItemConvert == null)
                {
                    dataItemConvert = new ReportDetailtForTotalMoneyType()
                    {
                        PartnerCode = item.PartnerCode,
                        PartnerName = item.PartnerName
                    };
                }
                else
                {
                    dataItemConvert.TongDS = dataItemConvert.VND + dataItemConvert.USD + dataItemConvert.EUR + dataItemConvert.CAD + dataItemConvert.AUD + dataItemConvert.GBP;
                }

                table.Rows.Add(
                    count++
                    , item.PartnerName
                    , dataItem.VND, dataItem.USD, dataItem.EUR, dataItem.CAD, dataItem.AUD, dataItem.GBP
                    , dataItemConvert.VND, dataItemConvert.USD, dataItemConvert.EUR, dataItemConvert.CAD, dataItemConvert.AUD, dataItemConvert.GBP, dataItemConvert.TongDS
                );
            }

            DataRow row = table.NewRow();
            row["STT"] = "";
            row["PartnerName"] = "Tổng";
            row["VND1"] = table.Compute("Sum(VND1)", "");
            row["USD1"] = table.Compute("Sum(USD1)", "");
            row["EUR1"] = table.Compute("Sum(EUR1)", "");
            row["CAD1"] = table.Compute("Sum(CAD1)", "");
            row["AUD1"] = table.Compute("Sum(AUD1)", "");
            row["GBP1"] = table.Compute("Sum(GBP1)", "");


            row["VND2"] = table.Compute("Sum(VND2)", "");
            row["USD2"] = table.Compute("Sum(USD2)", "");
            row["EUR2"] = table.Compute("Sum(EUR2)", "");
            row["CAD2"] = table.Compute("Sum(CAD2)", "");
            row["AUD2"] = table.Compute("Sum(AUD2)", "");
            row["GBP2"] = table.Compute("Sum(GBP2)", "");

            row["TDS2"] = table.Compute("Sum(TDS2)", "");
            table.Rows.Add(row);

            return Json(table.ToDataSourceResult(request), JsonRequestBehavior.AllowGet);
        }


        /// <summary>
        /// Search report day theo ngày nhập vào
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public ActionResult SearchPartnerForTotalForYear([DataSourceRequest]DataSourceRequest request, DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            List<ReportDetailtForTotalMoneyType> listData = new ReportBL().SearchReportDetailtPartnerLTForYear(fromDate, toDate, reportTypeID);
            List<ReportDetailtForTotalMoneyType> listDataConvert = new ReportBL().SearchReportDetailtPartnerLTForYearConvert(fromDate, toDate, reportTypeID);

            List<ReportDetailtForTotalMoneyType> listDataTotal = new List<ReportDetailtForTotalMoneyType>();

            foreach (ReportDetailtForTotalMoneyType item in listData)
            {
                listDataTotal.Add(
                    new ReportDetailtForTotalMoneyType()
                    {
                        PartnerCode = item.PartnerCode,
                        PartnerName = item.PartnerName,
                        VND = item.VND,
                        USD = item.USD,
                        EUR = item.EUR,
                        CAD = item.CAD,
                        AUD = item.AUD,
                        GBP = item.GBP,
                        typeID = 0
                    }
                );
            }

            foreach (ReportDetailtForTotalMoneyType item in listDataConvert)
            {
                listDataTotal.Add(
                    new ReportDetailtForTotalMoneyType()
                    {
                        PartnerCode = item.PartnerCode,
                        PartnerName = item.PartnerName,
                        VND = item.VND,
                        USD = item.USD,
                        EUR = item.EUR,
                        CAD = item.CAD,
                        AUD = item.AUD,
                        GBP = item.GBP,
                        typeID = 1
                    }
                );
            }

            // Khởi tạo datatable
            DataTable table = new DataTable();
            // Tạo các cột cho datatable
            table.Columns.Add("STT", typeof(String));
            table.Columns.Add("PartnerName", typeof(String));

            table.Columns.Add("VND1", typeof(double));
            table.Columns.Add("USD1", typeof(double));
            table.Columns.Add("EUR1", typeof(double));
            table.Columns.Add("CAD1", typeof(double));
            table.Columns.Add("AUD1", typeof(double));
            table.Columns.Add("GBP1", typeof(double));


            table.Columns.Add("VND2", typeof(double));
            table.Columns.Add("USD2", typeof(double));
            table.Columns.Add("EUR2", typeof(double));
            table.Columns.Add("CAD2", typeof(double));
            table.Columns.Add("AUD2", typeof(double));
            table.Columns.Add("GBP2", typeof(double));

            table.Columns.Add("TDS2", typeof(double));

            List<string> listPartner = new List<string>();
            int count = 1;

            foreach (ReportDetailtForTotalMoneyType item in listDataTotal)
            {
                if (listPartner.Contains(item.PartnerCode))
                {
                    continue;
                }

                listPartner.Add(item.PartnerCode);

                // Nguyên tệ
                ReportDetailtForTotalMoneyType dataItem = listDataTotal.Find(x => x.PartnerCode == item.PartnerCode && x.typeID == 0);
                // Quy USD
                ReportDetailtForTotalMoneyType dataItemConvert = listDataTotal.Find(x => x.PartnerCode == item.PartnerCode && x.typeID == 1);

                if (dataItem == null)
                {
                    dataItem = new ReportDetailtForTotalMoneyType()
                    {
                        PartnerCode = item.PartnerCode,
                        PartnerName = item.PartnerName
                    };
                }

                if (dataItemConvert == null)
                {
                    dataItemConvert = new ReportDetailtForTotalMoneyType()
                    {
                        PartnerCode = item.PartnerCode,
                        PartnerName = item.PartnerName
                    };
                }
                else
                {
                    dataItemConvert.TongDS = dataItemConvert.VND + dataItemConvert.USD + dataItemConvert.EUR + dataItemConvert.CAD + dataItemConvert.AUD + dataItemConvert.GBP;
                }

                table.Rows.Add(
                    count++
                    , item.PartnerName
                    , dataItem.VND, dataItem.USD, dataItem.EUR, dataItem.CAD, dataItem.AUD, dataItem.GBP
                    , dataItemConvert.VND, dataItemConvert.USD, dataItemConvert.EUR, dataItemConvert.CAD, dataItemConvert.AUD, dataItemConvert.GBP, dataItemConvert.TongDS
                );
            }
            
            DataRow row = table.NewRow();
            row["STT"] = "";
            row["PartnerName"] = "Tổng";
            row["VND1"] = table.Compute("Sum(VND1)", "");
            row["USD1"] = table.Compute("Sum(USD1)", "");
            row["EUR1"] = table.Compute("Sum(EUR1)", "");
            row["CAD1"] = table.Compute("Sum(CAD1)", "");
            row["AUD1"] = table.Compute("Sum(AUD1)", "");
            row["GBP1"] = table.Compute("Sum(GBP1)", "");


            row["VND2"] = table.Compute("Sum(VND2)", "");
            row["USD2"] = table.Compute("Sum(USD2)", "");
            row["EUR2"] = table.Compute("Sum(EUR2)", "");
            row["CAD2"] = table.Compute("Sum(CAD2)", "");
            row["AUD2"] = table.Compute("Sum(AUD2)", "");
            row["GBP2"] = table.Compute("Sum(GBP2)", "");

            row["TDS2"] = table.Compute("Sum(TDS2)", "");
            table.Rows.Add(row);


            return Json(table.ToDataSourceResult(request), JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// Search report chi tiết theo ngày cho từng đối tác
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/09/2020]
        /// </history>
        public ActionResult SearchPartnerForOne([DataSourceRequest]DataSourceRequest request, DateTime fromDay, DateTime toDay, string reportTypeID, string partnerID)
        {
            List<ReportDetailtForTotalMoneyType> listData = new ReportBL().SearchPartnerLTForOne(fromDay, toDay, reportTypeID, partnerID);
            List<ReportDetailtForTotalMoneyType> listDataConvert = new ReportBL().SearchPartnerLTForOneConvert(fromDay, toDay, reportTypeID, partnerID);
            List<ReportDetailtForTotalMoneyType> listDataTotal = new List<ReportDetailtForTotalMoneyType>();

            foreach (ReportDetailtForTotalMoneyType item in listData)
            {
                listDataTotal.Add(
                    new ReportDetailtForTotalMoneyType()
                    {
                        PartnerCode = item.PartnerCode,
                        PartnerName = item.PartnerName,
                        VND = item.VND,
                        USD = item.USD,
                        EUR = item.EUR,
                        CAD = item.CAD,
                        AUD = item.AUD,
                        GBP = item.GBP,
                        typeID = 0,
                        CreatedDate = item.CreatedDate
                    }
                );
            }

            foreach (ReportDetailtForTotalMoneyType item in listDataConvert)
            {
                listDataTotal.Add(
                    new ReportDetailtForTotalMoneyType()
                    {
                        PartnerCode = item.PartnerCode,
                        PartnerName = item.PartnerName,
                        VND = item.VND,
                        USD = item.USD,
                        EUR = item.EUR,
                        CAD = item.CAD,
                        AUD = item.AUD,
                        GBP = item.GBP,
                        typeID = 1,
                        CreatedDate = item.CreatedDate
                    }
                );
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


            table.Columns.Add("VND2", typeof(double));
            table.Columns.Add("USD2", typeof(double));
            table.Columns.Add("EUR2", typeof(double));
            table.Columns.Add("CAD2", typeof(double));
            table.Columns.Add("AUD2", typeof(double));
            table.Columns.Add("GBP2", typeof(double));

            table.Columns.Add("TDS2", typeof(double));

            List<string> listPartner = new List<string>();

            foreach (ReportDetailtForTotalMoneyType item in listDataTotal)
            {
                if (listPartner.Contains(item.CreatedDate.ToString()))
                {
                    continue;
                }

                listPartner.Add(item.CreatedDate.ToString());

                // Nguyên tệ
                ReportDetailtForTotalMoneyType dataItem = listDataTotal.Find(x => x.CreatedDate == item.CreatedDate && x.typeID == 0);
                // Quy USD
                ReportDetailtForTotalMoneyType dataItemConvert = listDataTotal.Find(x => x.CreatedDate == item.CreatedDate && x.typeID == 1);

                if (dataItem == null)
                {
                    dataItem = new ReportDetailtForTotalMoneyType()
                    {
                        PartnerCode = item.PartnerCode,
                        PartnerName = item.PartnerName
                    };
                }

                if (dataItemConvert == null)
                {
                    dataItemConvert = new ReportDetailtForTotalMoneyType()
                    {
                        PartnerCode = item.PartnerCode,
                        PartnerName = item.PartnerName
                    };
                }
                else
                {
                    dataItemConvert.TongDS = dataItemConvert.VND + dataItemConvert.USD + dataItemConvert.EUR + dataItemConvert.CAD + dataItemConvert.AUD + dataItemConvert.GBP;
                }

                table.Rows.Add(
                    item.CreatedDate.ToString("dd/MM/yyyy")
                    , dataItem.VND, dataItem.USD, dataItem.EUR, dataItem.CAD, dataItem.AUD, dataItem.GBP
                    , dataItemConvert.VND, dataItemConvert.USD, dataItemConvert.EUR, dataItemConvert.CAD, dataItemConvert.AUD, dataItemConvert.GBP, dataItemConvert.TongDS
                );
            }

            DataRow row = table.NewRow();
            row["PartnerName"] = "Tổng";
            row["VND1"] = table.Compute("Sum(VND1)", "");
            row["USD1"] = table.Compute("Sum(USD1)", "");
            row["EUR1"] = table.Compute("Sum(EUR1)", "");
            row["CAD1"] = table.Compute("Sum(CAD1)", "");
            row["AUD1"] = table.Compute("Sum(AUD1)", "");
            row["GBP1"] = table.Compute("Sum(GBP1)", "");


            row["VND2"] = table.Compute("Sum(VND2)", "");
            row["USD2"] = table.Compute("Sum(USD2)", "");
            row["EUR2"] = table.Compute("Sum(EUR2)", "");
            row["CAD2"] = table.Compute("Sum(CAD2)", "");
            row["AUD2"] = table.Compute("Sum(AUD2)", "");
            row["GBP2"] = table.Compute("Sum(GBP2)", "");

            row["TDS2"] = table.Compute("Sum(TDS2)", "");
            table.Rows.Add(row);

            return Json(table.ToDataSourceResult(request), JsonRequestBehavior.AllowGet);
        }


        /// <summary>
        /// Get data cho việc vẽ biểu đồ cột cho so sánh theo tháng và cùng kì năm trước
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult SearchChartPartnerForOne([DataSourceRequest]DataSourceRequest request, DateTime fromDay, DateTime toDay, string reportTypeID, string partnerID)
        {
            List<ReportDetailtForTotalMoneyType> listDataConvert = new ReportBL().SearchPartnerLTForOneConvert(fromDay, toDay, reportTypeID, partnerID);
            
            GradationCompare[] arrayGradation = null;

            if (listDataConvert.Count() > 0)
            {

                arrayGradation = new GradationCompare[6 * listDataConvert.Count()];
                int count = 0;
                foreach (ReportDetailtForTotalMoneyType item in listDataConvert)
                {
                    // tổng doanh số
                    item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;

                    // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                    arrayGradation[count] = new GradationCompare()
                    {
                        NameGradationCompare = "VND",
                        amount = item.VND,
                        NameType = item.CreatedDate.ToString("dd/MM/yyyy")
                    };

                    count++;
                    arrayGradation[count] = new GradationCompare()
                    {
                        NameGradationCompare = "USD",
                        amount = item.USD,
                        NameType = item.CreatedDate.ToString("dd/MM/yyyy")
                    };

                    count++;
                    arrayGradation[count] = new GradationCompare()
                    {
                        NameGradationCompare = "EUR",
                        amount = item.EUR,
                        NameType = item.CreatedDate.ToString("dd/MM/yyyy")
                    };
                    
                    count++;

                    arrayGradation[count] = new GradationCompare()
                    {
                        NameGradationCompare = "CAD",
                        amount = item.CAD,
                        NameType = item.CreatedDate.ToString("dd/MM/yyyy")
                    };

                    count++;

                    arrayGradation[count] = new GradationCompare()
                    {
                        NameGradationCompare = "AUD",
                        amount = item.AUD,
                        NameType = item.CreatedDate.ToString("dd/MM/yyyy")
                    };

                    count++;

                    arrayGradation[count] = new GradationCompare()
                    {
                        NameGradationCompare = "GBP",
                        amount = item.GBP,
                        NameType = item.CreatedDate.ToString("dd/MM/yyyy")
                    };

                    count++;
                }


            }
            else
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
        /// Search report chi tiết theo ngày cho từng đối tác
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/09/2020]
        /// </history>
        public ActionResult SearchPartnerForOneForMonth([DataSourceRequest]DataSourceRequest request, DateTime fromDate, DateTime toDate, string reportTypeID, string partnerID)
        {
            List<ReportDetailtForTotalMoneyType> listData = new ReportBL().SearchPartnerLTForOneForMonth(fromDate, toDate, reportTypeID, partnerID);
            List<ReportDetailtForTotalMoneyType> listDataConvert = new ReportBL().SearchPartnerLTForOneForMonthConvert(fromDate, toDate, reportTypeID, partnerID);
            List<ReportDetailtForTotalMoneyType> listDataTotal = new List<ReportDetailtForTotalMoneyType>();

            foreach (ReportDetailtForTotalMoneyType item in listData)
            {
                listDataTotal.Add(
                    new ReportDetailtForTotalMoneyType()
                    {
                        PartnerCode = item.PartnerCode,
                        PartnerName = item.PartnerName,
                        VND = item.VND,
                        USD = item.USD,
                        EUR = item.EUR,
                        CAD = item.CAD,
                        AUD = item.AUD,
                        GBP = item.GBP,
                        typeID = 0,
                        Month = item.Month,
                        Year = item.Year
                    }
                );
            }

            foreach (ReportDetailtForTotalMoneyType item in listDataConvert)
            {
                listDataTotal.Add(
                    new ReportDetailtForTotalMoneyType()
                    {
                        PartnerCode = item.PartnerCode,
                        PartnerName = item.PartnerName,
                        VND = item.VND,
                        USD = item.USD,
                        EUR = item.EUR,
                        CAD = item.CAD,
                        AUD = item.AUD,
                        GBP = item.GBP,
                        typeID = 1,
                        Month = item.Month,
                        Year = item.Year
                    }
                );
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


            table.Columns.Add("VND2", typeof(double));
            table.Columns.Add("USD2", typeof(double));
            table.Columns.Add("EUR2", typeof(double));
            table.Columns.Add("CAD2", typeof(double));
            table.Columns.Add("AUD2", typeof(double));
            table.Columns.Add("GBP2", typeof(double));

            table.Columns.Add("TDS2", typeof(double));

            List<string> listPartner = new List<string>();

            foreach (ReportDetailtForTotalMoneyType item in listDataTotal)
            {
                string value = string.Format("Tháng {0}/{1}", item.Month, item.Year);
                if (listPartner.Contains(value))
                {
                    continue;
                }

                listPartner.Add(value);

                // Nguyên tệ
                ReportDetailtForTotalMoneyType dataItem = listDataTotal.Find(x => x.Month == item.Month && x.Year == item.Year && x.typeID == 0);
                // Quy USD
                ReportDetailtForTotalMoneyType dataItemConvert = listDataTotal.Find(x => x.Month == item.Month && x.Year == item.Year && x.typeID == 1);

                if (dataItem == null)
                {
                    dataItem = new ReportDetailtForTotalMoneyType()
                    {
                        PartnerCode = item.PartnerCode,
                        PartnerName = item.PartnerName
                    };
                }

                if (dataItemConvert == null)
                {
                    dataItemConvert = new ReportDetailtForTotalMoneyType()
                    {
                        PartnerCode = item.PartnerCode,
                        PartnerName = item.PartnerName
                    };
                }
                else
                {
                    dataItemConvert.TongDS = dataItemConvert.VND + dataItemConvert.USD + dataItemConvert.EUR + dataItemConvert.CAD + dataItemConvert.AUD + dataItemConvert.GBP;
                }

                table.Rows.Add(
                    string.Format("Tháng {0}/{1}", item.Month, item.Year)
                    , dataItem.VND, dataItem.USD, dataItem.EUR, dataItem.CAD, dataItem.AUD, dataItem.GBP
                    , dataItemConvert.VND, dataItemConvert.USD, dataItemConvert.EUR, dataItemConvert.CAD, dataItemConvert.AUD, dataItemConvert.GBP, dataItemConvert.TongDS
                );
            }

            DataRow row = table.NewRow();
            row["PartnerName"] = "Tổng";
            row["VND1"] = table.Compute("Sum(VND1)", "");
            row["USD1"] = table.Compute("Sum(USD1)", "");
            row["EUR1"] = table.Compute("Sum(EUR1)", "");
            row["CAD1"] = table.Compute("Sum(CAD1)", "");
            row["AUD1"] = table.Compute("Sum(AUD1)", "");
            row["GBP1"] = table.Compute("Sum(GBP1)", "");


            row["VND2"] = table.Compute("Sum(VND2)", "");
            row["USD2"] = table.Compute("Sum(USD2)", "");
            row["EUR2"] = table.Compute("Sum(EUR2)", "");
            row["CAD2"] = table.Compute("Sum(CAD2)", "");
            row["AUD2"] = table.Compute("Sum(AUD2)", "");
            row["GBP2"] = table.Compute("Sum(GBP2)", "");

            row["TDS2"] = table.Compute("Sum(TDS2)", "");
            table.Rows.Add(row);

            return Json(table.ToDataSourceResult(request), JsonRequestBehavior.AllowGet);
        }


        /// <summary>
        /// Get data cho việc vẽ biểu đồ cột cho so sánh theo tháng và cùng kì năm trước
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult SearchChartPartnerForOneForMonth([DataSourceRequest]DataSourceRequest request, DateTime fromDate, DateTime toDate, string reportTypeID, string partnerID)
        {
            List<ReportDetailtForTotalMoneyType> listDataConvert = new ReportBL().SearchPartnerLTForOneForMonthConvert(fromDate, toDate, reportTypeID, partnerID);

            GradationCompare[] arrayGradation = null;

            if (listDataConvert.Count() > 0)
            {

                arrayGradation = new GradationCompare[6 * listDataConvert.Count()];
                int count = 0;
                foreach (ReportDetailtForTotalMoneyType item in listDataConvert)
                {
                    // tổng doanh số
                    item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;

                    // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                    arrayGradation[count] = new GradationCompare()
                    {
                        NameGradationCompare = "VND",
                        amount = item.VND,
                        NameType = string.Format("Tháng {0}/{1}", item.Month, item.Year)
                    };

                    count++;
                    arrayGradation[count] = new GradationCompare()
                    {
                        NameGradationCompare = "USD",
                        amount = item.USD,
                        NameType = string.Format("Tháng {0}/{1}", item.Month, item.Year)
                    };

                    count++;
                    arrayGradation[count] = new GradationCompare()
                    {
                        NameGradationCompare = "EUR",
                        amount = item.EUR,
                        NameType = string.Format("Tháng {0}/{1}", item.Month, item.Year)
                    };

                    count++;

                    arrayGradation[count] = new GradationCompare()
                    {
                        NameGradationCompare = "CAD",
                        amount = item.CAD,
                        NameType = string.Format("Tháng {0}/{1}", item.Month, item.Year)
                    };

                    count++;

                    arrayGradation[count] = new GradationCompare()
                    {
                        NameGradationCompare = "AUD",
                        amount = item.AUD,
                        NameType = string.Format("Tháng {0}/{1}", item.Month, item.Year)
                    };

                    count++;

                    arrayGradation[count] = new GradationCompare()
                    {
                        NameGradationCompare = "GBP",
                        amount = item.GBP,
                        NameType = string.Format("Tháng {0}/{1}", item.Month, item.Year)
                    };

                    count++;
                }


            }
            else
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
        /// Get data cho việc vẽ biểu đồ cột cho so sánh theo tháng và cùng kì năm trước
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult SearchChartPartnerForOneForYear([DataSourceRequest]DataSourceRequest request, DateTime fromDate, DateTime toDate, string reportTypeID, string partnerID)
        {
            List<ReportDetailtForTotalMoneyType> listDataConvert = new ReportBL().SearchPartnerLTForOneForYearConvert(fromDate, toDate, reportTypeID, partnerID);

            GradationCompare[] arrayGradation = null;

            if (listDataConvert.Count() > 0)
            {

                arrayGradation = new GradationCompare[6 * listDataConvert.Count()];
                int count = 0;
                foreach (ReportDetailtForTotalMoneyType item in listDataConvert)
                {
                    // tổng doanh số
                    item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;

                    // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                    arrayGradation[count] = new GradationCompare()
                    {
                        NameGradationCompare = "VND",
                        amount = item.VND,
                        NameType = string.Format("Năm {0}", item.Year)
                    };

                    count++;
                    arrayGradation[count] = new GradationCompare()
                    {
                        NameGradationCompare = "USD",
                        amount = item.USD,
                        NameType = string.Format("Năm {0}", item.Year)
                    };

                    count++;
                    arrayGradation[count] = new GradationCompare()
                    {
                        NameGradationCompare = "EUR",
                        amount = item.EUR,
                        NameType = string.Format("Năm {0}", item.Year)
                    };

                    count++;

                    arrayGradation[count] = new GradationCompare()
                    {
                        NameGradationCompare = "CAD",
                        amount = item.CAD,
                        NameType = string.Format("Năm {0}", item.Year)
                    };

                    count++;

                    arrayGradation[count] = new GradationCompare()
                    {
                        NameGradationCompare = "AUD",
                        amount = item.AUD,
                        NameType = string.Format("Năm {0}", item.Year)
                    };

                    count++;

                    arrayGradation[count] = new GradationCompare()
                    {
                        NameGradationCompare = "GBP",
                        amount = item.GBP,
                        NameType = string.Format("Năm {0}", item.Year)
                    };

                    count++;
                }


            }
            else
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
        /// Search report chi tiết theo ngày cho từng đối tác
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/09/2020]
        /// </history>
        public ActionResult SearchPartnerForOneForYear([DataSourceRequest]DataSourceRequest request, DateTime fromDate, DateTime toDate, string reportTypeID, string partnerID)
        {
            List<ReportDetailtForTotalMoneyType> listData = new ReportBL().SearchPartnerLTForOneForYear(fromDate, toDate, reportTypeID, partnerID);
            List<ReportDetailtForTotalMoneyType> listDataConvert = new ReportBL().SearchPartnerLTForOneForYearConvert(fromDate, toDate, reportTypeID, partnerID);
            List<ReportDetailtForTotalMoneyType> listDataTotal = new List<ReportDetailtForTotalMoneyType>();

            foreach (ReportDetailtForTotalMoneyType item in listData)
            {
                listDataTotal.Add(
                    new ReportDetailtForTotalMoneyType()
                    {
                        PartnerCode = item.PartnerCode,
                        PartnerName = item.PartnerName,
                        VND = item.VND,
                        USD = item.USD,
                        EUR = item.EUR,
                        CAD = item.CAD,
                        AUD = item.AUD,
                        GBP = item.GBP,
                        typeID = 0,
                        Month = item.Month,
                        Year = item.Year
                    }
                );
            }

            foreach (ReportDetailtForTotalMoneyType item in listDataConvert)
            {
                listDataTotal.Add(
                    new ReportDetailtForTotalMoneyType()
                    {
                        PartnerCode = item.PartnerCode,
                        PartnerName = item.PartnerName,
                        VND = item.VND,
                        USD = item.USD,
                        EUR = item.EUR,
                        CAD = item.CAD,
                        AUD = item.AUD,
                        GBP = item.GBP,
                        typeID = 1,
                        Month = item.Month,
                        Year = item.Year
                    }
                );
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


            table.Columns.Add("VND2", typeof(double));
            table.Columns.Add("USD2", typeof(double));
            table.Columns.Add("EUR2", typeof(double));
            table.Columns.Add("CAD2", typeof(double));
            table.Columns.Add("AUD2", typeof(double));
            table.Columns.Add("GBP2", typeof(double));

            table.Columns.Add("TDS2", typeof(double));

            List<string> listPartner = new List<string>();

            foreach (ReportDetailtForTotalMoneyType item in listDataTotal)
            {
                string value = string.Format("Tháng {0}/{1}", item.Month, item.Year);
                if (listPartner.Contains(value))
                {
                    continue;
                }

                listPartner.Add(value);

                // Nguyên tệ
                ReportDetailtForTotalMoneyType dataItem = listDataTotal.Find(x => x.Year == item.Year && x.typeID == 0);
                // Quy USD
                ReportDetailtForTotalMoneyType dataItemConvert = listDataTotal.Find(x => x.Year == item.Year && x.typeID == 1);

                if (dataItem == null)
                {
                    dataItem = new ReportDetailtForTotalMoneyType()
                    {
                        PartnerCode = item.PartnerCode,
                        PartnerName = item.PartnerName
                    };
                }

                if (dataItemConvert == null)
                {
                    dataItemConvert = new ReportDetailtForTotalMoneyType()
                    {
                        PartnerCode = item.PartnerCode,
                        PartnerName = item.PartnerName
                    };
                }
                else
                {
                    dataItemConvert.TongDS = dataItemConvert.VND + dataItemConvert.USD + dataItemConvert.EUR + dataItemConvert.CAD + dataItemConvert.AUD + dataItemConvert.GBP;
                }

                table.Rows.Add(
                    string.Format("Năm {0}",item.Year)
                    , dataItem.VND, dataItem.USD, dataItem.EUR, dataItem.CAD, dataItem.AUD, dataItem.GBP
                    , dataItemConvert.VND, dataItemConvert.USD, dataItemConvert.EUR, dataItemConvert.CAD, dataItemConvert.AUD, dataItemConvert.GBP, dataItemConvert.TongDS
                );
            }

            DataRow row = table.NewRow();
            row["PartnerName"] = "Tổng";
            row["VND1"] = table.Compute("Sum(VND1)", "");
            row["USD1"] = table.Compute("Sum(USD1)", "");
            row["EUR1"] = table.Compute("Sum(EUR1)", "");
            row["CAD1"] = table.Compute("Sum(CAD1)", "");
            row["AUD1"] = table.Compute("Sum(AUD1)", "");
            row["GBP1"] = table.Compute("Sum(GBP1)", "");


            row["VND2"] = table.Compute("Sum(VND2)", "");
            row["USD2"] = table.Compute("Sum(USD2)", "");
            row["EUR2"] = table.Compute("Sum(EUR2)", "");
            row["CAD2"] = table.Compute("Sum(CAD2)", "");
            row["AUD2"] = table.Compute("Sum(AUD2)", "");
            row["GBP2"] = table.Compute("Sum(GBP2)", "");

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
        public ActionResult SearchGridReportForGradation([DataSourceRequest]DataSourceRequest request, int gradation, int year, string reportTypeID)
        {
            List<ReportDetailtForTotalMoneyType> listDataGradation = new ReportBL().ReportDetailtPartnerLTGradationCompareForAll(year, gradation, reportTypeID);
            
            // Khởi tạo datatable
            DataTable table = new DataTable();
            // Tạo các cột cho datatable
            table.Columns.Add("STT", typeof(String));
            table.Columns.Add("PartnerName", typeof(String));

            table.Columns.Add("VND1", typeof(double));
            table.Columns.Add("VND2", typeof(double));
            table.Columns.Add("TDS1", typeof(double));

            table.Columns.Add("USD1", typeof(double));
            table.Columns.Add("USD2", typeof(double));
            table.Columns.Add("TDS2", typeof(double));

            table.Columns.Add("EUR1", typeof(double));
            table.Columns.Add("EUR2", typeof(double));
            table.Columns.Add("TDS3", typeof(double));

            table.Columns.Add("CAD1", typeof(double));
            table.Columns.Add("CAD2", typeof(double));
            table.Columns.Add("TDS4", typeof(double));

            table.Columns.Add("AUD1", typeof(double));
            table.Columns.Add("AUD2", typeof(double));
            table.Columns.Add("TDS5", typeof(double));

            table.Columns.Add("GBP1", typeof(double));
            table.Columns.Add("GBP2", typeof(double));
            table.Columns.Add("TDS6", typeof(double));

            int count = 1;
            List<string> listPartner = new List<string>();
            foreach (ReportDetailtForTotalMoneyType item in listDataGradation)
            {
                // Check trường hợp đã tồn tại đối tác
                if(listPartner.Contains(item.PartnerCode))
                {
                    continue;
                }

                listPartner.Add(item.PartnerCode);

                // Cùng kì
                ReportDetailtForTotalMoneyType dataItemLastYear = listDataGradation.Find(x => x.PartnerCode == item.PartnerCode && x.Year == (year - 1).ToString());
                ReportDetailtForTotalMoneyType dataItemYear = listDataGradation.Find(x => x.PartnerCode == item.PartnerCode && x.Year == year.ToString());

                // Last year
                if(dataItemLastYear == null)
                {
                    dataItemLastYear = new ReportDetailtForTotalMoneyType()
                    {
                        PartnerCode = item.PartnerCode,
                        PartnerName = item.PartnerName,
                        Year = (year - 1).ToString()
                    };
                }

                // Year hiện tại
                if (dataItemYear == null)
                {
                    dataItemYear = new ReportDetailtForTotalMoneyType()
                    {
                        PartnerCode = item.PartnerCode,
                        PartnerName = item.PartnerName,
                        Year = year.ToString()
                    };
                }

                double sumVND = dataItemYear.VND - dataItemLastYear.VND;
                double sumUSD = dataItemYear.USD - dataItemLastYear.USD;
                double sumEUR = dataItemYear.EUR - dataItemLastYear.EUR;
                double sumCAD = dataItemYear.CAD - dataItemLastYear.CAD;
                double sumAUD = dataItemYear.AUD - dataItemLastYear.AUD;
                double sumGBP = dataItemYear.GBP - dataItemLastYear.GBP;

                table.Rows.Add(
                    count++, item.PartnerName
                    , dataItemYear.VND, dataItemLastYear.VND, Math.Round(sumVND, 2, MidpointRounding.ToEven)
                    , dataItemYear.USD, dataItemLastYear.USD, Math.Round(sumUSD, 2, MidpointRounding.ToEven)
                    , dataItemYear.EUR, dataItemLastYear.EUR, Math.Round(sumEUR, 2, MidpointRounding.ToEven)
                    , dataItemYear.CAD, dataItemLastYear.CAD, Math.Round(sumCAD, 2, MidpointRounding.ToEven)
                    , dataItemYear.AUD, dataItemLastYear.AUD, Math.Round(sumAUD, 2, MidpointRounding.ToEven)
                    , dataItemYear.GBP, dataItemLastYear.GBP, Math.Round(sumGBP, 2, MidpointRounding.ToEven)
                );
            }

            DataRow row = table.NewRow();
            row["STT"] = "";
            row["PartnerName"] = "Tổng";

            row["VND1"] = table.Compute("Sum(VND1)", "");
            row["VND2"] = table.Compute("Sum(VND2)", "");
            row["TDS1"] = table.Compute("Sum(TDS1)", "");

            row["USD1"] = table.Compute("Sum(USD1)", "");
            row["USD2"] = table.Compute("Sum(USD2)", "");
            row["TDS2"] = table.Compute("Sum(TDS2)", "");

            row["EUR1"] = table.Compute("Sum(EUR1)", "");
            row["EUR2"] = table.Compute("Sum(EUR2)", "");
            row["TDS3"] = table.Compute("Sum(TDS3)", "");

            row["CAD1"] = table.Compute("Sum(CAD1)", "");
            row["CAD2"] = table.Compute("Sum(CAD2)", "");
            row["TDS4"] = table.Compute("Sum(TDS4)", "");

            row["AUD1"] = table.Compute("Sum(AUD1)", "");
            row["AUD2"] = table.Compute("Sum(AUD2)", "");
            row["TDS5"] = table.Compute("Sum(TDS5)", "");

            row["GBP1"] = table.Compute("Sum(GBP1)", "");
            row["GBP2"] = table.Compute("Sum(GBP2)", "");
            row["TDS6"] = table.Compute("Sum(TDS6)", "");

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
        public ActionResult SearchGridReportForGradationConvert([DataSourceRequest]DataSourceRequest request, int gradation, int year, string reportTypeID)
        {
            List<ReportDetailtForTotalMoneyType> listDataGradation = new ReportBL().ReportDetailtPartnerLTGradationCompareForAllConvert(year, gradation, reportTypeID);

            // Khởi tạo datatable
            DataTable table = new DataTable();
            // Tạo các cột cho datatable
            table.Columns.Add("STT", typeof(String));
            table.Columns.Add("PartnerName", typeof(String));

            table.Columns.Add("VND1", typeof(double));
            table.Columns.Add("VND2", typeof(double));
            table.Columns.Add("TDS1", typeof(double));

            table.Columns.Add("USD1", typeof(double));
            table.Columns.Add("USD2", typeof(double));
            table.Columns.Add("TDS2", typeof(double));

            table.Columns.Add("EUR1", typeof(double));
            table.Columns.Add("EUR2", typeof(double));
            table.Columns.Add("TDS3", typeof(double));

            table.Columns.Add("CAD1", typeof(double));
            table.Columns.Add("CAD2", typeof(double));
            table.Columns.Add("TDS4", typeof(double));

            table.Columns.Add("AUD1", typeof(double));
            table.Columns.Add("AUD2", typeof(double));
            table.Columns.Add("TDS5", typeof(double));

            table.Columns.Add("GBP1", typeof(double));
            table.Columns.Add("GBP2", typeof(double));
            table.Columns.Add("TDS6", typeof(double));

            int count = 1;
            List<string> listPartner = new List<string>();
            foreach (ReportDetailtForTotalMoneyType item in listDataGradation)
            {
                // Check trường hợp đã tồn tại đối tác
                if (listPartner.Contains(item.PartnerCode))
                {
                    continue;
                }

                listPartner.Add(item.PartnerCode);

                // Cùng kì
                ReportDetailtForTotalMoneyType dataItemLastYear = listDataGradation.Find(x => x.PartnerCode == item.PartnerCode && x.Year == (year - 1).ToString());
                ReportDetailtForTotalMoneyType dataItemYear = listDataGradation.Find(x => x.PartnerCode == item.PartnerCode && x.Year == year.ToString());

                // Last year
                if (dataItemLastYear == null)
                {
                    dataItemLastYear = new ReportDetailtForTotalMoneyType()
                    {
                        PartnerCode = item.PartnerCode,
                        PartnerName = item.PartnerName,
                        Year = (year - 1).ToString()
                    };
                }

                // Year hiện tại
                if (dataItemYear == null)
                {
                    dataItemYear = new ReportDetailtForTotalMoneyType()
                    {
                        PartnerCode = item.PartnerCode,
                        PartnerName = item.PartnerName,
                        Year = year.ToString()
                    };
                }

                double sumVND = dataItemYear.VND - dataItemLastYear.VND;
                double sumUSD = dataItemYear.USD - dataItemLastYear.USD;
                double sumEUR = dataItemYear.EUR - dataItemLastYear.EUR;
                double sumCAD = dataItemYear.CAD - dataItemLastYear.CAD;
                double sumAUD = dataItemYear.AUD - dataItemLastYear.AUD;
                double sumGBP = dataItemYear.GBP - dataItemLastYear.GBP;

                table.Rows.Add(
                    count++, item.PartnerName
                    , dataItemYear.VND, dataItemLastYear.VND, Math.Round(sumVND, 2, MidpointRounding.ToEven)
                    , dataItemYear.USD, dataItemLastYear.USD, Math.Round(sumUSD, 2, MidpointRounding.ToEven)
                    , dataItemYear.EUR, dataItemLastYear.EUR, Math.Round(sumEUR, 2, MidpointRounding.ToEven)
                    , dataItemYear.CAD, dataItemLastYear.CAD, Math.Round(sumCAD, 2, MidpointRounding.ToEven)
                    , dataItemYear.AUD, dataItemLastYear.AUD, Math.Round(sumAUD, 2, MidpointRounding.ToEven)
                    , dataItemYear.GBP, dataItemLastYear.GBP, Math.Round(sumGBP, 2, MidpointRounding.ToEven)
                );
            }

            DataRow row = table.NewRow();
            row["STT"] = "";
            row["PartnerName"] = "Tổng";

            row["VND1"] = table.Compute("Sum(VND1)", "");
            row["VND2"] = table.Compute("Sum(VND2)", "");
            row["TDS1"] = table.Compute("Sum(TDS1)", "");

            row["USD1"] = table.Compute("Sum(USD1)", "");
            row["USD2"] = table.Compute("Sum(USD2)", "");
            row["TDS2"] = table.Compute("Sum(TDS2)", "");

            row["EUR1"] = table.Compute("Sum(EUR1)", "");
            row["EUR2"] = table.Compute("Sum(EUR2)", "");
            row["TDS3"] = table.Compute("Sum(TDS3)", "");

            row["CAD1"] = table.Compute("Sum(CAD1)", "");
            row["CAD2"] = table.Compute("Sum(CAD2)", "");
            row["TDS4"] = table.Compute("Sum(TDS4)", "");

            row["AUD1"] = table.Compute("Sum(AUD1)", "");
            row["AUD2"] = table.Compute("Sum(AUD2)", "");
            row["TDS5"] = table.Compute("Sum(TDS5)", "");

            row["GBP1"] = table.Compute("Sum(GBP1)", "");
            row["GBP2"] = table.Compute("Sum(GBP2)", "");
            row["TDS6"] = table.Compute("Sum(TDS6)", "");

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
        public ActionResult SearchGridReportForGradationForOne([DataSourceRequest]DataSourceRequest request, int gradation, int year, string reportTypeID, string partnerID)
        {
            List<ReportDetailtForTotalMoneyType> listDataGradation = new ReportBL().ReportDetailtPartnerLTGradationCompareForOne(year, gradation, reportTypeID, partnerID);
            List<ReportDetailtForTotalMoneyType> listDataGradationConvert = new ReportBL().ReportDetailtPartnerLTGradationCompareForOneConvert(year, gradation, reportTypeID, partnerID);
            List<ReportDetailtForTotalMoneyType> listDataTotal = new List<ReportDetailtForTotalMoneyType>();

            string text = string.Empty;

            switch (gradation)
            {
                case 1:
                    text = "3 tháng";
                    break;
                case 2:
                    text = "6 tháng";
                    break;
                case 3:
                    text = "9 tháng";
                    break;
                default:
                    text = "12 tháng";
                    break;
            }


            foreach (ReportDetailtForTotalMoneyType item in listDataGradation)
            {
                listDataTotal.Add(
                    new ReportDetailtForTotalMoneyType()
                    {
                        PartnerCode = item.PartnerCode,
                        PartnerName = item.PartnerName,
                        VND = item.VND,
                        USD = item.USD,
                        EUR = item.EUR,
                        CAD = item.CAD,
                        AUD = item.AUD,
                        GBP = item.GBP,
                        typeID = 0,
                        Year = item.Year
                    }
                );
            }

            foreach (ReportDetailtForTotalMoneyType item in listDataGradationConvert)
            {
                listDataTotal.Add(
                    new ReportDetailtForTotalMoneyType()
                    {
                        PartnerCode = item.PartnerCode,
                        PartnerName = item.PartnerName,
                        VND = item.VND,
                        USD = item.USD,
                        EUR = item.EUR,
                        CAD = item.CAD,
                        AUD = item.AUD,
                        GBP = item.GBP,
                        typeID = 1,
                        Year = item.Year
                    }
                );
            }
            
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

            List<int> listPartner = new List<int>();
            string value = string.Empty;

            foreach (ReportDetailtForTotalMoneyType item in listDataTotal)
            {
                value = "Nguyên tệ";

                if(item.typeID.Equals(1))
                {
                    value = "Quy USD";
                }
                if(listPartner.Contains(item.typeID))
                {
                    continue;
                }

                listPartner.Add(item.typeID);

                // Nguyên tệ
                ReportDetailtForTotalMoneyType dataItemYear = listDataTotal.Find(x => x.Year == year.ToString() && x.typeID == item.typeID);
                ReportDetailtForTotalMoneyType dataItemLastYear = listDataTotal.Find(x => x.Year == (year - 1).ToString() && x.typeID == item.typeID);

                if (dataItemYear == null)
                {
                    dataItemYear = new ReportDetailtForTotalMoneyType()
                    {
                        Year = year.ToString()
                    };
                }

                if (dataItemLastYear == null)
                {
                    dataItemLastYear = new ReportDetailtForTotalMoneyType()
                    {
                        Year = (year - 1).ToString()
                    };
                }

                table.Rows.Add(
                     value
                    , dataItemYear.VND, dataItemLastYear.VND, dataItemYear.VND - dataItemLastYear.VND
                    , dataItemYear.USD, dataItemLastYear.USD, dataItemYear.USD - dataItemLastYear.USD
                    , dataItemYear.EUR, dataItemLastYear.EUR, dataItemYear.EUR - dataItemLastYear.EUR
                    , dataItemYear.CAD, dataItemLastYear.CAD, dataItemYear.CAD - dataItemLastYear.CAD
                    , dataItemYear.AUD, dataItemLastYear.AUD, dataItemYear.AUD - dataItemLastYear.AUD
                    , dataItemYear.GBP, dataItemLastYear.GBP, dataItemYear.GBP - dataItemLastYear.GBP
                );
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
        public ActionResult SearchColumnsChartGradationCompareForOne([DataSourceRequest]DataSourceRequest request, int gradation, int year, string reportTypeID, string partnerID)
        {
            List<ReportDetailtForTotalMoneyType> listDataGradationConvert = new ReportBL().ReportDetailtPartnerLTGradationCompareForOneConvert(year, gradation, reportTypeID, partnerID);

            string text = string.Empty;
            switch(gradation)
            {
                case 1:
                    text = "3";
                    break;
                case 2:
                    text = "6";
                    break;

                case 3:
                    text = "9";
                    break;

                default:
                    text = "12";
                    break;
            }
            // Số record của mảng
            int countArray = 6;
            GradationCompare[] arrayGradation = new GradationCompare[countArray * listDataGradationConvert.Count];
            int count = 0;
            foreach (ReportDetailtForTotalMoneyType item in listDataGradationConvert)
            {
                // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                arrayGradation[count] = new GradationCompare()
                {
                    NameGradationCompare = string.Format("Lũy kế {0} tháng năm {1}", text, item.Year),
                    amount = item.VND,
                    NameType = "VND"
                };

                count++;

                arrayGradation[count] = new GradationCompare()
                {
                    NameGradationCompare = string.Format("Lũy kế {0} tháng năm {1}", text, item.Year),
                    amount = item.USD,
                    NameType = "USD"
                };

                count++;

                arrayGradation[count] = new GradationCompare()
                {
                    NameGradationCompare = string.Format("Lũy kế {0} tháng năm {1}", text, item.Year),
                    amount = item.EUR,
                    NameType = "EUR"
                };

                count++;

                arrayGradation[count] = new GradationCompare()
                {
                    NameGradationCompare = string.Format("Lũy kế {0} tháng năm {1}", text, item.Year),
                    amount = item.CAD,
                    NameType = "CAD"
                };

                count++;

                arrayGradation[count] = new GradationCompare()
                {
                    NameGradationCompare = string.Format("Lũy kế {0} tháng năm {1}", text, item.Year),
                    amount = item.AUD,
                    NameType = "AUD"
                };

                count++;

                arrayGradation[count] = new GradationCompare()
                {
                    NameGradationCompare = string.Format("Lũy kế {0} tháng năm {1}", text, item.Year),
                    amount = item.GBP,
                    NameType = "GBP"
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
        /// search data cho biểu đồ của năm hiện tại
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult SearchGradationComparePieForYear([DataSourceRequest]DataSourceRequest request, string gradation, int year, string reportTypeID, string partnerID)
        {
            List<ReportDetailtForTotalMoneyType> listDataGradationConvert = new ReportBL().ReportDetailtPartnerLTGradationCompareForOneConvert(year, int.Parse(gradation), reportTypeID, partnerID);
            ReportDetailtForTotalMoneyType dataGradationPercent = null;

            foreach (ReportDetailtForTotalMoneyType item in listDataGradationConvert)
            {
                if(item.Year == year.ToString())
                {
                    item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;

                    dataGradationPercent = new ReportDetailtForTotalMoneyType()
                    {
                        VND = item.TongDS == 0 ? 0 : Math.Round((item.VND / item.TongDS) * 100, 2, MidpointRounding.ToEven),
                        USD = item.TongDS == 0 ? 0 : Math.Round((item.USD / item.TongDS) * 100, 2, MidpointRounding.ToEven),
                        EUR = item.TongDS == 0 ? 0 : Math.Round((item.EUR / item.TongDS) * 100, 2, MidpointRounding.ToEven),
                        CAD = item.TongDS == 0 ? 0 : Math.Round((item.CAD / item.TongDS) * 100, 2, MidpointRounding.ToEven),
                        AUD = item.TongDS == 0 ? 0 : Math.Round((item.AUD / item.TongDS) * 100, 2, MidpointRounding.ToEven),
                        GBP = item.TongDS == 0 ? 0 : Math.Round((item.GBP / item.TongDS) * 100, 2, MidpointRounding.ToEven),
                    };
                }
            }

            // # dòng record
            GradationChartPie[] arrayGradation = new GradationChartPie[1];
            arrayGradation[0] = new GradationChartPie()
            {
                category = "1",
                value = 0,
                color = "#9de219"

            };

            int count = 0;
            if(dataGradationPercent != null)
            {
                // tạo mảng gồm 8 object
                arrayGradation = new GradationChartPie[6];

                // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                arrayGradation[count] = new GradationChartPie()
                {
                    category = "VND",
                    value = dataGradationPercent.VND,
                    color = "#FFBF00"
                };

                count++;

                arrayGradation[count] = new GradationChartPie()
                {
                    category = "USD",
                    value = dataGradationPercent.USD,
                    color = "#40FF00"
                };

                count++;

                arrayGradation[count] = new GradationChartPie()
                {
                    category = "EUR",
                    value = dataGradationPercent.EUR,
                    color = "#2ECCFA"
                };

                count++;

                arrayGradation[count] = new GradationChartPie()
                {
                    category = "CAD",
                    value = dataGradationPercent.CAD,
                    color = "#9A2EFE"
                };

                count++;

                arrayGradation[count] = new GradationChartPie()
                {
                    category = "AUD",
                    value = dataGradationPercent.AUD,
                    color = "#FE2EF7"
                };

                count++;

                arrayGradation[count] = new GradationChartPie()
                {
                    category = "GBP",
                    value = dataGradationPercent.GBP,
                    color = "#0000FF"
                };

                count++;
            }

           return Json(arrayGradation);
        }

        /// <summary>
        /// search data cho biểu đồ của năm hiện tại
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult SearchGradationComparePieForLastYear([DataSourceRequest]DataSourceRequest request, string gradation, int year, string reportTypeID, string partnerID)
        {
            List<ReportDetailtForTotalMoneyType> listDataGradationConvert = new ReportBL().ReportDetailtPartnerLTGradationCompareForOneConvert(year, int.Parse(gradation), reportTypeID, partnerID);
            ReportDetailtForTotalMoneyType dataGradationPercent = null;

            foreach (ReportDetailtForTotalMoneyType item in listDataGradationConvert)
            {
                if (item.Year == (year - 1).ToString())
                {
                    item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;

                    dataGradationPercent = new ReportDetailtForTotalMoneyType()
                    {
                        VND = item.TongDS == 0 ? 0 : Math.Round((item.VND / item.TongDS) * 100, 2, MidpointRounding.ToEven),
                        USD = item.TongDS == 0 ? 0 : Math.Round((item.USD / item.TongDS) * 100, 2, MidpointRounding.ToEven),
                        EUR = item.TongDS == 0 ? 0 : Math.Round((item.EUR / item.TongDS) * 100, 2, MidpointRounding.ToEven),
                        CAD = item.TongDS == 0 ? 0 : Math.Round((item.CAD / item.TongDS) * 100, 2, MidpointRounding.ToEven),
                        AUD = item.TongDS == 0 ? 0 : Math.Round((item.AUD / item.TongDS) * 100, 2, MidpointRounding.ToEven),
                        GBP = item.TongDS == 0 ? 0 : Math.Round((item.GBP / item.TongDS) * 100, 2, MidpointRounding.ToEven),
                    };
                }
            }

            // # dòng record
            GradationChartPie[] arrayGradation = new GradationChartPie[1];
            arrayGradation[0] = new GradationChartPie()
            {
                category = "1",
                value = 0,
                color = "#9de219"

            };

            int count = 0;
            if (dataGradationPercent != null)
            {
                // tạo mảng gồm 8 object
                arrayGradation = new GradationChartPie[6];

                // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                arrayGradation[count] = new GradationChartPie()
                {
                    category = "VND",
                    value = dataGradationPercent.VND,
                    color = "#FFBF00"
                };

                count++;

                arrayGradation[count] = new GradationChartPie()
                {
                    category = "USD",
                    value = dataGradationPercent.USD,
                    color = "#40FF00"
                };

                count++;

                arrayGradation[count] = new GradationChartPie()
                {
                    category = "EUR",
                    value = dataGradationPercent.EUR,
                    color = "#2ECCFA"
                };

                count++;

                arrayGradation[count] = new GradationChartPie()
                {
                    category = "CAD",
                    value = dataGradationPercent.CAD,
                    color = "#9A2EFE"
                };

                count++;

                arrayGradation[count] = new GradationChartPie()
                {
                    category = "AUD",
                    value = dataGradationPercent.AUD,
                    color = "#FE2EF7"
                };

                count++;

                arrayGradation[count] = new GradationChartPie()
                {
                    category = "GBP",
                    value = dataGradationPercent.GBP,
                    color = "#0000FF"
                };

                count++;
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
        public ActionResult SearchGridReportForGradationForOnePercent([DataSourceRequest]DataSourceRequest request, int gradation, int year, string reportTypeID, string partnerID)
        {
            List<ReportDetailtForTotalMoneyType> listDataGradationConvert = new ReportBL().ReportDetailtPartnerLTGradationCompareForOneConvert(year, gradation, reportTypeID, partnerID);
            List<ReportDetailtForTotalMoneyType> listdataGradationPercent = new List<ReportDetailtForTotalMoneyType>();

            string text = string.Empty;

            switch (gradation)
            {
                case 1:
                    text = "3 tháng";
                    break;
                case 2:
                    text = "6 tháng";
                    break;
                case 3:
                    text = "9 tháng";
                    break;
                default:
                    text = "12 tháng";
                    break;
            }
            

            foreach(ReportDetailtForTotalMoneyType item in listDataGradationConvert)
            {
                item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;

                listdataGradationPercent.Add(
                    new ReportDetailtForTotalMoneyType()
                    {
                        PartnerCode = item.PartnerCode,
                        PartnerName = item.PartnerName,
                        VND = item.TongDS == 0 ? 0 : Math.Round((item.VND / item.TongDS) * 100, 2, MidpointRounding.ToEven),
                        USD = item.TongDS == 0 ? 0 : Math.Round((item.USD / item.TongDS) * 100, 2, MidpointRounding.ToEven),
                        EUR = item.TongDS == 0 ? 0 : Math.Round((item.EUR / item.TongDS) * 100, 2, MidpointRounding.ToEven),
                        CAD = item.TongDS == 0 ? 0 : Math.Round((item.CAD / item.TongDS) * 100, 2, MidpointRounding.ToEven),
                        AUD = item.TongDS == 0 ? 0 : Math.Round((item.AUD / item.TongDS) * 100, 2, MidpointRounding.ToEven),
                        GBP = item.TongDS == 0 ? 0 : Math.Round((item.GBP / item.TongDS) * 100, 2, MidpointRounding.ToEven),
                        Year = item.Year
                    }

                );
            }

            // Khởi tạo datatable
            DataTable table = new DataTable();
            // Tạo các cột cho datatable
            table.Columns.Add("PartnerName", typeof(String));

            table.Columns.Add("TDS1", typeof(double));
            table.Columns.Add("TDS2", typeof(double));
            table.Columns.Add("TDS3", typeof(double));

            List<int> listPartner = new List<int>();
            string value = string.Empty;

            string[] listTypeMoney = { "VND", "USD", "EUR", "CAD", "AUD", "GBP" };

            if(listdataGradationPercent.Count > 0)
            {
                ReportDetailtForTotalMoneyType dataItemYear = listdataGradationPercent.Find(x => x.Year == year.ToString());
                ReportDetailtForTotalMoneyType dataItemLastYear = listdataGradationPercent.Find(x => x.Year == (year - 1).ToString());

                foreach (string item in listTypeMoney)
                {
                    var propertyInfoYear = dataItemYear.GetType().GetProperty(item);
                    var valueDataYear = propertyInfoYear.GetValue(dataItemYear, null);

                    var propertyInfoLastYear = dataItemLastYear.GetType().GetProperty(item);
                    var valueDataLastYear = propertyInfoYear.GetValue(dataItemLastYear, null);

                    double sum = Math.Round(Convert.ToDouble(valueDataYear) - Convert.ToDouble(valueDataLastYear), 2, MidpointRounding.ToEven);

                    table.Rows.Add(
                        item
                        , valueDataYear, valueDataLastYear, sum
                    );

                }
            }

            DataRow row = table.NewRow();

            row["PartnerName"] = "Tổng";

            row["TDS1"] = 100;
            row["TDS2"] = 100;
            row["TDS3"] = table.Compute("Sum(TDS3)", "");

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
        public ActionResult SearchReportDetailtCompareMonthForAll([DataSourceRequest]DataSourceRequest request, int month, int year, string reportTypeID)
        {
            List<ReportDetailtForTotalMoneyType> listDataCompareMonth = new ReportBL().ReportDetailtPartnerLTCompareMonthForAll(year, month, reportTypeID);

            DataTable table = new DataTable();
            // Khởi tạo datatable
            table = new DataTable();
            // Tạo các cột cho datatable
            // tháng hiện tại
            table.Columns.Add("STT", typeof(String));

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

            List<string> listPartner = new List<string>();
            int count = 1;

            foreach (ReportDetailtForTotalMoneyType item in listDataCompareMonth)
            {
                // Check tồn tại của đối tác
                if (listPartner.Contains(item.PartnerCode))
                {
                    continue;
                }
                listPartner.Add(item.PartnerCode);

                // Cùng kì
                ReportDetailtForTotalMoneyType dataItemLastYear = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == month.ToString() && x.Year == (year - 1).ToString());
                ReportDetailtForTotalMoneyType dataItemYear = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == month.ToString() && x.Year == year.ToString());
                ReportDetailtForTotalMoneyType dataItemLastMonth = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == (month - 1).ToString() && x.Year == year.ToString());
                // Trường hợp tháng 1
                if (month == 1)
                {
                    dataItemLastMonth = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == "12" && x.Year == (year - 1).ToString());
                }
                // Cung kì
                if (dataItemLastYear == null)
                {
                    dataItemLastYear = new ReportDetailtForTotalMoneyType()
                    {
                        PartnerCode = item.PartnerCode,
                        PartnerName = item.PartnerName,
                        Month = month.ToString(),
                        Year = (year - 1).ToString()
                    };
                }

                // Tháng hiện tại
                if (dataItemYear == null)
                {
                    dataItemYear = new ReportDetailtForTotalMoneyType()
                    {
                        PartnerCode = item.PartnerCode,
                        PartnerName = item.PartnerName,
                        Month = month.ToString(),
                        Year = year.ToString()
                    };
                }

                // Tháng trước
                if (dataItemLastMonth == null)
                {
                    // Trường hợp tháng 1
                    if (month == 1)
                    {
                        dataItemLastMonth = new ReportDetailtForTotalMoneyType()
                        {
                            PartnerCode = item.PartnerCode,
                            PartnerName = item.PartnerName,
                            Month = "12",
                            Year = (year - 1).ToString()
                        };
                    }
                    else
                    {
                        dataItemLastMonth = new ReportDetailtForTotalMoneyType()
                        {
                            PartnerCode = item.PartnerCode,
                            PartnerName = item.PartnerName,
                            Month = (month - 1).ToString(),
                            Year = year.ToString()
                        };
                    }
                    
                }

                if (dataItemLastYear != null && dataItemYear != null && dataItemLastMonth != null)
                {
                    // add item vào table
                    table.Rows.Add(count++, item.PartnerName
                        , dataItemYear.VND, dataItemLastMonth.VND, dataItemLastYear.VND
                        , dataItemYear.USD, dataItemLastMonth.USD, dataItemLastYear.USD
                        , dataItemYear.EUR, dataItemLastMonth.EUR, dataItemLastYear.EUR
                        , dataItemYear.CAD, dataItemLastMonth.CAD, dataItemLastYear.CAD
                        , dataItemYear.AUD, dataItemLastMonth.AUD, dataItemLastYear.AUD
                        , dataItemYear.GBP, dataItemLastMonth.GBP, dataItemLastYear.GBP
                    );
                }
            }

            // Add dòng tổng
            DataRow row = table.NewRow();
            row["STT"] = "";
            row["PartnerName"] = "Tổng";
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
            table.Rows.Add(row);

            return Json(table.ToDataSourceResult(request), JsonRequestBehavior.AllowGet);
        }


        /// <summary>
        /// Get data cho báo cáo chi tiết so sánh theo tháng - Quy USD
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult SearchReportDetailtCompareMonthForAllConvert([DataSourceRequest]DataSourceRequest request, int month, int year, string reportTypeID)
        {
            List<ReportDetailtForTotalMoneyType> listDataCompareMonth = new ReportBL().ReportDetailtPartnerLTCompareMonthForAllConvert(year, month, reportTypeID);

            DataTable table = new DataTable();
            // Khởi tạo datatable
            table = new DataTable();
            // Tạo các cột cho datatable
            // tháng hiện tại
            table.Columns.Add("STT", typeof(String));

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

            List<string> listPartner = new List<string>();
            int count = 1;

            foreach (ReportDetailtForTotalMoneyType item in listDataCompareMonth)
            {
                // Check tồn tại của đối tác
                if (listPartner.Contains(item.PartnerCode))
                {
                    continue;
                }
                listPartner.Add(item.PartnerCode);

                // Cùng kì
                ReportDetailtForTotalMoneyType dataItemLastYear = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == month.ToString() && x.Year == (year - 1).ToString());
                ReportDetailtForTotalMoneyType dataItemYear = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == month.ToString() && x.Year == year.ToString());
                ReportDetailtForTotalMoneyType dataItemLastMonth = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == (month - 1).ToString() && x.Year == year.ToString());
                // Trường hợp tháng 1
                if (month == 1)
                {
                    dataItemLastMonth = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == "12" && x.Year == (year - 1).ToString());
                }
                // Cung kì
                if (dataItemLastYear == null)
                {
                    dataItemLastYear = new ReportDetailtForTotalMoneyType()
                    {
                        PartnerCode = item.PartnerCode,
                        PartnerName = item.PartnerName,
                        Month = month.ToString(),
                        Year = (year - 1).ToString()
                    };
                }

                // Tháng hiện tại
                if (dataItemYear == null)
                {
                    dataItemYear = new ReportDetailtForTotalMoneyType()
                    {
                        PartnerCode = item.PartnerCode,
                        PartnerName = item.PartnerName,
                        Month = month.ToString(),
                        Year = year.ToString()
                    };
                }

                // Tháng trước
                if (dataItemLastMonth == null)
                {
                    // Trường hợp tháng 1
                    if (month == 1)
                    {
                        dataItemLastMonth = new ReportDetailtForTotalMoneyType()
                        {
                            PartnerCode = item.PartnerCode,
                            PartnerName = item.PartnerName,
                            Month = "12",
                            Year = (year - 1).ToString()
                        };
                    }
                    else
                    {
                        dataItemLastMonth = new ReportDetailtForTotalMoneyType()
                        {
                            PartnerCode = item.PartnerCode,
                            PartnerName = item.PartnerName,
                            Month = (month - 1).ToString(),
                            Year = year.ToString()
                        };
                    }

                }

                if (dataItemLastYear != null && dataItemYear != null && dataItemLastMonth != null)
                {
                    // add item vào table
                    table.Rows.Add(count++, item.PartnerName
                        , dataItemYear.VND, dataItemLastMonth.VND, dataItemLastYear.VND
                        , dataItemYear.USD, dataItemLastMonth.USD, dataItemLastYear.USD
                        , dataItemYear.EUR, dataItemLastMonth.EUR, dataItemLastYear.EUR
                        , dataItemYear.CAD, dataItemLastMonth.CAD, dataItemLastYear.CAD
                        , dataItemYear.AUD, dataItemLastMonth.AUD, dataItemLastYear.AUD
                        , dataItemYear.GBP, dataItemLastMonth.GBP, dataItemLastYear.GBP
                    );
                }
            }

            // Add dòng tổng
            DataRow row = table.NewRow();
            row["STT"] = "";
            row["PartnerName"] = "Tổng";
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
            table.Rows.Add(row);

            return Json(table.ToDataSourceResult(request), JsonRequestBehavior.AllowGet);
        }


        /// <summary>
        /// Get data cho báo cáo chi tiết so sánh theo tháng - Nguyên tệ
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult SearchReportDetailtCompareMonthForAllCompare([DataSourceRequest]DataSourceRequest request, int month, int year, string reportTypeID)
        {
            List<ReportDetailtForTotalMoneyType> listDataCompareMonth = new ReportBL().ReportDetailtPartnerLTCompareMonthForAll(year, month, reportTypeID);

            DataTable table = new DataTable();
            // Khởi tạo datatable
            table = new DataTable();
            // Tạo các cột cho datatable
            // tháng hiện tại
            table.Columns.Add("STT", typeof(String));

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

            List<string> listPartner = new List<string>();
            int count = 1;

            foreach (ReportDetailtForTotalMoneyType item in listDataCompareMonth)
            {
                // Check tồn tại của đối tác
                if (listPartner.Contains(item.PartnerCode))
                {
                    continue;
                }
                listPartner.Add(item.PartnerCode);

                // Cùng kì
                ReportDetailtForTotalMoneyType dataItemLastYear = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == month.ToString() && x.Year == (year - 1).ToString());
                ReportDetailtForTotalMoneyType dataItemYear = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == month.ToString() && x.Year == year.ToString());
                ReportDetailtForTotalMoneyType dataItemLastMonth = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == (month - 1).ToString() && x.Year == year.ToString());

                // Cung kì
                if (dataItemLastYear == null)
                {
                    dataItemLastYear = new ReportDetailtForTotalMoneyType()
                    {
                        PartnerCode = item.PartnerCode,
                        PartnerName = item.PartnerName,
                        Month = month.ToString(),
                        Year = (year - 1).ToString()
                    };
                }

                // Tháng hiện tại
                if (dataItemYear == null)
                {
                    dataItemYear = new ReportDetailtForTotalMoneyType()
                    {
                        PartnerCode = item.PartnerCode,
                        PartnerName = item.PartnerName,
                        Month = month.ToString(),
                        Year = year.ToString()
                    };
                }

                // Tháng trước
                if (dataItemLastMonth == null)
                {
                    dataItemLastMonth = new ReportDetailtForTotalMoneyType()
                    {
                        PartnerCode = item.PartnerCode,
                        PartnerName = item.PartnerName,
                        Month = (month - 1).ToString(),
                        Year = year.ToString()
                    };
                }

                if (dataItemLastYear != null && dataItemYear != null && dataItemLastMonth != null)
                {
                    double sumVNDYear = Math.Round(dataItemYear.VND - dataItemLastMonth.VND, 2, MidpointRounding.ToEven);
                    double sumUSDYear = Math.Round(dataItemYear.USD - dataItemLastMonth.USD, 2, MidpointRounding.ToEven);
                    double sumEURYear = Math.Round(dataItemYear.EUR - dataItemLastMonth.EUR, 2, MidpointRounding.ToEven);
                    double sumCADYear = Math.Round(dataItemYear.CAD - dataItemLastMonth.CAD, 2, MidpointRounding.ToEven);
                    double sumAUDYear = Math.Round(dataItemYear.AUD - dataItemLastMonth.AUD, 2, MidpointRounding.ToEven);
                    double sumGBPYear = Math.Round(dataItemYear.GBP - dataItemLastMonth.GBP, 2, MidpointRounding.ToEven);

                    double sumVNDLastYear = Math.Round(dataItemYear.VND - dataItemLastYear.VND, 2, MidpointRounding.ToEven);
                    double sumUSDLastYear = Math.Round(dataItemYear.USD - dataItemLastYear.USD, 2, MidpointRounding.ToEven);
                    double sumEURLastYear = Math.Round(dataItemYear.EUR - dataItemLastYear.EUR, 2, MidpointRounding.ToEven);
                    double sumCADLastYear = Math.Round(dataItemYear.CAD - dataItemLastYear.CAD, 2, MidpointRounding.ToEven);
                    double sumAUDLastYear = Math.Round(dataItemYear.AUD - dataItemLastYear.AUD, 2, MidpointRounding.ToEven);
                    double sumGBPLastYear = Math.Round(dataItemYear.GBP - dataItemLastYear.GBP, 2, MidpointRounding.ToEven);

                    // add item vào table
                    table.Rows.Add(count++, item.PartnerName
                        , sumVNDYear, sumVNDLastYear
                        , sumUSDYear, sumUSDLastYear
                        , sumEURYear, sumEURLastYear
                        , sumCADYear, sumCADLastYear
                        , sumAUDYear, sumAUDLastYear
                        , sumGBPYear, sumGBPLastYear
                    );
                }
            }

            // Add dòng tổng
            DataRow row = table.NewRow();
            row["STT"] = "";
            row["PartnerName"] = "Tổng";
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
            table.Rows.Add(row);

            return Json(table.ToDataSourceResult(request), JsonRequestBehavior.AllowGet);
        }


        /// <summary>
        /// Get data cho báo cáo chi tiết so sánh theo tháng - Quy USD
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult SearchReportDetailtCompareMonthForAllCompareConvert([DataSourceRequest]DataSourceRequest request, int month, int year, string reportTypeID)
        {
            List<ReportDetailtForTotalMoneyType> listDataCompareMonth = new ReportBL().ReportDetailtPartnerLTCompareMonthForAllConvert(year, month, reportTypeID);

            DataTable table = new DataTable();
            // Khởi tạo datatable
            table = new DataTable();
            // Tạo các cột cho datatable
            // tháng hiện tại
            table.Columns.Add("STT", typeof(String));

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

            List<string> listPartner = new List<string>();
            int count = 1;

            foreach (ReportDetailtForTotalMoneyType item in listDataCompareMonth)
            {
                // Check tồn tại của đối tác
                if (listPartner.Contains(item.PartnerCode))
                {
                    continue;
                }
                listPartner.Add(item.PartnerCode);

                // Cùng kì
                ReportDetailtForTotalMoneyType dataItemLastYear = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == month.ToString() && x.Year == (year - 1).ToString());
                ReportDetailtForTotalMoneyType dataItemYear = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == month.ToString() && x.Year == year.ToString());
                ReportDetailtForTotalMoneyType dataItemLastMonth = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == (month - 1).ToString() && x.Year == year.ToString());

                // Cung kì
                if (dataItemLastYear == null)
                {
                    dataItemLastYear = new ReportDetailtForTotalMoneyType()
                    {
                        PartnerCode = item.PartnerCode,
                        PartnerName = item.PartnerName,
                        Month = month.ToString(),
                        Year = (year - 1).ToString()
                    };
                }

                // Tháng hiện tại
                if (dataItemYear == null)
                {
                    dataItemYear = new ReportDetailtForTotalMoneyType()
                    {
                        PartnerCode = item.PartnerCode,
                        PartnerName = item.PartnerName,
                        Month = month.ToString(),
                        Year = year.ToString()
                    };
                }

                // Tháng trước
                if (dataItemLastMonth == null)
                {
                    dataItemLastMonth = new ReportDetailtForTotalMoneyType()
                    {
                        PartnerCode = item.PartnerCode,
                        PartnerName = item.PartnerName,
                        Month = (month - 1).ToString(),
                        Year = year.ToString()
                    };
                }

                if (dataItemLastYear != null && dataItemYear != null && dataItemLastMonth != null)
                {
                    double sumVNDYear = Math.Round(dataItemYear.VND - dataItemLastMonth.VND, 2, MidpointRounding.ToEven);
                    double sumUSDYear = Math.Round(dataItemYear.USD - dataItemLastMonth.USD, 2, MidpointRounding.ToEven);
                    double sumEURYear = Math.Round(dataItemYear.EUR - dataItemLastMonth.EUR, 2, MidpointRounding.ToEven);
                    double sumCADYear = Math.Round(dataItemYear.CAD - dataItemLastMonth.CAD, 2, MidpointRounding.ToEven);
                    double sumAUDYear = Math.Round(dataItemYear.AUD - dataItemLastMonth.AUD, 2, MidpointRounding.ToEven);
                    double sumGBPYear = Math.Round(dataItemYear.GBP - dataItemLastMonth.GBP, 2, MidpointRounding.ToEven);

                    double sumVNDLastYear = Math.Round(dataItemYear.VND - dataItemLastYear.VND, 2, MidpointRounding.ToEven);
                    double sumUSDLastYear = Math.Round(dataItemYear.USD - dataItemLastYear.USD, 2, MidpointRounding.ToEven);
                    double sumEURLastYear = Math.Round(dataItemYear.EUR - dataItemLastYear.EUR, 2, MidpointRounding.ToEven);
                    double sumCADLastYear = Math.Round(dataItemYear.CAD - dataItemLastYear.CAD, 2, MidpointRounding.ToEven);
                    double sumAUDLastYear = Math.Round(dataItemYear.AUD - dataItemLastYear.AUD, 2, MidpointRounding.ToEven);
                    double sumGBPLastYear = Math.Round(dataItemYear.GBP - dataItemLastYear.GBP, 2, MidpointRounding.ToEven);

                    // add item vào table
                    table.Rows.Add(count++, item.PartnerName
                        , sumVNDYear, sumVNDLastYear
                        , sumUSDYear, sumUSDLastYear
                        , sumEURYear, sumEURLastYear
                        , sumCADYear, sumCADLastYear
                        , sumAUDYear, sumAUDLastYear
                        , sumGBPYear, sumGBPLastYear
                    );
                }
            }

            // Add dòng tổng
            DataRow row = table.NewRow();
            row["STT"] = "";
            row["PartnerName"] = "Tổng";
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
        public ActionResult SearchReportDetailtCompareMonthForOne([DataSourceRequest]DataSourceRequest request, int month, int year, string reportTypeID, string partnerID)
        {
            List<ReportDetailtForTotalMoneyType> listDataCompareMonth = new ReportBL().ReportDetailtPartnerLTCompareMonthForOne(year, month, reportTypeID, partnerID);
            List<ReportDetailtForTotalMoneyType> listDataCompareMonthConvert = new ReportBL().ReportDetailtPartnerLTCompareMonthForOneConvert(year, month, reportTypeID, partnerID);
            List<ReportDetailtForTotalMoneyType> listDataCompareMonthTotal = new List<ReportDetailtForTotalMoneyType>();

            foreach (ReportDetailtForTotalMoneyType item in listDataCompareMonth)
            {
                listDataCompareMonthTotal.Add(
                    new ReportDetailtForTotalMoneyType()
                    {
                        PartnerCode = item.PartnerCode,
                        PartnerName = item.PartnerName,
                        VND = item.VND,
                        USD = item.USD,
                        EUR = item.EUR,
                        CAD = item.CAD,
                        AUD = item.AUD,
                        GBP = item.GBP,
                        typeID = 0,
                        Month = item.Month,
                        Year = item.Year
                    }
                );
            }

            foreach (ReportDetailtForTotalMoneyType item in listDataCompareMonthConvert)
            {
                listDataCompareMonthTotal.Add(
                    new ReportDetailtForTotalMoneyType()
                    {
                        PartnerCode = item.PartnerCode,
                        PartnerName = item.PartnerName,
                        VND = item.VND,
                        USD = item.USD,
                        EUR = item.EUR,
                        CAD = item.CAD,
                        AUD = item.AUD,
                        GBP = item.GBP,
                        TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP,
                        typeID = 1,
                        Month = item.Month,
                        Year = item.Year
                    }
                );
            }

            DataTable table = new DataTable();
            // Khởi tạo datatable
            table = new DataTable();
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

            string value = string.Empty;
            List<int> listPartner = new List<int>();

            foreach (ReportDetailtForTotalMoneyType item in listDataCompareMonthTotal)
            {
                value = "Nguyên tệ";

                if (item.typeID.Equals(1))
                {
                    value = "Quy USD";
                }
                if (listPartner.Contains(item.typeID))
                {
                    continue;
                }

                listPartner.Add(item.typeID);
                
                ReportDetailtForTotalMoneyType dataItemLastYear = listDataCompareMonthTotal.Find(x => x.Month == month.ToString() && x.Year == (year - 1).ToString() && x.typeID == item.typeID);
                ReportDetailtForTotalMoneyType dataItemYear = listDataCompareMonthTotal.Find(x => x.Month == month.ToString() && x.Year == year.ToString() && x.typeID == item.typeID);
                ReportDetailtForTotalMoneyType dataItemLastMonth = listDataCompareMonthTotal.Find(x => x.Month == (month - 1).ToString() && x.Year == year.ToString() && x.typeID == item.typeID);
                // Trường hợp tháng 1
                if (month == 1)
                {
                    dataItemLastMonth = listDataCompareMonthTotal.Find(x => x.Month == "12" && x.Year == (year - 1).ToString() && x.typeID == item.typeID);
                }

                // Tháng hiện tại
                if (dataItemYear == null)
                {
                    dataItemYear = new ReportDetailtForTotalMoneyType()
                    {
                        Month = month.ToString(),
                        Year = year.ToString()
                    };
                }

                // Tháng trước
                if (dataItemLastMonth == null)
                {
                    // Trường hợp tháng 1
                    if (month == 1)
                    {
                        dataItemLastMonth = new ReportDetailtForTotalMoneyType()
                        {
                            Month = "12",
                            Year = (year - 1).ToString()
                        };
                    }
                    else
                    {
                        dataItemLastMonth = new ReportDetailtForTotalMoneyType()
                        {
                            Month = (month - 1).ToString(),
                            Year = year.ToString()
                        };
                    }
                    
                }

                // Cùng kì năm trước
                if (dataItemLastYear == null)
                {
                    dataItemLastYear = new ReportDetailtForTotalMoneyType()
                    {
                        Month = month.ToString(),
                        Year = (year - 1).ToString()
                    };
                }

                table.Rows.Add(
                        value
                    , dataItemYear.VND, dataItemLastMonth.VND, dataItemLastYear.VND
                    , dataItemYear.USD, dataItemLastMonth.USD, dataItemLastYear.USD
                    , dataItemYear.EUR, dataItemLastMonth.EUR, dataItemLastYear.EUR
                    , dataItemYear.CAD, dataItemLastMonth.CAD, dataItemLastYear.CAD
                    , dataItemYear.AUD, dataItemLastMonth.AUD, dataItemLastYear.AUD
                    , dataItemYear.GBP, dataItemLastMonth.GBP, dataItemLastYear.GBP
                );
            }
            
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
        public ActionResult SearchReportDetailtCompareMonthForOneCompare([DataSourceRequest]DataSourceRequest request, int month, int year, string reportTypeID, string partnerID)
        {
            List<ReportDetailtForTotalMoneyType> listDataCompareMonth = new ReportBL().ReportDetailtPartnerLTCompareMonthForOne(year, month, reportTypeID, partnerID);
            List<ReportDetailtForTotalMoneyType> listDataCompareMonthConvert = new ReportBL().ReportDetailtPartnerLTCompareMonthForOneConvert(year, month, reportTypeID, partnerID);
            List<ReportDetailtForTotalMoneyType> listDataCompareMonthTotal = new List<ReportDetailtForTotalMoneyType>();

            foreach (ReportDetailtForTotalMoneyType item in listDataCompareMonth)
            {
                listDataCompareMonthTotal.Add(
                    new ReportDetailtForTotalMoneyType()
                    {
                        PartnerCode = item.PartnerCode,
                        PartnerName = item.PartnerName,
                        VND = item.VND,
                        USD = item.USD,
                        EUR = item.EUR,
                        CAD = item.CAD,
                        AUD = item.AUD,
                        GBP = item.GBP,
                        typeID = 0,
                        Month = item.Month,
                        Year = item.Year
                    }
                );
            }

            foreach (ReportDetailtForTotalMoneyType item in listDataCompareMonthConvert)
            {
                listDataCompareMonthTotal.Add(
                    new ReportDetailtForTotalMoneyType()
                    {
                        PartnerCode = item.PartnerCode,
                        PartnerName = item.PartnerName,
                        VND = item.VND,
                        USD = item.USD,
                        EUR = item.EUR,
                        CAD = item.CAD,
                        AUD = item.AUD,
                        GBP = item.GBP,
                        TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP,
                        typeID = 1,
                        Month = item.Month,
                        Year = item.Year
                    }
                );
            }

            DataTable table = new DataTable();
            // Khởi tạo datatable
            table = new DataTable();
            // Tạo các cột cho datatable
            table.Columns.Add("PartnerName", typeof(String));

            table.Columns.Add("VND1", typeof(double));
            table.Columns.Add("VND2", typeof(double));
            table.Columns.Add("VND3", typeof(double));
            table.Columns.Add("VND4", typeof(double));

            table.Columns.Add("USD1", typeof(double));
            table.Columns.Add("USD2", typeof(double));
            table.Columns.Add("USD3", typeof(double));
            table.Columns.Add("USD4", typeof(double));

            table.Columns.Add("EUR1", typeof(double));
            table.Columns.Add("EUR2", typeof(double));
            table.Columns.Add("EUR3", typeof(double));
            table.Columns.Add("EUR4", typeof(double));

            table.Columns.Add("CAD1", typeof(double));
            table.Columns.Add("CAD2", typeof(double));
            table.Columns.Add("CAD3", typeof(double));
            table.Columns.Add("CAD4", typeof(double));

            table.Columns.Add("AUD1", typeof(double));
            table.Columns.Add("AUD2", typeof(double));
            table.Columns.Add("AUD3", typeof(double));
            table.Columns.Add("AUD4", typeof(double));

            table.Columns.Add("GBP1", typeof(double));
            table.Columns.Add("GBP2", typeof(double));
            table.Columns.Add("GBP3", typeof(double));
            table.Columns.Add("GBP4", typeof(double));

            string value = string.Empty;
            List<int> listPartner = new List<int>();

            foreach (ReportDetailtForTotalMoneyType item in listDataCompareMonthTotal)
            {
                value = "Nguyên tệ";

                if (item.typeID.Equals(1))
                {
                    value = "Quy USD";
                }
                if (listPartner.Contains(item.typeID))
                {
                    continue;
                }

                listPartner.Add(item.typeID);

                ReportDetailtForTotalMoneyType dataItemLastYear = listDataCompareMonthTotal.Find(x => x.Month == month.ToString() && x.Year == (year - 1).ToString() && x.typeID == item.typeID);
                ReportDetailtForTotalMoneyType dataItemYear = listDataCompareMonthTotal.Find(x => x.Month == month.ToString() && x.Year == year.ToString() && x.typeID == item.typeID);
                ReportDetailtForTotalMoneyType dataItemLastMonth = listDataCompareMonthTotal.Find(x => x.Month == (month - 1).ToString() && x.Year == year.ToString() && x.typeID == item.typeID);
                // Trường hợp tháng 1
                if (month == 1)
                {
                    dataItemLastMonth = listDataCompareMonthTotal.Find(x => x.Month == "12" && x.Year == (year - 1).ToString() && x.typeID == item.typeID);
                }

                // Tháng hiện tại
                if (dataItemYear == null)
                {
                    dataItemYear = new ReportDetailtForTotalMoneyType()
                    {
                        Month = month.ToString(),
                        Year = year.ToString()
                    };
                }

                // Tháng trước
                if (dataItemLastMonth == null)
                {
                    // Trường hợp tháng 1
                    if (month == 1)
                    {
                        dataItemLastMonth = new ReportDetailtForTotalMoneyType()
                        {
                            Month = "12",
                            Year = (year - 1).ToString()
                        };
                    }
                    else
                    {
                        dataItemLastMonth = new ReportDetailtForTotalMoneyType()
                        {
                            Month = (month - 1).ToString(),
                            Year = year.ToString()
                        };
                    }
                    
                }

                // Cùng kì năm trước
                if (dataItemLastYear == null)
                {
                    dataItemLastYear = new ReportDetailtForTotalMoneyType()
                    {
                        Month = month.ToString(),
                        Year = (year - 1).ToString()
                    };
                }

                if(dataItemYear != null && dataItemLastMonth != null && dataItemLastYear != null)
                {
                    double sumVNDYear = dataItemYear.VND - dataItemLastMonth.VND;
                    double sumUSDYear = dataItemYear.USD - dataItemLastMonth.USD;
                    double sumEURYear = dataItemYear.EUR - dataItemLastMonth.EUR;
                    double sumCADYear = dataItemYear.CAD - dataItemLastMonth.CAD;
                    double sumAUDYear = dataItemYear.AUD - dataItemLastMonth.AUD;
                    double sumGBPYear = dataItemYear.GBP - dataItemLastMonth.GBP;

                    double sumVNDYearPercent = dataItemLastMonth.VND == 0 ? 0 : Math.Round((sumVNDYear / dataItemLastMonth.VND) * 100, 2, MidpointRounding.ToEven);
                    double sumUSDYearPercent = dataItemLastMonth.USD == 0 ? 0 : Math.Round((sumUSDYear / dataItemLastMonth.USD) * 100, 2, MidpointRounding.ToEven);
                    double sumEURYearPercent = dataItemLastMonth.EUR == 0 ? 0 : Math.Round((sumEURYear / dataItemLastMonth.EUR) * 100, 2, MidpointRounding.ToEven);
                    double sumCADYearPercent = dataItemLastMonth.CAD == 0 ? 0 : Math.Round((sumCADYear / dataItemLastMonth.CAD) * 100, 2, MidpointRounding.ToEven);
                    double sumAUDYearPercent = dataItemLastMonth.AUD == 0 ? 0 : Math.Round((sumAUDYear / dataItemLastMonth.AUD) * 100, 2, MidpointRounding.ToEven);
                    double sumGBPYearPercent = dataItemLastMonth.GBP == 0 ? 0 : Math.Round((sumGBPYear / dataItemLastMonth.GBP) * 100, 2, MidpointRounding.ToEven);

                    double sumVNDLastYear = dataItemYear.VND - dataItemLastYear.VND;
                    double sumUSDLastYear = dataItemYear.USD - dataItemLastYear.USD;
                    double sumEURLastYear = dataItemYear.EUR - dataItemLastYear.EUR;
                    double sumCADLastYear = dataItemYear.CAD - dataItemLastYear.CAD;
                    double sumAUDLastYear = dataItemYear.AUD - dataItemLastYear.AUD;
                    double sumGBPLastYear = dataItemYear.GBP - dataItemLastYear.GBP;

                    double sumVNDLastYearPercent = dataItemLastYear.VND == 0 ? 0 : Math.Round((sumVNDLastYear / dataItemLastYear.VND) * 100, 2, MidpointRounding.ToEven);
                    double sumUSDLastYearPercent = dataItemLastYear.USD == 0 ? 0 : Math.Round((sumUSDLastYear / dataItemLastYear.USD) * 100, 2, MidpointRounding.ToEven);
                    double sumEURLastYearPercent = dataItemLastYear.EUR == 0 ? 0 : Math.Round((sumEURLastYear / dataItemLastYear.EUR) * 100, 2, MidpointRounding.ToEven);
                    double sumCADLastYearPercent = dataItemLastYear.CAD == 0 ? 0 : Math.Round((sumCADLastYear / dataItemLastYear.CAD) * 100, 2, MidpointRounding.ToEven);
                    double sumAUDLastYearPercent = dataItemLastYear.AUD == 0 ? 0 : Math.Round((sumAUDLastYear / dataItemLastYear.AUD) * 100, 2, MidpointRounding.ToEven);
                    double sumGBPLastYearPercent = dataItemLastYear.GBP == 0 ? 0 : Math.Round((sumGBPLastYear / dataItemLastYear.GBP) * 100, 2, MidpointRounding.ToEven);

                    table.Rows.Add(
                        value
                        , sumVNDYear, sumVNDYearPercent, sumVNDLastYear, sumVNDLastYearPercent
                        , sumUSDYear, sumUSDYearPercent, sumUSDLastYear, sumUSDLastYearPercent
                        , sumEURYear, sumEURYearPercent, sumEURLastYear, sumEURLastYearPercent
                        , sumCADYear, sumCADYearPercent, sumCADLastYear, sumCADLastYearPercent
                        , sumAUDYear, sumAUDYearPercent, sumAUDLastYear, sumAUDLastYearPercent
                        , sumGBPYear, sumGBPYearPercent, sumGBPLastYear, sumGBPLastYearPercent
                    );
                }
            }

            return Json(table.ToDataSourceResult(request), JsonRequestBehavior.AllowGet);
        }


        /// <summary>
        /// Get data cho việc vẽ biểu đồ cột cho so sánh theo tháng và cùng kì năm trước
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult SearchDataChartCompareMonth([DataSourceRequest]DataSourceRequest request, int month, int year, string reportTypeID, string partnerID)
        {

            List<ReportDetailtForTotalMoneyType> listDataCompareMonth = new ReportBL().ReportDetailtPartnerLTCompareMonthForOneConvert(year, month, reportTypeID, partnerID);

            // # dòng record
            GradationCompare[] arrayGradation = null;

            if (listDataCompareMonth.Count.Equals(3))
            {
                // tạo mảng gồm 8 object
                arrayGradation = new GradationCompare[18];
                int count = 0;
                foreach (ReportDetailtForTotalMoneyType item in listDataCompareMonth)
                {
                    // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                    arrayGradation[count] = new GradationCompare()
                    {
                        NameGradationCompare = string.Format("Tháng {0}/{1}", item.Month, item.Year),
                        amount = item.VND,
                        NameType = "VND"
                    };

                    count++;

                    arrayGradation[count] = new GradationCompare()
                    {
                        NameGradationCompare = string.Format("Tháng {0}/{1}", item.Month, item.Year),
                        amount = item.USD,
                        NameType = "USD"
                    };

                    count++;
                    arrayGradation[count] = new GradationCompare()
                    {
                        NameGradationCompare = string.Format("Tháng {0}/{1}", item.Month, item.Year),
                        amount = item.EUR,
                        NameType = "EUR"
                    };

                    count++;

                    arrayGradation[count] = new GradationCompare()
                    {
                        NameGradationCompare = string.Format("Tháng {0}/{1}", item.Month, item.Year),
                        amount = item.CAD,
                        NameType = "CAD"
                    };

                    count++;

                    arrayGradation[count] = new GradationCompare()
                    {
                        NameGradationCompare = string.Format("Tháng {0}/{1}", item.Month, item.Year),
                        amount = item.AUD,
                        NameType = "AUD"
                    };

                    count++;

                    arrayGradation[count] = new GradationCompare()
                    {
                        NameGradationCompare = string.Format("Tháng {0}/{1}", item.Month, item.Year),
                        amount = item.GBP,
                        NameType = "GBP"
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
        /// Get data cho báo cáo chi tiết so sánh theo tháng - tháng hiện tại
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult SearchCompareForMonthPieForYear([DataSourceRequest]DataSourceRequest request, int month, int year, string reportTypeID, string partnerID)
        {
            List<ReportDetailtForTotalMoneyType> listDataCompareMonth = new ReportBL().ReportDetailtPartnerLTCompareMonthForOneConvert(year, month, reportTypeID, partnerID);

            List<ReportDetailtForTotalMoneyType> listDataCompareMonthConvert = new List<ReportDetailtForTotalMoneyType>();

            foreach (ReportDetailtForTotalMoneyType item in listDataCompareMonth)
            {
                item.PartnerName = string.Format("Tháng {0}/{1}", item.Month, item.Year);
                item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;

                listDataCompareMonthConvert.Add(
                    new ReportDetailtForTotalMoneyType()
                    {
                        PartnerName = string.Format("Tháng {0}/{1}", item.Month, item.Year),
                        VND = item.TongDS == 0 ? 0 : Math.Round((item.VND / item.TongDS) * 100, 2, MidpointRounding.ToEven),
                        USD = item.TongDS == 0 ? 0 : Math.Round((item.USD / item.TongDS) * 100, 2, MidpointRounding.ToEven),
                        EUR = item.TongDS == 0 ? 0 : Math.Round((item.EUR / item.TongDS) * 100, 2, MidpointRounding.ToEven),
                        CAD = item.TongDS == 0 ? 0 : Math.Round((item.CAD / item.TongDS) * 100, 2, MidpointRounding.ToEven),
                        AUD = item.TongDS == 0 ? 0 : Math.Round((item.AUD / item.TongDS) * 100, 2, MidpointRounding.ToEven),
                        GBP = item.TongDS == 0 ? 0 : Math.Round((item.GBP / item.TongDS) * 100, 2, MidpointRounding.ToEven),
                        TongDS = 100,
                        Month = item.Month,
                        Year = item.Year
                    }
                );
            }

            // # dòng record
            GradationChartPie[] arrayGradation = new GradationChartPie[1];
            arrayGradation[0] = new GradationChartPie()
            {
                category = "1",
                value = 0,
                color = "#9de219"

            };

            int count = 0;
            foreach (ReportDetailtForTotalMoneyType item in listDataCompareMonthConvert)
            {
                if (item.Year == year.ToString() && item.Month == month.ToString())
                {
                    // tạo mảng gồm 8 object
                    arrayGradation = new GradationChartPie[6];

                    // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                    arrayGradation[count] = new GradationChartPie()
                    {
                        category = "VND",
                        value = item.VND,
                        color = "#FFBF00"
                    };

                    count++;

                    arrayGradation[count] = new GradationChartPie()
                    {
                        category = "USD",
                        value = item.USD,
                        color = "#40FF00"
                    };

                    count++;

                    arrayGradation[count] = new GradationChartPie()
                    {
                        category = "EUR",
                        value = item.EUR,
                        color = "#2ECCFA"
                    };

                    count++;

                    arrayGradation[count] = new GradationChartPie()
                    {
                        category = "CAD",
                        value = item.CAD,
                        color = "#9A2EFE"
                    };

                    count++;

                    arrayGradation[count] = new GradationChartPie()
                    {
                        category = "AUD",
                        value = item.AUD,
                        color = "#FE2EF7"
                    };

                    count++;

                    arrayGradation[count] = new GradationChartPie()
                    {
                        category = "GBP",
                        value = item.GBP,
                        color = "#0000FF"
                    };

                    count++;

                }
            }

            return Json(arrayGradation);
        }


        /// <summary>
        /// Get data cho báo cáo chi tiết so sánh theo tháng - tháng hiện tại
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult SearchCompareForMonthPieForLastYear([DataSourceRequest]DataSourceRequest request, int month, int year, string reportTypeID, string partnerID)
        {
            List<ReportDetailtForTotalMoneyType> listDataCompareMonth = new ReportBL().ReportDetailtPartnerLTCompareMonthForOneConvert(year, month, reportTypeID, partnerID);

            List<ReportDetailtForTotalMoneyType> listDataCompareMonthConvert = new List<ReportDetailtForTotalMoneyType>();

            foreach (ReportDetailtForTotalMoneyType item in listDataCompareMonth)
            {
                item.PartnerName = string.Format("Tháng {0}/{1}", item.Month, item.Year);
                item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;

                listDataCompareMonthConvert.Add(
                    new ReportDetailtForTotalMoneyType()
                    {
                        PartnerName = string.Format("Tháng {0}/{1}", item.Month, item.Year),
                        VND = item.TongDS == 0 ? 0 : Math.Round((item.VND / item.TongDS) * 100, 2, MidpointRounding.ToEven),
                        USD = item.TongDS == 0 ? 0 : Math.Round((item.USD / item.TongDS) * 100, 2, MidpointRounding.ToEven),
                        EUR = item.TongDS == 0 ? 0 : Math.Round((item.EUR / item.TongDS) * 100, 2, MidpointRounding.ToEven),
                        CAD = item.TongDS == 0 ? 0 : Math.Round((item.CAD / item.TongDS) * 100, 2, MidpointRounding.ToEven),
                        AUD = item.TongDS == 0 ? 0 : Math.Round((item.AUD / item.TongDS) * 100, 2, MidpointRounding.ToEven),
                        GBP = item.TongDS == 0 ? 0 : Math.Round((item.GBP / item.TongDS) * 100, 2, MidpointRounding.ToEven),
                        TongDS = 100,
                        Month = item.Month,
                        Year = item.Year
                    }
                );
            }

            // # dòng record
            GradationChartPie[] arrayGradation = new GradationChartPie[1];
            arrayGradation[0] = new GradationChartPie()
            {
                category = "1",
                value = 0,
                color = "#9de219"

            };

            int count = 0;
            foreach (ReportDetailtForTotalMoneyType item in listDataCompareMonthConvert)
            {
                if (item.Year == (year - 1).ToString() && item.Month == month.ToString())
                {
                    // tạo mảng gồm 8 object
                    arrayGradation = new GradationChartPie[6];

                    // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                    arrayGradation[count] = new GradationChartPie()
                    {
                        category = "VND",
                        value = item.VND,
                        color = "#FFBF00"
                    };

                    count++;

                    arrayGradation[count] = new GradationChartPie()
                    {
                        category = "USD",
                        value = item.USD,
                        color = "#40FF00"
                    };

                    count++;

                    arrayGradation[count] = new GradationChartPie()
                    {
                        category = "EUR",
                        value = item.EUR,
                        color = "#2ECCFA"
                    };

                    count++;

                    arrayGradation[count] = new GradationChartPie()
                    {
                        category = "CAD",
                        value = item.CAD,
                        color = "#9A2EFE"
                    };

                    count++;

                    arrayGradation[count] = new GradationChartPie()
                    {
                        category = "AUD",
                        value = item.AUD,
                        color = "#FE2EF7"
                    };

                    count++;

                    arrayGradation[count] = new GradationChartPie()
                    {
                        category = "GBP",
                        value = item.GBP,
                        color = "#0000FF"
                    };

                    count++;

                }
            }

            return Json(arrayGradation);
        }


        /// <summary>
        /// Get data cho báo cáo chi tiết so sánh theo tháng - tháng hiện tại
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult SearchCompareForMonthPieForLastMonth([DataSourceRequest]DataSourceRequest request, int month, int year, string reportTypeID, string partnerID)
        {
            List<ReportDetailtForTotalMoneyType> listDataCompareMonth = new ReportBL().ReportDetailtPartnerLTCompareMonthForOneConvert(year, month, reportTypeID, partnerID);

            List<ReportDetailtForTotalMoneyType> listDataCompareMonthConvert = new List<ReportDetailtForTotalMoneyType>();

            foreach (ReportDetailtForTotalMoneyType item in listDataCompareMonth)
            {
                item.PartnerName = string.Format("Tháng {0}/{1}", item.Month, item.Year);
                item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;

                listDataCompareMonthConvert.Add(
                    new ReportDetailtForTotalMoneyType()
                    {
                        PartnerName = string.Format("Tháng {0}/{1}", item.Month, item.Year),
                        VND = item.TongDS == 0 ? 0 : Math.Round((item.VND / item.TongDS) * 100, 2, MidpointRounding.ToEven),
                        USD = item.TongDS == 0 ? 0 : Math.Round((item.USD / item.TongDS) * 100, 2, MidpointRounding.ToEven),
                        EUR = item.TongDS == 0 ? 0 : Math.Round((item.EUR / item.TongDS) * 100, 2, MidpointRounding.ToEven),
                        CAD = item.TongDS == 0 ? 0 : Math.Round((item.CAD / item.TongDS) * 100, 2, MidpointRounding.ToEven),
                        AUD = item.TongDS == 0 ? 0 : Math.Round((item.AUD / item.TongDS) * 100, 2, MidpointRounding.ToEven),
                        GBP = item.TongDS == 0 ? 0 : Math.Round((item.GBP / item.TongDS) * 100, 2, MidpointRounding.ToEven),
                        TongDS = 100,
                        Month = item.Month,
                        Year = item.Year
                    }
                );
            }

            // # dòng record
            GradationChartPie[] arrayGradation = new GradationChartPie[1];
            arrayGradation[0] = new GradationChartPie()
            {
                category = "1",
                value = 0,
                color = "#9de219"

            };

            int count = 0;
            int lastMonth = 0;
            int lastYear = 0;
            foreach (ReportDetailtForTotalMoneyType item in listDataCompareMonthConvert)
            {
                lastMonth = (month - 1);
                lastYear = year;

                // Trường hợp thuộc tháng 1
                if (month == 1)
                {
                    lastMonth = 12;
                    lastYear = year - 1;
                }

                if (item.Year == lastYear.ToString() && item.Month == lastMonth.ToString())
                {
                    // tạo mảng gồm 8 object
                    arrayGradation = new GradationChartPie[6];

                    // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                    arrayGradation[count] = new GradationChartPie()
                    {
                        category = "VND",
                        value = item.VND,
                        color = "#FFBF00"
                    };

                    count++;

                    arrayGradation[count] = new GradationChartPie()
                    {
                        category = "USD",
                        value = item.USD,
                        color = "#40FF00"
                    };

                    count++;

                    arrayGradation[count] = new GradationChartPie()
                    {
                        category = "EUR",
                        value = item.EUR,
                        color = "#2ECCFA"
                    };

                    count++;

                    arrayGradation[count] = new GradationChartPie()
                    {
                        category = "CAD",
                        value = item.CAD,
                        color = "#9A2EFE"
                    };

                    count++;

                    arrayGradation[count] = new GradationChartPie()
                    {
                        category = "AUD",
                        value = item.AUD,
                        color = "#FE2EF7"
                    };

                    count++;

                    arrayGradation[count] = new GradationChartPie()
                    {
                        category = "GBP",
                        value = item.GBP,
                        color = "#0000FF"
                    };

                    count++;

                }
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
        public ActionResult SearchReportDetailtCompareMonthForOnePercent([DataSourceRequest]DataSourceRequest request, int month, int year, string reportTypeID, string partnerID)
        {
            List<ReportDetailtForTotalMoneyType> listDataCompareMonth = new ReportBL().ReportDetailtPartnerLTCompareMonthForOneConvert(year, month, reportTypeID, partnerID);
            List<ReportDetailtForTotalMoneyType> listDataCompareMonthConvert = new List<ReportDetailtForTotalMoneyType>();

            foreach (ReportDetailtForTotalMoneyType item in listDataCompareMonth)
            {
                item.PartnerName = string.Format("Tháng {0}/{1}", item.Month, item.Year);
                item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;

                listDataCompareMonthConvert.Add(
                    new ReportDetailtForTotalMoneyType()
                    {
                        PartnerName = string.Format("Tháng {0}/{1}", item.Month, item.Year),
                        VND = item.TongDS == 0 ? 0 : Math.Round((item.VND / item.TongDS) * 100, 2, MidpointRounding.ToEven),
                        USD = item.TongDS == 0 ? 0 : Math.Round((item.USD / item.TongDS) * 100, 2, MidpointRounding.ToEven),
                        EUR = item.TongDS == 0 ? 0 : Math.Round((item.EUR / item.TongDS) * 100, 2, MidpointRounding.ToEven),
                        CAD = item.TongDS == 0 ? 0 : Math.Round((item.CAD / item.TongDS) * 100, 2, MidpointRounding.ToEven),
                        AUD = item.TongDS == 0 ? 0 : Math.Round((item.AUD / item.TongDS) * 100, 2, MidpointRounding.ToEven),
                        GBP = item.TongDS == 0 ? 0 : Math.Round((item.GBP / item.TongDS) * 100, 2, MidpointRounding.ToEven),
                        TongDS = 100,
                        Month = item.Month,
                        Year = item.Year
                    }
                );
            }

            DataTable table = new DataTable();
            // Khởi tạo datatable
            table = new DataTable();
            // Tạo các cột cho datatable
            table.Columns.Add("PartnerName", typeof(String));

            table.Columns.Add("COL1", typeof(double));
            table.Columns.Add("COL2", typeof(double));
            table.Columns.Add("COL3", typeof(double));

            table.Columns.Add("TDS1", typeof(double));
            table.Columns.Add("TDS2", typeof(double));

            string[] listTypeMoney = { "VND", "USD", "EUR", "CAD", "AUD", "GBP" };

            if (listDataCompareMonthConvert.Count > 0)
            {
                ReportDetailtForTotalMoneyType dataItemLastYear = listDataCompareMonthConvert.Find(x => x.Month == month.ToString() && x.Year == (year - 1).ToString());
                ReportDetailtForTotalMoneyType dataItemYear = listDataCompareMonthConvert.Find(x => x.Month == month.ToString() && x.Year == year.ToString());
                ReportDetailtForTotalMoneyType dataItemLastMonth = listDataCompareMonthConvert.Find(x => x.Month == (month - 1).ToString() && x.Year == year.ToString());
                // Trường hợp tháng 1
                if (month == 1)
                {
                    dataItemLastMonth = listDataCompareMonthConvert.Find(x => x.Month == "12" && x.Year == (year - 1).ToString());
                }

                foreach (string item in listTypeMoney)
                {
                    // Tháng hiện tại
                    var propertyInfoYear = dataItemYear.GetType().GetProperty(item);
                    var valueDataYear = propertyInfoYear.GetValue(dataItemYear, null);

                    // Tháng trước
                    var propertyInfoLastMonth = dataItemLastMonth.GetType().GetProperty(item);
                    var valueDataLastMonth = propertyInfoLastMonth.GetValue(dataItemLastMonth, null);

                    //Cùng kì
                    var propertyInfoLastYear = dataItemLastYear.GetType().GetProperty(item);
                    var valueDataLastYear = propertyInfoYear.GetValue(dataItemLastYear, null);

                    double sumYear = Math.Round(Convert.ToDouble(valueDataYear) - Convert.ToDouble(valueDataLastMonth), 2, MidpointRounding.ToEven);
                    double sumLastYear = Math.Round(Convert.ToDouble(valueDataYear) - Convert.ToDouble(valueDataLastYear), 2, MidpointRounding.ToEven);

                    table.Rows.Add(
                        item
                        , valueDataYear, valueDataLastMonth, valueDataLastYear, sumYear, sumLastYear
                    );

                }
            }

            DataRow row = table.NewRow();
            row["PartnerName"] = "Tổng";
            row["COL1"] = 100;
            row["COL2"] = 100;
            row["COL3"] = 100;

            row["TDS1"] = table.Compute("Sum(TDS1)", "");
            row["TDS2"] = table.Compute("Sum(TDS2)", "");

            return Json(table.ToDataSourceResult(request), JsonRequestBehavior.AllowGet);
        }
    }
}
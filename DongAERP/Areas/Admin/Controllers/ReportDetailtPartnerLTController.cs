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
            string nameUrl = "Báo cáo chi tiết/Theo doanh số theo đối tác/Loại hình dịch vụ/Tất cả/Theo ngày";
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
            string nameUrl = "Báo cáo chi tiết/Theo doanh số theo đối tác/Loại hình dịch vụ/Tất cả/Theo tháng";
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
            string nameUrl = "Báo cáo chi tiết/Theo doanh số theo đối tác/Loại hình dịch vụ/Tất cả/Theo năm";
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
    }
}
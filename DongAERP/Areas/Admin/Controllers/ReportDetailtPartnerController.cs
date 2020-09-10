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
    public class ReportDetailtPartnerController : Controller
    {
        // GET: Admin/ReportDetailtPartner
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
        /// Màn hình báo cáo cho tháng
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
        /// Màn hình báo cáo cho năm
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
            string nameUrl = "Báo cáo chi tiết/Theo doanh số chi trả/Theo thị trường/Loại hình dịch vụ/Từng thị trường";
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
            string nameUrl = "Báo cáo chi tiết/Theo đối tác/Loại hình chi trả/So sánh - Giai đoạn - Tất cả";
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
            string nameUrl = "Báo cáo chi tiết/Theo đối tác/Loại hình chi trả/So sánh - Giai đoạn - Từng đối tác";
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
            // Khởi tạo datatable
            DataTable table = new DataTable();
            // Tạo các cột cho datatable
            table.Columns.Add("MarketName", typeof(String));
            table.Columns.Add("CQ1", typeof(double));
            table.Columns.Add("CQ2", typeof(double));
            table.Columns.Add("CQ3", typeof(double));

            table.Columns.Add("CN1", typeof(double));
            table.Columns.Add("CN2", typeof(double));
            table.Columns.Add("CN3", typeof(double));

            table.Columns.Add("CK1", typeof(double));
            table.Columns.Add("CK2", typeof(double));
            table.Columns.Add("CK3", typeof(double));

            table.Columns.Add("TDS1", typeof(double));
            table.Columns.Add("TDS2", typeof(double));
            table.Columns.Add("TDS3", typeof(double));

            table.Columns.Add("ReportID", typeof(String));

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

            return View(table);
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
            List<ReportDetailtForPartner> listData = new ReportBL().SearchReportDetailtPartnerForDay(fromDay, toDay, reportTypeID);

            int count = 1;
            foreach (ReportDetailtForPartner item in listData)
            {
                item.STT = (count++).ToString();
                item.TongDS = item.DSChiNha + item.DSChiQuay + item.DSCK;
            }

            listData.Add(
                new ReportDetailtForPartner()
                {
                    PartnerName = "Tổng",
                    DSChiQuay = listData.Sum(x => x.DSChiQuay),
                    DSChiNha = listData.Sum(x => x.DSChiNha),
                    DSCK = listData.Sum(x => x.DSCK),
                    TongDS = listData.Sum(x => x.TongDS)
                }
            );

            return Json(listData.ToDataSourceResult(request), JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// Search report detailt theo tháng
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public ActionResult SearchPartnerForTotalForMonth([DataSourceRequest]DataSourceRequest request, DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            List<ReportDetailtForPartner> listData = new ReportBL().SearchPartnerForTotalForMonth(fromDate, toDate, reportTypeID);

            int count = 1;
            foreach (ReportDetailtForPartner item in listData)
            {
                item.STT = (count++).ToString();
                item.TongDS = item.DSChiNha + item.DSChiQuay + item.DSCK;
            }

            listData.Add(
                new ReportDetailtForPartner()
                {
                    PartnerName = "Tổng",
                    DSChiQuay = listData.Sum(x => x.DSChiQuay),
                    DSChiNha = listData.Sum(x => x.DSChiNha),
                    DSCK = listData.Sum(x => x.DSCK),
                    TongDS = listData.Sum(x => x.TongDS)
                }
            );

            return Json(listData.ToDataSourceResult(request), JsonRequestBehavior.AllowGet);
        }


        /// <summary>
        /// Search report detailt theo tháng
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public ActionResult SearchPartnerForTotalForYear([DataSourceRequest]DataSourceRequest request, DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            List<ReportDetailtForPartner> listData = new ReportBL().SearchPartnerForTotalForYear(fromDate, toDate, reportTypeID);

            int count = 1;
            foreach (ReportDetailtForPartner item in listData)
            {
                item.STT = (count++).ToString();
                item.TongDS = item.DSChiNha + item.DSChiQuay + item.DSCK;
            }

            listData.Add(
                new ReportDetailtForPartner()
                {
                    PartnerName = "Tổng",
                    DSChiQuay = listData.Sum(x => x.DSChiQuay),
                    DSChiNha = listData.Sum(x => x.DSChiNha),
                    DSCK = listData.Sum(x => x.DSCK),
                    TongDS = listData.Sum(x => x.TongDS)
                }
            );

            return Json(listData.ToDataSourceResult(request), JsonRequestBehavior.AllowGet);
        }
        
        /// <summary>
        /// Search report chi tiết theo ngày cho từng đối tác
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/09/2020]
        /// </history>
        public ActionResult SearchMarketForOne([DataSourceRequest]DataSourceRequest request, DateTime fromDay, DateTime toDay, string reportTypeID, string partnerID)
        {
            List<ReportDetailtForPartner> listData = new ReportBL().SearchPartnerForOne(fromDay, toDay, reportTypeID, partnerID);
            int count = 1;
            foreach (ReportDetailtForPartner item in listData)
            {
                item.STT = item.CreatedDate.ToString("dd/MM/yyyy");
                item.TongDS = item.DSChiNha + item.DSChiQuay + item.DSCK;
            }

            listData.Add(
                new ReportDetailtForPartner()
                {
                    STT = "Tổng",
                    DSChiQuay = listData.Sum(x => x.DSChiQuay),
                    DSChiNha = listData.Sum(x => x.DSChiNha),
                    DSCK = listData.Sum(x => x.DSCK),
                    TongDS = listData.Sum(x => x.TongDS)
                }
            );

            return Json(listData.ToDataSourceResult(request), JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// Search report chi tiết theo ngày cho từng đối tác
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/09/2020]
        /// </history>
        public ActionResult SearchMarketForOneForMonth([DataSourceRequest]DataSourceRequest request, DateTime fromDate, DateTime toDate, string reportTypeID, string partnerID)
        {
            List<ReportDetailtForPartner> listData = new ReportBL().SearchPartnerForOneForMonth(fromDate, toDate, reportTypeID, partnerID);
            int count = 1;
            foreach (ReportDetailtForPartner item in listData)
            {
                item.STT = string.Format("Tháng {0}/{1}", item.Month, item.Year);
                item.TongDS = item.DSChiNha + item.DSChiQuay + item.DSCK;
            }

            listData.Add(
                new ReportDetailtForPartner()
                {
                    STT = "Tổng",
                    DSChiQuay = listData.Sum(x => x.DSChiQuay),
                    DSChiNha = listData.Sum(x => x.DSChiNha),
                    DSCK = listData.Sum(x => x.DSCK),
                    TongDS = listData.Sum(x => x.TongDS)
                }
            );

            return Json(listData.ToDataSourceResult(request), JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// Search report chi tiết theo ngày cho từng đối tác
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/09/2020]
        /// </history>
        public ActionResult SearchMarketForOneForYear([DataSourceRequest]DataSourceRequest request, DateTime fromDate, DateTime toDate, string reportTypeID, string partnerID)
        {
            List<ReportDetailtForPartner> listData = new ReportBL().SearchPartnerForOneForYear(fromDate, toDate, reportTypeID, partnerID);

            foreach (ReportDetailtForPartner item in listData)
            {
                item.STT = string.Format("năm {0}", item.Year);
                item.TongDS = item.DSChiNha + item.DSChiQuay + item.DSCK;
            }

            listData.Add(
                new ReportDetailtForPartner()
                {
                    STT = "Tổng",
                    DSChiQuay = listData.Sum(x => x.DSChiQuay),
                    DSChiNha = listData.Sum(x => x.DSChiNha),
                    DSCK = listData.Sum(x => x.DSCK),
                    TongDS = listData.Sum(x => x.TongDS)
                }
            );

            return Json(listData.ToDataSourceResult(request), JsonRequestBehavior.AllowGet);
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
            List<ReportDetailtForPartner> listDataGradation = new ReportBL().ReportDetailtPartnerGradationCompareForAll(year, gradation, reportTypeID);

            // Khởi tạo datatable
            DataTable table = new DataTable();
            foreach (ReportDetailtForPartner item in listDataGradation)
            {
                item.TongDS = item.DSChiQuay + item.DSChiNha + item.DSCK;
            }

            // Khởi tạo datatable
            table = new DataTable();
            // Tạo các cột cho datatable
            table.Columns.Add("STT", typeof(String));
            table.Columns.Add("PartnerName", typeof(String));

            table.Columns.Add("CQ1", typeof(double));
            table.Columns.Add("CQ2", typeof(double));
            table.Columns.Add("CQ3", typeof(double));

            table.Columns.Add("CN1", typeof(double));
            table.Columns.Add("CN2", typeof(double));
            table.Columns.Add("CN3", typeof(double));

            table.Columns.Add("CK1", typeof(double));
            table.Columns.Add("CK2", typeof(double));
            table.Columns.Add("CK3", typeof(double));

            table.Columns.Add("TDS1", typeof(double));
            table.Columns.Add("TDS2", typeof(double));
            table.Columns.Add("TDS3", typeof(double));
            
            if (listDataGradation.Count() > 0)
            {
                List<string> listPartnerName = new List<string>();
                int count = 1;
                foreach (ReportDetailtForPartner item in listDataGradation)
                {
                    if (listPartnerName.Contains(item.PartnerName))
                    {
                        continue;
                    }
                    // Thêm đối tác vào bảng tạm
                    listPartnerName.Add(item.PartnerName);

                    // Cùng kì
                    ReportDetailtForPartner dataItemLastYear = listDataGradation.Find(x => x.PartnerCode == item.PartnerCode && x.Year == (year - 1).ToString());
                    ReportDetailtForPartner dataItemYear = listDataGradation.Find(x => x.PartnerCode == item.PartnerCode && x.Year == year.ToString());

                    // Trường hợp năm hiện tại không có
                    if (dataItemLastYear == null)
                    {
                        dataItemLastYear = new ReportDetailtForPartner()
                        {
                            PartnerCode = item.PartnerCode,
                            PartnerName = item.PartnerName,
                            Year = (year - 1).ToString()
                        };
                    }

                    // Trường hợp năm hiện tại không có
                    if (dataItemYear == null)
                    {
                        dataItemYear = new ReportDetailtForPartner()
                        {
                            PartnerCode = item.PartnerCode,
                            PartnerName = item.PartnerName,
                            Year = year.ToString()
                        };
                    }
                    // add item vào table
                    table.Rows.Add(
                        count++ 
                        , item.PartnerName
                        , dataItemYear.DSChiQuay, dataItemLastYear.DSChiQuay, Math.Round(dataItemYear.DSChiQuay - dataItemLastYear.DSChiQuay, 2, MidpointRounding.ToEven)
                        , dataItemYear.DSChiNha, dataItemLastYear.DSChiNha, Math.Round(dataItemYear.DSChiNha - dataItemLastYear.DSChiNha, 2, MidpointRounding.ToEven)
                        , dataItemYear.DSCK, dataItemLastYear.DSCK, Math.Round(dataItemYear.DSCK - dataItemLastYear.DSCK, 2, MidpointRounding.ToEven)
                        , dataItemYear.TongDS, dataItemLastYear.TongDS, Math.Round(dataItemYear.TongDS - dataItemLastYear.TongDS, 2, MidpointRounding.ToEven)
                    );
                }

                DataRow row = table.NewRow();
                row["STT"] = "";
                row["PartnerName"] = "Tổng";
                row["CQ1"] = table.Compute("Sum(CQ1)", "");
                row["CQ2"] = table.Compute("Sum(CQ2)", "");
                row["CQ3"] = table.Compute("Sum(CQ3)", "");

                row["CN1"] = table.Compute("Sum(CN1)", "");
                row["CN2"] = table.Compute("Sum(CN2)", "");
                row["CN3"] = table.Compute("Sum(CN3)", "");

                row["CK1"] = table.Compute("Sum(CK1)", "");
                row["CK2"] = table.Compute("Sum(CK2)", "");
                row["CK3"] = table.Compute("Sum(CK3)", "");

                row["TDS1"] = table.Compute("Sum(TDS1)", "");
                row["TDS2"] = table.Compute("Sum(TDS2)", "");
                row["TDS3"] = table.Compute("Sum(TDS3)", "");
                table.Rows.Add(row);
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
        public ActionResult SearchGridReportForGradationForOne([DataSourceRequest]DataSourceRequest request, int gradation, int year, string reportTypeID, string partnerID)
        {
            List<ReportDetailtForPartner> listDataGradation = new ReportBL().ReportDetailtPartnerGradationCompareForOne(year, gradation, reportTypeID, partnerID);

            string text = string.Empty;

            switch(gradation)
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

            foreach (ReportDetailtForPartner item in listDataGradation)
            {
                item.PartnerName = string.Format("Lũy kế {0} năm {1}", text, item.Year);
                item.TongDS = item.DSChiQuay + item.DSChiNha + item.DSCK;
            }
            
            if(listDataGradation.Count().Equals(2))
            {
                double sumDSChiQuay = listDataGradation[1].DSChiQuay - listDataGradation[0].DSChiQuay;
                double sumDSChiNha = listDataGradation[1].DSChiNha - listDataGradation[0].DSChiNha;
                double sumDSCK = listDataGradation[1].DSCK - listDataGradation[0].DSCK;
                double sumTongDS = listDataGradation[1].TongDS - listDataGradation[0].TongDS;

                // Tăng giảm
                listDataGradation.Add(
                    new ReportDetailtForPartner()
                    {
                        PartnerName = string.Format("Tăng giảm so với cùng kì {0} (+/-)", year - 1),
                        DSChiQuay = Math.Round(sumDSChiQuay, 2, MidpointRounding.ToEven),
                        DSChiNha = Math.Round(sumDSChiNha, 2, MidpointRounding.ToEven),
                        DSCK = Math.Round(sumDSCK, 2, MidpointRounding.ToEven),
                        TongDS = Math.Round(sumTongDS, 2, MidpointRounding.ToEven),
                    }
                );

                // tỉ trọng
                listDataGradation.Add(
                    new ReportDetailtForPartner()
                    {
                        PartnerName = string.Format("Tăng giảm so với cùng kì {0} (%)", year - 1),
                        DSChiQuay = listDataGradation[0].DSChiQuay == 0 ? 0 : Math.Round((sumDSChiQuay/listDataGradation[0].DSChiQuay) * 100, 2, MidpointRounding.ToEven),
                        DSChiNha = listDataGradation[0].DSChiNha == 0 ? 0 : Math.Round((sumDSChiNha/listDataGradation[0].DSChiNha) * 100, 2, MidpointRounding.ToEven),
                        DSCK = listDataGradation[0].DSCK == 0 ? 0 : Math.Round((sumDSCK/listDataGradation[0].DSCK) * 100, 2, MidpointRounding.ToEven),
                        TongDS = listDataGradation[0].TongDS == 0 ? 0 : Math.Round((sumTongDS/listDataGradation[0].TongDS) * 100, 2, MidpointRounding.ToEven),
                    }
                );
            }
            
            return Json(listDataGradation.ToDataSourceResult(request), JsonRequestBehavior.AllowGet);
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
            List<ReportDetailtForPartner> listDataGradation = new ReportBL().ReportDetailtPartnerGradationCompareForOne(year, gradation, reportTypeID, partnerID);
            List<ReportDetailtForPartner> listDataGradationConvert = new List<ReportDetailtForPartner>();

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

            foreach (ReportDetailtForPartner item in listDataGradation)
            {
                item.PartnerName = string.Format("Lũy kế {0} năm {1}", text, item.Year);
                item.TongDS = item.DSChiQuay + item.DSChiNha + item.DSCK;

                listDataGradationConvert.Add(
                    new ReportDetailtForPartner()
                    {
                        PartnerCode = item.PartnerCode,
                        PartnerName = item.PartnerName,
                        DSChiQuay = item.TongDS == 0 ? 0 : Math.Round((item.DSChiQuay / item.TongDS) * 100, 2, MidpointRounding.ToEven),
                        DSChiNha = item.TongDS == 0 ? 0 : Math.Round((item.DSChiNha / item.TongDS) * 100, 2, MidpointRounding.ToEven),
                        DSCK = item.TongDS == 0 ? 0 : Math.Round((item.DSCK / item.TongDS) * 100, 2, MidpointRounding.ToEven),
                        TongDS = 100
                    }
                );
            }

            double sumDSChiQuay = 0;
            double sumDSChiNha = 0;
            double sumDSCK = 0;
            double sumTongDS = 0;

            if (listDataGradation.Count().Equals(2))
            {
                sumDSChiQuay = listDataGradation[1].DSChiQuay - listDataGradation[0].DSChiQuay;
                sumDSChiNha = listDataGradation[1].DSChiNha - listDataGradation[0].DSChiNha;
                sumDSCK = listDataGradation[1].DSCK - listDataGradation[0].DSCK;
                sumTongDS = listDataGradation[1].TongDS - listDataGradation[0].TongDS;
            }

            if (listDataGradationConvert.Count().Equals(2))
            {
                listDataGradationConvert.Add(
                    new ReportDetailtForPartner()
                    {
                        PartnerName = string.Format("Tăng giảm so với cùng kì {0}", year - 1),
                        DSChiQuay = listDataGradation[0].DSChiQuay == 0 ? 0 : Math.Round((sumDSChiQuay / listDataGradation[0].DSChiQuay) * 100, 2, MidpointRounding.ToEven),
                        DSChiNha = listDataGradation[0].DSChiNha == 0 ? 0 : Math.Round((sumDSChiNha/ listDataGradation[0].DSChiNha) * 100, 2, MidpointRounding.ToEven),
                        DSCK = listDataGradation[0].DSCK == 0 ? 0 : Math.Round((sumDSCK/ listDataGradation[0].DSCK) * 100, 2, MidpointRounding.ToEven),
                    }
                );
            }
            

            return Json(listDataGradationConvert.ToDataSourceResult(request), JsonRequestBehavior.AllowGet);
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
            List<ReportDetailtForPartner> listData = new ReportBL().ReportDetailtPartnerGradationCompareForOne(year, int.Parse(gradation), reportTypeID, partnerID);
            List<ReportDetailtForPartner> listDataConvert = new List<ReportDetailtForPartner>();

            foreach (ReportDetailtForPartner item in listData)
            {
                if (item.Year == year.ToString())
                {
                    item.TongDS = item.DSChiQuay + item.DSChiNha + item.DSCK;

                    listDataConvert.Add(
                        new ReportDetailtForPartner()
                        {
                            PartnerCode = item.PartnerCode,
                            PartnerName = item.PartnerName,
                            DSChiQuay = item.TongDS == 0 ? 0 : Math.Round((item.DSChiQuay / item.TongDS) * 100, 2, MidpointRounding.ToEven),
                            DSChiNha = item.TongDS == 0 ? 0 : Math.Round((item.DSChiNha / item.TongDS) * 100, 2, MidpointRounding.ToEven),
                            DSCK = item.TongDS == 0 ? 0 : Math.Round((item.DSCK / item.TongDS) * 100, 2, MidpointRounding.ToEven),
                            TongDS = 100,
                            Year = item.Year
                        }
                    );
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
            foreach (ReportDetailtForPartner item in listDataConvert)
            {
                if (item.Year == year.ToString())
                {
                    // tạo mảng gồm 8 object
                    arrayGradation = new GradationChartPie[3];

                    // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                    arrayGradation[count] = new GradationChartPie()
                    {
                        category = "Chi Quầy",
                        value = item.DSChiQuay,
                        color = "#FFBF00"
                    };

                    count++;
                    arrayGradation[count] = new GradationChartPie()
                    {
                        category = "Chi Nhà",
                        value = item.DSChiNha,
                        color = "#40FF00"
                    };

                    count++;
                    arrayGradation[count] = new GradationChartPie()
                    {
                        category = "Chuyển Khoản",
                        value = item.DSCK,
                        color = "#2ECCFA"
                    };

                    count++;
                }
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
            List<ReportDetailtForPartner> listData = new ReportBL().ReportDetailtPartnerGradationCompareForOne(year, int.Parse(gradation), reportTypeID, partnerID);
            List<ReportDetailtForPartner> listDataConvert = new List<ReportDetailtForPartner>();

            foreach (ReportDetailtForPartner item in listData)
            {
                if (item.Year == (year - 1).ToString())
                {
                    item.TongDS = item.DSChiQuay + item.DSChiNha + item.DSCK;

                    listDataConvert.Add(
                        new ReportDetailtForPartner()
                        {
                            PartnerCode = item.PartnerCode,
                            PartnerName = item.PartnerName,
                            DSChiQuay = item.TongDS == 0 ? 0 : Math.Round((item.DSChiQuay / item.TongDS) * 100, 2, MidpointRounding.ToEven),
                            DSChiNha = item.TongDS == 0 ? 0 : Math.Round((item.DSChiNha / item.TongDS) * 100, 2, MidpointRounding.ToEven),
                            DSCK = item.TongDS == 0 ? 0 : Math.Round((item.DSCK / item.TongDS) * 100, 2, MidpointRounding.ToEven),
                            TongDS = 100,
                            Year = item.Year
                        }
                    );
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
            foreach (ReportDetailtForPartner item in listDataConvert)
            {
                if (item.Year == (year - 1).ToString())
                {
                    // tạo mảng gồm 8 object
                    arrayGradation = new GradationChartPie[3];

                    // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                    arrayGradation[count] = new GradationChartPie()
                    {
                        category = "Chi Quầy",
                        value = item.DSChiQuay,
                        color = "#FFBF00"
                    };

                    count++;
                    arrayGradation[count] = new GradationChartPie()
                    {
                        category = "Chi Nhà",
                        value = item.DSChiNha,
                        color = "#40FF00"
                    };

                    count++;
                    arrayGradation[count] = new GradationChartPie()
                    {
                        category = "Chuyển Khoản",
                        value = item.DSCK,
                        color = "#2ECCFA"
                    };

                    count++;
                }
            }

            return Json(arrayGradation);
        }
    }
}
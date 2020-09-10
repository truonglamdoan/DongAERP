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
            
            foreach (ReportDetailtForPartner item in listDataGradation)
            {
                item.TongDS = item.DSChiQuay + item.DSChiNha + item.DSCK;
            }
            
            return Json(listDataGradation.ToDataSourceResult(request), JsonRequestBehavior.AllowGet);
        }
    }
}
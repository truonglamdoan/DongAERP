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
    public class ReportHSDetailtPartnerLHController : Controller
    {
        // GET: Admin/ReportHSDetailtPartnerLH
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
            string nameUrl = "Báo cáo chi tiết/Theo Hồ sơ theo đối tác/Loại hình dịch vụ/Tất cả/Theo ngày";
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
        /// Search report day theo ngày nhập vào
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public ActionResult SearchPartnerForTotal([DataSourceRequest]DataSourceRequest request, DateTime fromDay, DateTime toDay, string reportTypeID)
        {
            List<ReportDetailtForPartner> listData = new HSReportBL().SearchReportDetailtPartnerForDay(fromDay, toDay, reportTypeID);

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
        /// Màn hình báo cáo cho tháng
        /// </summary>
        /// <returns></returns>
        public ActionResult PartnerForTotalForMonth(DateTime? fromDate, DateTime? toDate, string reportTypeID)
        {
            string nameUrl = "Báo cáo chi tiết/Theo Hồ sơ theo đối tác/Loại hình dịch vụ/Tất cả/Theo tháng";
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
        /// Search report detailt theo tháng
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public ActionResult SearchPartnerForTotalForMonth([DataSourceRequest]DataSourceRequest request, DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            List<ReportDetailtForPartner> listData = new HSReportBL().SearchReportDetailtPartnerForMonth(fromDate, toDate, reportTypeID);

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
        /// Màn hình báo cáo cho năm
        /// </summary>
        /// <returns></returns>
        public ActionResult PartnerForTotalForYear(DateTime? fromDate, DateTime? toDate, string reportTypeID)
        {
            string nameUrl = "Báo cáo chi tiết/Theo Hồ sơ theo đối tác/Loại hình dịch vụ/Tất cả/Theo năm";
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
        /// Search report detailt theo tháng
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public ActionResult SearchPartnerForTotalForYear([DataSourceRequest]DataSourceRequest request, DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            List<ReportDetailtForPartner> listData = new HSReportBL().SearchReportDetailtPartnerForYear(fromDate, toDate, reportTypeID);

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
        /// Màn hình báo cáo chi tiết theo ngày theo từng đối tác
        /// </summary>
        /// <returns></returns>
        public ActionResult PartnerForOne(DateTime? fromDay, DateTime? toDay, string reportTypeID, string partnerID)
        {
            string nameUrl = "Báo cáo chi tiết/Theo Hồ sơ chi trả/Theo thị trường/Loại hình dịch vụ/Từng thị trường";
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
        /// Search report chi tiết theo ngày cho từng đối tác
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/09/2020]
        /// </history>
        public ActionResult SearchMarketForOne([DataSourceRequest]DataSourceRequest request, DateTime fromDay, DateTime toDay, string reportTypeID, string partnerID)
        {
            List<ReportDetailtForPartner> listData = new HSReportBL().SearchPartnerForOne(fromDay, toDay, reportTypeID, partnerID);
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
        /// Search report day theo ngày nhập vào
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult SearchLineChartMarketForOne([DataSourceRequest]DataSourceRequest request, DateTime fromDay, DateTime toDay, string reportTypeID, string partnerID)
        {
            List<ReportDetailtForPartner> listData = new HSReportBL().SearchPartnerForOne(fromDay, toDay, reportTypeID, partnerID);

            foreach (ReportDetailtForPartner item in listData)
            {
                item.STT = item.CreatedDate.ToString("dd/MM/yyyy");
                item.TongDS = item.DSChiNha + item.DSChiQuay + item.DSCK;
            }

            return Json(listData);
        }



        /// <summary>
        /// Màn hình báo cáo chi tiết theo ngày theo từng đối tác
        /// </summary>
        /// <returns></returns>
        public ActionResult PartnerForOneForMonth(DateTime? fromDate, DateTime? toDate, string reportTypeID, string partnerID)
        {
            string nameUrl = "Báo cáo chi tiết/Theo Hồ sơ chi trả/Theo thị trường/Loại hình dịch vụ/Từng thị trường";
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
        /// Search report chi tiết theo ngày cho từng đối tác
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/09/2020]
        /// </history>
        public ActionResult SearchMarketForOneForMonth([DataSourceRequest]DataSourceRequest request, DateTime fromDate, DateTime toDate, string reportTypeID, string partnerID)
        {
            List<ReportDetailtForPartner> listData = new HSReportBL().SearchPartnerForOneForMonth(fromDate, toDate, reportTypeID, partnerID);

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
        /// Search report day theo ngày nhập vào
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult SearchLineChartMarketForOneForMonth([DataSourceRequest]DataSourceRequest request, DateTime fromDate, DateTime toDate, string reportTypeID, string partnerID)
        {
            List<ReportDetailtForPartner> listData = new HSReportBL().SearchPartnerForOneForMonth(fromDate, toDate, reportTypeID, partnerID);

            foreach (ReportDetailtForPartner item in listData)
            {
                item.STT = string.Format("Tháng {0}/{1}", item.Month, item.Year);
                item.TongDS = item.DSChiNha + item.DSChiQuay + item.DSCK;
            }

            return Json(listData);
        }


        /// <summary>
        /// Màn hình báo cáo chi tiết theo ngày theo từng đối tác
        /// </summary>
        /// <returns></returns>
        public ActionResult PartnerForOneForYear(DateTime? fromDate, DateTime? toDate, string reportTypeID, string partnerID)
        {
            string nameUrl = "Báo cáo chi tiết/Theo Hồ sơ chi trả/Theo thị trường/Loại hình dịch vụ/Từng thị trường";
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
        /// Search report chi tiết theo ngày cho từng đối tác
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/09/2020]
        /// </history>
        public ActionResult SearchMarketForOneForYear([DataSourceRequest]DataSourceRequest request, DateTime fromDate, DateTime toDate, string reportTypeID, string partnerID)
        {
            List<ReportDetailtForPartner> listData = new HSReportBL().SearchPartnerForOneForYear(fromDate, toDate, reportTypeID, partnerID);

            foreach (ReportDetailtForPartner item in listData)
            {
                item.STT = string.Format("Năm {0}", item.Year);
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
        /// Search report day theo ngày nhập vào
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult SearchLineChartMarketForOneForYear([DataSourceRequest]DataSourceRequest request, DateTime fromDate, DateTime toDate, string reportTypeID, string partnerID)
        {
            List<ReportDetailtForPartner> listData = new HSReportBL().SearchPartnerForOneForYear(fromDate, toDate, reportTypeID, partnerID);

            foreach (ReportDetailtForPartner item in listData)
            {
                item.STT = string.Format("Năm {0}", item.Year);
                item.TongDS = item.DSChiNha + item.DSChiQuay + item.DSCK;
            }

            return Json(listData);
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
            List<ReportDetailtForPartner> listDataGradation = new HSReportBL().ReportDetailtPartnerGradationCompareForAll(year, gradation, reportTypeID);

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
        /// Get data cho việc vẽ biểu đồ cột cho so sánh theo giai đoạn
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult SearchGridReportForGradationForOne([DataSourceRequest]DataSourceRequest request, int gradation, int year, string reportTypeID, string partnerID)
        {
            List<ReportDetailtForPartner> listDataGradation = new HSReportBL().ReportDetailtPartnerGradationCompareForOne(year, gradation, reportTypeID, partnerID);

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
            }

            if (listDataGradation.Count().Equals(2))
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
                        DSChiQuay = listDataGradation[0].DSChiQuay == 0 ? 0 : Math.Round((sumDSChiQuay / listDataGradation[0].DSChiQuay) * 100, 2, MidpointRounding.ToEven),
                        DSChiNha = listDataGradation[0].DSChiNha == 0 ? 0 : Math.Round((sumDSChiNha / listDataGradation[0].DSChiNha) * 100, 2, MidpointRounding.ToEven),
                        DSCK = listDataGradation[0].DSCK == 0 ? 0 : Math.Round((sumDSCK / listDataGradation[0].DSCK) * 100, 2, MidpointRounding.ToEven),
                        TongDS = listDataGradation[0].TongDS == 0 ? 0 : Math.Round((sumTongDS / listDataGradation[0].TongDS) * 100, 2, MidpointRounding.ToEven),
                    }
                );
            }

            return Json(listDataGradation.ToDataSourceResult(request), JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// Get data cho việc vẽ biểu đồ cột cho so sánh theo tháng và cùng kì năm trước
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult SearchChartReportForGradation([DataSourceRequest]DataSourceRequest request, int gradation, int year, string reportTypeID, string partnerID)
        {
            List<ReportDetailtForPartner> listDataGradation = new HSReportBL().ReportDetailtPartnerGradationCompareForOne(year, gradation, reportTypeID, partnerID);
            GradationCompare[] arrayGradation = null;

            if (listDataGradation.Any(x => x.Year == year.ToString()) && listDataGradation.Any(x => x.Year == (year - 1).ToString()))
            {

                arrayGradation = new GradationCompare[6];
                string text = " tháng năm";
                switch (gradation)
                {
                    case 1:
                        text = string.Concat("3", text);
                        break;
                    case 2:
                        text = string.Concat("6", text);
                        break;
                    case 3:
                        text = string.Concat("9", text);
                        break;
                    default:
                        text = string.Concat("12", text);
                        break;
                }

                int count = 0;
                foreach (ReportDetailtForPartner item in listDataGradation)
                {
                    // tổng Hồ sơ
                    item.TongDS = item.DSChiNha + item.DSChiQuay + item.DSCK;

                    // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                    arrayGradation[count] = new GradationCompare()
                    {
                        NameGradationCompare = string.Format(" Lũy kế {0} {1}", text, item.Year),
                        amount = item.DSChiQuay,
                        NameType = "Hồ sơ \n chi quầy"
                    };

                    count++;
                    arrayGradation[count] = new GradationCompare()
                    {
                        NameGradationCompare = string.Format(" Lũy kế {0} {1}", text, item.Year),
                        amount = item.DSChiNha,
                        NameType = "Hồ sơ \n chi nhà"
                    };

                    count++;
                    arrayGradation[count] = new GradationCompare()
                    {
                        NameGradationCompare = string.Format(" Lũy kế {0} {1}", text, item.Year),
                        amount = item.DSCK,
                        NameType = "Hồ sơ \n chuyển khoản"
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
        /// Get data cho việc vẽ biểu đồ cột cho so sánh theo giai đoạn
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult SearchGridReportForGradationForOnePercent([DataSourceRequest]DataSourceRequest request, int gradation, int year, string reportTypeID, string partnerID)
        {
            List<ReportDetailtForPartner> listDataGradation = new HSReportBL().ReportDetailtPartnerGradationCompareForOne(year, gradation, reportTypeID, partnerID);
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
                        DSChiQuay = listDataGradationConvert[1].DSChiQuay - listDataGradationConvert[0].DSChiQuay,
                        DSChiNha = listDataGradationConvert[1].DSChiNha - listDataGradationConvert[0].DSChiNha,
                        DSCK = listDataGradationConvert[1].DSCK - listDataGradationConvert[0].DSCK,
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
        public ActionResult SearchGradationComparePieForLastYear([DataSourceRequest]DataSourceRequest request, string gradation, int year, string reportTypeID, string partnerID)
        {
            List<ReportDetailtForPartner> listData = new HSReportBL().ReportDetailtPartnerGradationCompareForOne(year, int.Parse(gradation), reportTypeID, partnerID);
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
            List<ReportDetailtForPartner> listData = new HSReportBL().ReportDetailtPartnerGradationCompareForOne(year, int.Parse(gradation), reportTypeID, partnerID);
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
            List<ReportDetailtForPartner> listDataCompareMonth = new HSReportBL().ReportDetailtPartnerCompareMonthForAll(year, month, reportTypeID);

            foreach (ReportDetailtForPartner item in listDataCompareMonth)
            {
                item.TongDS = item.DSChiQuay + item.DSChiNha + item.DSCK;
            }

            DataTable table = new DataTable();
            // Khởi tạo datatable
            table = new DataTable();
            // Tạo các cột cho datatable
            // tháng hiện tại
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


            List<string> listPartner = new List<string>();
            int count = 1;

            foreach (ReportDetailtForPartner item in listDataCompareMonth)
            {
                // Check tồn tại của đối tác
                if (listPartner.Contains(item.PartnerCode))
                {
                    continue;
                }
                listPartner.Add(item.PartnerCode);

                // Cùng kì
                ReportDetailtForPartner dataItemLastYear = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == month.ToString() && x.Year == (year - 1).ToString());
                ReportDetailtForPartner dataItemYear = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == month.ToString() && x.Year == year.ToString());
                ReportDetailtForPartner dataItemLastMonth = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == (month - 1).ToString() && x.Year == year.ToString());
                // Trường hợp tháng 1
                if (month == 1)
                {
                    dataItemLastMonth = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == "12" && x.Year == (year - 1).ToString());
                }
                // Cung kì
                if (dataItemLastYear == null)
                {
                    dataItemLastYear = new ReportDetailtForPartner()
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
                    dataItemYear = new ReportDetailtForPartner()
                    {
                        PartnerCode = item.PartnerCode,
                        PartnerName = item.PartnerName,
                        Month = month.ToString(),
                        Year = year.ToString()
                    };
                }

                // Tháng trước
                // Trường hợp tháng 1
                if (month == 1)
                {
                    if (dataItemLastMonth == null)
                    {
                        dataItemLastMonth = new ReportDetailtForPartner()
                        {
                            PartnerCode = item.PartnerCode,
                            PartnerName = item.PartnerName,
                            Month = "12",
                            Year = (year - 1).ToString()
                        };
                    }
                }
                else
                {
                    if (dataItemLastMonth == null)
                    {
                        dataItemLastMonth = new ReportDetailtForPartner()
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
                        , dataItemYear.DSChiQuay, dataItemLastMonth.DSChiQuay, dataItemLastYear.DSChiQuay
                        , dataItemYear.DSChiNha, dataItemLastMonth.DSChiNha, dataItemLastYear.DSChiNha
                        , dataItemYear.DSCK, dataItemLastMonth.DSCK, dataItemLastYear.DSCK
                        , dataItemYear.TongDS, dataItemLastMonth.TongDS, dataItemLastYear.TongDS
                    );
                }
            }

            // Add dòng tổng
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

            return Json(table.ToDataSourceResult(request), JsonRequestBehavior.AllowGet);
        }


        /// <summary>
        /// Get data cho báo cáo chi tiết so sánh theo tháng - compare
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult SearchReportDetailtCompareMonthForAllCompare([DataSourceRequest]DataSourceRequest request, int month, int year, string reportTypeID)
        {
            List<ReportDetailtForPartner> listDataCompareMonth = new HSReportBL().ReportDetailtPartnerCompareMonthForAll(year, month, reportTypeID);

            foreach (ReportDetailtForPartner item in listDataCompareMonth)
            {
                item.TongDS = item.DSChiQuay + item.DSChiNha + item.DSCK;
            }

            DataTable table = new DataTable();
            // Khởi tạo datatable
            table = new DataTable();
            // Tạo các cột cho datatable
            // tháng hiện tại
            table.Columns.Add("STT", typeof(String));

            table.Columns.Add("PartnerName", typeof(String));
            table.Columns.Add("CQ1", typeof(double));
            table.Columns.Add("CQ2", typeof(double));

            table.Columns.Add("CN1", typeof(double));
            table.Columns.Add("CN2", typeof(double));

            table.Columns.Add("CK1", typeof(double));
            table.Columns.Add("CK2", typeof(double));

            table.Columns.Add("TDS1", typeof(double));
            table.Columns.Add("TDS2", typeof(double));


            List<string> listPartner = new List<string>();
            int count = 1;

            foreach (ReportDetailtForPartner item in listDataCompareMonth)
            {
                // Check tồn tại của đối tác
                if (listPartner.Contains(item.PartnerCode))
                {
                    continue;
                }
                listPartner.Add(item.PartnerCode);

                // Cùng kì
                ReportDetailtForPartner dataItemLastYear = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == month.ToString() && x.Year == (year - 1).ToString());
                ReportDetailtForPartner dataItemYear = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == month.ToString() && x.Year == year.ToString());
                ReportDetailtForPartner dataItemLastMonth = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == (month - 1).ToString() && x.Year == year.ToString());
                // Trường hợp tháng 1
                if (month == 1)
                {
                    dataItemLastMonth = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == "12" && x.Year == (year - 1).ToString());
                }
                // Cung kì
                if (dataItemLastYear == null)
                {
                    dataItemLastYear = new ReportDetailtForPartner()
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
                    dataItemYear = new ReportDetailtForPartner()
                    {
                        PartnerCode = item.PartnerCode,
                        PartnerName = item.PartnerName,
                        Month = month.ToString(),
                        Year = year.ToString()
                    };
                }

                // Tháng trước
                // Trường hợp tháng 1
                if (month == 1)
                {
                    if (dataItemLastMonth == null)
                    {
                        dataItemLastMonth = new ReportDetailtForPartner()
                        {
                            PartnerCode = item.PartnerCode,
                            PartnerName = item.PartnerName,
                            Month = "12",
                            Year = (year - 1).ToString()
                        };
                    }
                }
                else
                {
                    if (dataItemLastMonth == null)
                    {
                        dataItemLastMonth = new ReportDetailtForPartner()
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
                    // So với tháng trước
                    double sumDSChiQuayYear = dataItemYear.DSChiQuay - dataItemLastMonth.DSChiQuay;
                    double sumDSChiNhaYear = dataItemYear.DSChiNha - dataItemLastMonth.DSChiNha;
                    double sumDSCKYear = dataItemYear.DSCK - dataItemLastMonth.DSCK;
                    double sumTongDSYear = dataItemYear.TongDS - dataItemLastMonth.TongDS;

                    // So với cùng kì năm trước
                    double sumDSChiQuayLastYear = dataItemYear.DSChiQuay - dataItemLastYear.DSChiQuay;
                    double sumDSChiNhaLastYear = dataItemYear.DSChiNha - dataItemLastYear.DSChiNha;
                    double sumDSCKLastYear = dataItemYear.DSCK - dataItemLastYear.DSCK;
                    double sumTongDSLastYear = dataItemYear.TongDS - dataItemLastYear.TongDS;

                    // add item vào table
                    table.Rows.Add(count++, item.PartnerName
                        , sumDSChiQuayYear, sumDSChiQuayLastYear
                        , sumDSChiNhaYear, sumDSChiNhaLastYear
                        , sumDSCKYear, sumDSCKLastYear
                        , sumTongDSYear, sumTongDSLastYear
                    );
                }
            }

            // Add dòng tổng
            DataRow row = table.NewRow();
            row["STT"] = "";
            row["PartnerName"] = "Tổng";
            row["CQ1"] = table.Compute("Sum(CQ1)", "");
            row["CQ2"] = table.Compute("Sum(CQ2)", "");

            row["CN1"] = table.Compute("Sum(CN1)", "");
            row["CN2"] = table.Compute("Sum(CN2)", "");

            row["CK1"] = table.Compute("Sum(CK1)", "");
            row["CK2"] = table.Compute("Sum(CK2)", "");

            row["TDS1"] = table.Compute("Sum(TDS1)", "");
            row["TDS2"] = table.Compute("Sum(TDS2)", "");
            table.Rows.Add(row);

            return Json(table.ToDataSourceResult(request), JsonRequestBehavior.AllowGet);
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
        /// Get data cho báo cáo chi tiết so sánh theo tháng
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult SearchReportDetailtCompareMonthForOne([DataSourceRequest]DataSourceRequest request, int month, int year, string reportTypeID, string partnerID)
        {
            List<ReportDetailtForPartner> listDataCompareMonth = new HSReportBL().ReportDetailtPartnerCompareMonthForOne(year, month, reportTypeID, partnerID);
            List<ReportDetailtForPartner> listDataCompareMonthConvert = new List<ReportDetailtForPartner>();

            foreach (ReportDetailtForPartner item in listDataCompareMonth)
            {
                item.PartnerName = string.Format("Tháng {0}/{1}", item.Month, item.Year);
                item.TongDS = item.DSChiQuay + item.DSChiNha + item.DSCK;
            }

            if (listDataCompareMonth.Count > 0)
            {
                // Month
                ReportDetailtForPartner dataItemMonth = listDataCompareMonth.Find(x => x.Month == month.ToString() && x.Year == year.ToString());
                // Last month
                ReportDetailtForPartner dataItemLastMonth = listDataCompareMonth.Find(x => x.Month == (month - 1).ToString() && x.Year == year.ToString());
                // month == 1
                if (month == 1)
                {
                    dataItemLastMonth = listDataCompareMonth.Find(x => x.Month == "12" && x.Year == (year - 1).ToString());
                }

                // Month last Year

                ReportDetailtForPartner dataItemMonthLastYear = listDataCompareMonth.Find(x => x.Month == month.ToString() && x.Year == (year - 1).ToString());

                // month
                if (dataItemMonth == null)
                {
                    dataItemMonth = new ReportDetailtForPartner()
                    {
                        Month = month.ToString(),
                        Year = year.ToString()
                    };
                }

                // last month
                if (dataItemLastMonth == null)
                {
                    if (month == 1)
                    {
                        dataItemLastMonth = new ReportDetailtForPartner()
                        {
                            Month = "12",
                            Year = (year - 1).ToString()
                        };
                    }
                    else
                    {
                        dataItemLastMonth = new ReportDetailtForPartner()
                        {
                            Month = (month - 1).ToString(),
                            Year = year.ToString()
                        };
                    }

                }

                // month last year
                if (dataItemMonthLastYear == null)
                {
                    dataItemMonthLastYear = new ReportDetailtForPartner()
                    {
                        Month = month.ToString(),
                        Year = (year - 1).ToString()
                    };
                }

                // Tháng hiện tại
                listDataCompareMonthConvert.Add(
                    new ReportDetailtForPartner()
                    {
                        PartnerName = dataItemMonth.PartnerName,
                        DSChiQuay = dataItemMonth.DSChiQuay,
                        DSChiNha = dataItemMonth.DSChiNha,
                        DSCK = dataItemMonth.DSCK,
                        TongDS = dataItemMonth.TongDS,
                    }
                );

                // Tháng trước
                listDataCompareMonthConvert.Add(
                    new ReportDetailtForPartner()
                    {
                        PartnerName = dataItemLastMonth.PartnerName,
                        DSChiQuay = dataItemLastMonth.DSChiQuay,
                        DSChiNha = dataItemLastMonth.DSChiNha,
                        DSCK = dataItemLastMonth.DSCK,
                        TongDS = dataItemLastMonth.TongDS,
                    }
                );

                // Cùng kì năm trước
                listDataCompareMonthConvert.Add(
                    new ReportDetailtForPartner()
                    {
                        PartnerName = dataItemMonthLastYear.PartnerName,
                        DSChiQuay = dataItemMonthLastYear.DSChiQuay,
                        DSChiNha = dataItemMonthLastYear.DSChiNha,
                        DSCK = dataItemMonthLastYear.DSCK,
                        TongDS = dataItemMonthLastYear.TongDS,
                    }
                );

                // Tăng giảm so với tháng trước (+/-)
                listDataCompareMonthConvert.Add(
                    new ReportDetailtForPartner()
                    {
                        PartnerName = "Tăng giảm so với tháng trước (+/-)",
                        DSChiQuay = dataItemMonth.DSChiQuay - dataItemLastMonth.DSChiQuay,
                        DSChiNha = dataItemMonth.DSChiNha - dataItemLastMonth.DSChiNha,
                        DSCK = dataItemMonth.DSCK - dataItemLastMonth.DSCK,
                        TongDS = dataItemMonth.TongDS - dataItemLastMonth.TongDS,
                    }
                );

                // Tăng giảm so với tháng trước (%)
                listDataCompareMonthConvert.Add(
                    new ReportDetailtForPartner()
                    {
                        PartnerName = "Tăng giảm so với tháng trước (%)",
                        DSChiQuay = dataItemLastMonth.DSChiQuay == 0 ? 0 : Math.Round(((dataItemMonth.DSChiQuay - dataItemLastMonth.DSChiQuay) / dataItemLastMonth.DSChiQuay) * 100, 2, MidpointRounding.ToEven),
                        DSChiNha = dataItemLastMonth.DSChiNha == 0 ? 0 : Math.Round(((dataItemMonth.DSChiNha - dataItemLastMonth.DSChiNha) / dataItemLastMonth.DSChiNha) * 100, 2, MidpointRounding.ToEven),
                        DSCK = dataItemLastMonth.DSCK == 0 ? 0 : Math.Round(((dataItemMonth.DSCK - dataItemLastMonth.DSCK) / dataItemLastMonth.DSCK) * 100, 2, MidpointRounding.ToEven),
                        TongDS = dataItemLastMonth.TongDS == 0 ? 0 : Math.Round(((dataItemMonth.TongDS - dataItemLastMonth.TongDS) / dataItemLastMonth.TongDS) * 100, 2, MidpointRounding.ToEven),
                    }
                );

                // Tăng giảm so với cùng kì năm trước (+/-)
                listDataCompareMonthConvert.Add(
                    new ReportDetailtForPartner()
                    {
                        PartnerName = "Tăng giảm so với cùng kì năm trước (+/-)",
                        DSChiQuay = dataItemMonth.DSChiQuay - dataItemMonthLastYear.DSChiQuay,
                        DSChiNha = dataItemMonth.DSChiNha - dataItemMonthLastYear.DSChiNha,
                        DSCK = dataItemMonth.DSCK - dataItemMonthLastYear.DSCK,
                        TongDS = dataItemMonth.TongDS - dataItemMonthLastYear.TongDS,
                    }
                );

                // Tăng giảm so với cùng kì năm trước (%)
                listDataCompareMonthConvert.Add(
                    new ReportDetailtForPartner()
                    {
                        PartnerName = "Tăng giảm so với cùng kì năm trước (%)",
                        DSChiQuay = dataItemMonthLastYear.DSChiQuay == 0 ? 0 : Math.Round(((dataItemMonth.DSChiQuay - dataItemMonthLastYear.DSChiQuay) / dataItemMonthLastYear.DSChiQuay) * 100, 2, MidpointRounding.ToEven),
                        DSChiNha = dataItemMonthLastYear.DSChiNha == 0 ? 0 : Math.Round(((dataItemMonth.DSChiNha - dataItemMonthLastYear.DSChiNha) / dataItemMonthLastYear.DSChiNha) * 100, 2, MidpointRounding.ToEven),
                        DSCK = dataItemMonthLastYear.DSCK == 0 ? 0 : Math.Round(((dataItemMonth.DSCK - dataItemMonthLastYear.DSCK) / dataItemMonthLastYear.DSCK) * 100, 2, MidpointRounding.ToEven),
                        TongDS = dataItemMonthLastYear.TongDS == 0 ? 0 : Math.Round(((dataItemMonth.TongDS - dataItemMonthLastYear.TongDS) / dataItemMonthLastYear.TongDS) * 100, 2, MidpointRounding.ToEven),
                    }
                );

                if (listDataCompareMonthConvert.Count > 0)
                {
                    listDataCompareMonth = new List<ReportDetailtForPartner>(listDataCompareMonthConvert);
                }
            }


            return Json(listDataCompareMonth.ToDataSourceResult(request), JsonRequestBehavior.AllowGet);
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
            List<ReportDetailtForPartner> listDataCompareMonth = new HSReportBL().ReportDetailtPartnerCompareMonthForOne(year, month, reportTypeID, partnerID);

            // # dòng record
            GradationCompare[] arrayGradation = null;

            if (listDataCompareMonth.Count.Equals(3))
            {
                // tạo mảng gồm 8 object
                arrayGradation = new GradationCompare[9];
                int count = 0;
                foreach (ReportDetailtForPartner item in listDataCompareMonth)
                {
                    // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                    arrayGradation[count] = new GradationCompare()
                    {
                        NameGradationCompare = string.Format("Tháng {0}/{1}", item.Month, item.Year),
                        amount = item.DSChiQuay,
                        NameType = "Hồ sơ \n chi quầy"
                    };

                    count++;
                    arrayGradation[count] = new GradationCompare()
                    {
                        NameGradationCompare = string.Format("Tháng {0}/{1}", item.Month, item.Year),
                        amount = item.DSChiNha,
                        NameType = "Hồ sơ \n chi nhà"
                    };

                    count++;
                    arrayGradation[count] = new GradationCompare()
                    {
                        NameGradationCompare = string.Format("Tháng {0}/{1}", item.Month, item.Year),
                        amount = item.DSCK,
                        NameType = "Hồ sơ \n chuyển khoản"
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
        /// Get data cho báo cáo chi tiết so sánh theo tháng
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult SearchReportDetailtCompareMonthForOnePercent([DataSourceRequest]DataSourceRequest request, int month, int year, string reportTypeID, string partnerID)
        {
            List<ReportDetailtForPartner> listDataCompareMonth = new HSReportBL().ReportDetailtPartnerCompareMonthForOne(year, month, reportTypeID, partnerID);
            List<ReportDetailtForPartner> listDataCompareMonthConvert = new List<ReportDetailtForPartner>();

            foreach (ReportDetailtForPartner item in listDataCompareMonth)
            {
                item.PartnerName = string.Format("Tháng {0}/{1}", item.Month, item.Year);
                item.TongDS = item.DSChiQuay + item.DSChiNha + item.DSCK;

                listDataCompareMonthConvert.Add(
                    new ReportDetailtForPartner()
                    {
                        PartnerName = string.Format("Tháng {0}/{1}", item.Month, item.Year),
                        DSChiQuay = item.TongDS == 0 ? 0 : Math.Round((item.DSChiQuay / item.TongDS) * 100, 2, MidpointRounding.ToEven),
                        DSChiNha = item.TongDS == 0 ? 0 : Math.Round((item.DSChiNha / item.TongDS) * 100, 2, MidpointRounding.ToEven),
                        DSCK = item.TongDS == 0 ? 0 : Math.Round((item.DSCK / item.TongDS) * 100, 2, MidpointRounding.ToEven),
                        TongDS = 100
                    }
                );
            }


            if (listDataCompareMonthConvert.Count().Equals(3))
            {
                // so với tháng trước
                double sumDSChiQuayYear = listDataCompareMonthConvert[2].DSChiQuay - listDataCompareMonthConvert[1].DSChiQuay;
                double sumDSChiNhaYear = listDataCompareMonthConvert[2].DSChiNha - listDataCompareMonthConvert[1].DSChiNha;
                double sumDSCKYear = listDataCompareMonthConvert[2].DSCK - listDataCompareMonthConvert[1].DSCK;
                double sumTongDSYear = listDataCompareMonthConvert[2].TongDS - listDataCompareMonthConvert[1].TongDS;

                // so với cùng kì
                double sumDSChiQuayLastYear = listDataCompareMonthConvert[2].DSChiQuay - listDataCompareMonthConvert[0].DSChiQuay;
                double sumDSChiNhaLastYear = listDataCompareMonthConvert[2].DSChiNha - listDataCompareMonthConvert[0].DSChiNha;
                double sumDSCKLastYear = listDataCompareMonthConvert[2].DSCK - listDataCompareMonthConvert[0].DSCK;
                double sumTongDSLastYear = listDataCompareMonthConvert[2].TongDS - listDataCompareMonthConvert[0].TongDS;

                // Tăng giảm so với tháng trước (%)
                listDataCompareMonthConvert.Add(
                    new ReportDetailtForPartner()
                    {
                        PartnerName = "Tăng giảm so với tháng trước (%)",
                        DSChiQuay = Math.Round(sumDSChiQuayYear, 2, MidpointRounding.ToEven),
                        DSChiNha = Math.Round(sumDSChiNhaYear, 2, MidpointRounding.ToEven),
                        DSCK = Math.Round(sumDSCKYear, 2, MidpointRounding.ToEven),
                        TongDS = 0
                    }
                );

                // Tăng giảm so với cùng kì năm trước (%)
                listDataCompareMonthConvert.Add(
                    new ReportDetailtForPartner()
                    {
                        PartnerName = "Tăng giảm so với cùng kì năm trước (%)",
                        DSChiQuay = Math.Round(sumDSChiQuayLastYear, 2, MidpointRounding.ToEven),
                        DSChiNha = Math.Round(sumDSChiNhaLastYear, 2, MidpointRounding.ToEven),
                        DSCK = Math.Round(sumDSCKLastYear, 2, MidpointRounding.ToEven),
                        TongDS = 0
                    }
                );
            }

            return Json(listDataCompareMonthConvert.ToDataSourceResult(request), JsonRequestBehavior.AllowGet);
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
            List<ReportDetailtForPartner> listDataCompareMonth = new HSReportBL().ReportDetailtPartnerCompareMonthForOne(year, month, reportTypeID, partnerID);
            List<ReportDetailtForPartner> listDataCompareMonthConvert = new List<ReportDetailtForPartner>();

            foreach (ReportDetailtForPartner item in listDataCompareMonth)
            {
                item.PartnerName = string.Format("Tháng {0}/{1}", item.Month, item.Year);
                item.TongDS = item.DSChiQuay + item.DSChiNha + item.DSCK;

                listDataCompareMonthConvert.Add(
                    new ReportDetailtForPartner()
                    {
                        PartnerName = string.Format("Tháng {0}/{1}", item.Month, item.Year),
                        DSChiQuay = item.TongDS == 0 ? 0 : Math.Round((item.DSChiQuay / item.TongDS) * 100, 2, MidpointRounding.ToEven),
                        DSChiNha = item.TongDS == 0 ? 0 : Math.Round((item.DSChiNha / item.TongDS) * 100, 2, MidpointRounding.ToEven),
                        DSCK = item.TongDS == 0 ? 0 : Math.Round((item.DSCK / item.TongDS) * 100, 2, MidpointRounding.ToEven),
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
            foreach (ReportDetailtForPartner item in listDataCompareMonthConvert)
            {
                if (item.Year == (year - 1).ToString() && item.Month == month.ToString())
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
        /// Get data cho báo cáo chi tiết so sánh theo tháng - tháng hiện tại
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult SearchCompareForMonthPieForLastMonth([DataSourceRequest]DataSourceRequest request, int month, int year, string reportTypeID, string partnerID)
        {
            List<ReportDetailtForPartner> listDataCompareMonth = new HSReportBL().ReportDetailtPartnerCompareMonthForOne(year, month, reportTypeID, partnerID);
            List<ReportDetailtForPartner> listDataCompareMonthConvert = new List<ReportDetailtForPartner>();

            foreach (ReportDetailtForPartner item in listDataCompareMonth)
            {
                item.PartnerName = string.Format("Tháng {0}/{1}", item.Month, item.Year);
                item.TongDS = item.DSChiQuay + item.DSChiNha + item.DSCK;

                listDataCompareMonthConvert.Add(
                    new ReportDetailtForPartner()
                    {
                        PartnerName = string.Format("Tháng {0}/{1}", item.Month, item.Year),
                        DSChiQuay = item.TongDS == 0 ? 0 : Math.Round((item.DSChiQuay / item.TongDS) * 100, 2, MidpointRounding.ToEven),
                        DSChiNha = item.TongDS == 0 ? 0 : Math.Round((item.DSChiNha / item.TongDS) * 100, 2, MidpointRounding.ToEven),
                        DSCK = item.TongDS == 0 ? 0 : Math.Round((item.DSCK / item.TongDS) * 100, 2, MidpointRounding.ToEven),
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
            int lastYear = 0;
            int lastMonth = 0;
            foreach (ReportDetailtForPartner item in listDataCompareMonthConvert)
            {
                if (month == 1)
                {
                    lastMonth = 12;
                    lastYear = year - 1;
                }
                else
                {
                    lastMonth = month - 1;
                    lastYear = year;
                }
                if (item.Year == lastYear.ToString() && item.Month == lastMonth.ToString())
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
        /// Get data cho báo cáo chi tiết so sánh theo tháng - tháng hiện tại
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult SearchCompareForMonthPieForYear([DataSourceRequest]DataSourceRequest request, int month, int year, string reportTypeID, string partnerID)
        {
            List<ReportDetailtForPartner> listDataCompareMonth = new HSReportBL().ReportDetailtPartnerCompareMonthForOne(year, month, reportTypeID, partnerID);
            List<ReportDetailtForPartner> listDataCompareMonthConvert = new List<ReportDetailtForPartner>();

            foreach (ReportDetailtForPartner item in listDataCompareMonth)
            {
                item.PartnerName = string.Format("Tháng {0}/{1}", item.Month, item.Year);
                item.TongDS = item.DSChiQuay + item.DSChiNha + item.DSCK;

                listDataCompareMonthConvert.Add(
                    new ReportDetailtForPartner()
                    {
                        PartnerName = string.Format("Tháng {0}/{1}", item.Month, item.Year),
                        DSChiQuay = item.TongDS == 0 ? 0 : Math.Round((item.DSChiQuay / item.TongDS) * 100, 2, MidpointRounding.ToEven),
                        DSChiNha = item.TongDS == 0 ? 0 : Math.Round((item.DSChiNha / item.TongDS) * 100, 2, MidpointRounding.ToEven),
                        DSCK = item.TongDS == 0 ? 0 : Math.Round((item.DSCK / item.TongDS) * 100, 2, MidpointRounding.ToEven),
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
            foreach (ReportDetailtForPartner item in listDataCompareMonthConvert)
            {
                if (item.Year == year.ToString() && item.Month == month.ToString())
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
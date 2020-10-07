using DongA.Bussiness;
using DongA.Entities;
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
    public class ReportHSTotalHSController : Controller
    {
        // GET: Admin/ReportHSTotalHS
        public ActionResult Index()
        {
            return View();
        }


        public ActionResult ReportDay(DateTime? fromDate, DateTime? toDate, string reportTypeID)
        {
            string nameUrl = "Hồ sơ/Tổng Hồ sơ chi trả/Chi tiết/Theo Ngày";
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
        /// Search báo cáo theo ngày của tổng doanh số chi trả
        /// </summary>
        /// <param name="request"></param>
        /// <param name="fromDate"></param>
        /// <param name="toDate"></param>
        /// <returns></returns>
        [HttpPost]
        public ActionResult SearchReportDay([DataSourceRequest]DataSourceRequest request, DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            List<ReportForTotalPayment> listData = new HSReportBL().SearchReportTPForDay(fromDate, toDate, reportTypeID);
            foreach (ReportForTotalPayment item in listData)
            {
                item.ReportID = string.Concat("Ngày ", item.CreatedDate.Day, "/", item.CreatedDate.Month, "/", item.CreatedDate.Year);
                item.Type = 0;
            }

            ReportForTotalPayment dataItem = new ReportForTotalPayment()
            {
                ReportID = "Tổng",
                Payed = listData.Sum(x => x.Payed),
                Type = 0
            };
            listData.Add(dataItem);

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
        public ActionResult SearchLineChartTotalPaymentForDay([DataSourceRequest]DataSourceRequest request, DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            List<ReportForTotalPayment> listData = new HSReportBL().SearchReportTPForDay(fromDate, toDate, reportTypeID);
            foreach (ReportForTotalPayment item in listData)
            {
                item.ReportID = string.Concat("Ngày ", item.CreatedDate.Day, "/", item.CreatedDate.Month, "/", item.CreatedDate.Year);
                item.Type = 0;
            }

            return Json(listData);
        }

        public ActionResult ReportMonth(DateTime? fromDate, DateTime? toDate, string reportTypeID)
        {
            string nameUrl = "Hồ sơ/Tổng Hồ sơ chi trả/Chi tiết/Theo Tháng";
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
        /// Search báo cáo theo ngày của tổng doanh số chi trả
        /// </summary>
        /// <param name="request"></param>
        /// <param name="fromDate"></param>
        /// <param name="toDate"></param>
        /// <returns></returns>
        [HttpPost]
        public ActionResult SearchReportMonth([DataSourceRequest]DataSourceRequest request, DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            List<ReportForTotalPayment> listData = new HSReportBL().SearchReportTPForMonth(fromDate, toDate, reportTypeID);
            foreach (ReportForTotalPayment item in listData)
            {
                item.ReportID = string.Format("Tháng {0}/{1}", item.Month, item.Year);
                item.Type = 0;
            }

            ReportForTotalPayment dataItem = new ReportForTotalPayment()
            {
                ReportID = "Tổng",
                Payed = listData.Sum(x => x.Payed),
                Type = 0
            };
            listData.Add(dataItem);

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
        public ActionResult SearchLineChartTotalPaymentForMonth([DataSourceRequest]DataSourceRequest request, DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            List<ReportForTotalPayment> listData = new HSReportBL().SearchReportTPForMonth(fromDate, toDate, reportTypeID);
            foreach (ReportForTotalPayment item in listData)
            {
                item.ReportID = string.Format("Tháng {0}/{1}", item.Month, item.Year);
                item.Type = 0;
            }

            return Json(listData);
        }

        public ActionResult ReportYear(DateTime? fromDate, DateTime? toDate, string reportTypeID)
        {
            string nameUrl = "Hồ sơ/Tổng Hồ sơ chi trả/Chi tiết/Theo Năm";
            ViewBag.NameURL = nameUrl;
            DataTable table = new DataTable();
            table.Columns.Add("ReportID", typeof(String));
            table.Columns.Add("year1", typeof(double));
            table.Columns.Add("year2", typeof(double));
            table.Columns.Add("year3", typeof(double));
            table.Columns.Add("year4", typeof(double));
            table.Columns.Add("year5", typeof(double));
            table.PrimaryKey = new DataColumn[] { table.Columns["ReportID"] };

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
            return View(table);
        }


        /// <summary>
        /// Search báo cáo theo ngày của tổng doanh số chi trả
        /// </summary>
        /// <param name="request"></param>
        /// <param name="fromDate"></param>
        /// <param name="toDate"></param>
        /// <returns></returns>
        [HttpPost]
        public ActionResult SearchReportYear([DataSourceRequest]DataSourceRequest request, DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            List<ReportForTotalPayment> listData = new HSReportBL().SearchReportTPForYear(fromDate, toDate, reportTypeID);
            foreach (ReportForTotalPayment item in listData)
            {
                item.ReportID = string.Format("Năm {0}", item.Year);
                item.Type = 0;
            }

            ReportForTotalPayment dataItem = new ReportForTotalPayment()
            {
                ReportID = "Tổng",
                Payed = listData.Sum(x => x.Payed),
                Type = 0
            };
            listData.Add(dataItem);

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
        public ActionResult SearchLineChartTotalPaymentForYear([DataSourceRequest]DataSourceRequest request, DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            List<ReportForTotalPayment> listData = new HSReportBL().SearchReportTPForYear(fromDate, toDate, reportTypeID);
            foreach (ReportForTotalPayment item in listData)
            {
                item.ReportID = string.Format("Năm {0}", item.Year);
                item.Type = 0;
            }

            return Json(listData);
        }

        public ActionResult ReportGradationCompare(int gradation, int year, string reportTypeID)
        {
            string nameUrl = "Hồ sơ/Tổng Hồ sơ chi trả/So sánh/Theo giai đoạn";
            ViewBag.NameURL = nameUrl;

            if (gradation > 0 && year > 0 && reportTypeID != null)
            {
                List<string> listData = new List<string>()
                {
                    gradation.ToString(),
                    year.ToString(),
                    reportTypeID
                };

                ViewData["listData"] = listData;
            }
            return View();
        }


        /// <summary>
        /// Search báo cáo theo ngày của thị trường
        /// </summary>
        /// <param name="request"></param>
        /// <param name="fromDay"></param>
        /// <param name="toDay"></param>
        /// <returns></returns>
        public ActionResult SearchReportGradationCompare([DataSourceRequest]DataSourceRequest request, int gradation, int year, string reportTypeID)
        {
            List<ReportForTotalPayment> listData = new HSReportBL().SearchReportTotalHSForGradationCompare(year, gradation, reportTypeID);

            string text = " tháng đầu năm ";
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

            if (listData.Count.Equals(2))
            {
                bool check = true;

                foreach (ReportForTotalPayment item in listData)
                {
                    item.ReportID = string.Concat("Lũy kế ", text, item.Year);
                    item.Type = 0;
                    // Set lại giá trị cho check để lấy giá trị của năm trước
                    check = false;
                }

                double totalPaymentPercent = listData[0].Payed - listData[1].Payed;

                // Object báo cáo tăng giảm so với cùng kỳ (%)
                ReportForTotalPayment dataDifferencePercent = new ReportForTotalPayment()
                {
                    ReportID = string.Format("Tăng giảm so với cùng kì {0} (%)", year - 1),
                    Payed = Math.Round(totalPaymentPercent / listData[1].Payed * 100, 2, MidpointRounding.ToEven),
                };

                listData.Add(dataDifferencePercent);

                // Object báo cáo tăng giảm so với cùng kỳ (+/-)
                ReportForTotalPayment dataDifference = new ReportForTotalPayment()
                {
                    ReportID = string.Format("Tăng giảm so với cùng kì {0} (+/-)", year - 1),
                    Payed = Math.Round(totalPaymentPercent, 2, MidpointRounding.ToEven),
                };
                listData.Add(dataDifference);
            }

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
        public ActionResult SearchColumnChartMaketReportForGradation([DataSourceRequest]DataSourceRequest request, int gradation, int year, string reportTypeID)
        {
            //string typeID = "1";
            List<ReportForTotalPayment> listData = new HSReportBL().SearchReportTotalHSForGradationCompare(year, gradation, reportTypeID);

            GradationCompare[] arrayGradation = null;

            string text = " tháng đầu năm ";
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

            if (listData.Count.Equals(2))
            {
                arrayGradation = new GradationCompare[2];
                int count = 0;
                foreach (ReportForTotalPayment item in listData)
                {
                    // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                    arrayGradation[count] = new GradationCompare()
                    {
                        NameGradationCompare = "Tổng hồ sơ",
                        amount = item.Payed,
                        NameType = string.Format("Lũy kề {0} {1}", text, item.Year)
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


        public ActionResult ReportCompareForMonth(int? month, int? year, string reportTypeID)
        {
            string nameUrl = "Hồ sơ/Tổng Hồ sơ chi trả/So sánh/Theo tháng";
            ViewBag.NameURL = nameUrl;

            if (month > 0 && year > 0 && reportTypeID != null)
            {
                List<string> listData = new List<string>()
                {
                    month.ToString(),
                    year.ToString(),
                    reportTypeID
                };

                ViewData["listData"] = listData;
            }
            return View();
        }


        /// <summary>
        /// Get data init cho màn hình so sánh theo tháng
        /// </summary>
        /// <param name="request"></param>
        /// <returns></returns>
        [HttpPost]
        public ActionResult SearchReportCompareForMonth([DataSourceRequest]DataSourceRequest request, int month, int year, string reportTypeID)
        {
            List<ReportForTotalPayment> listData = new HSReportBL().SearchReportTotalHSForCompareMonth(year, month, reportTypeID);

            if (listData.Count.Equals(3))
            {
                foreach (ReportForTotalPayment item in listData)
                {
                    item.ReportID = string.Format("Tháng {0}/{1}", item.Month, item.Year);
                }

                double totalPayment = listData[0].Payed - listData[1].Payed;
                double totalPaymentLastYear = listData[0].Payed - listData[2].Payed;

                // Object báo cáo tăng giảm so với tháng trước (%)
                ReportForTotalPayment dataDifference = null;
                dataDifference = new ReportForTotalPayment()
                {
                    ReportID = "Tăng giảm so với tháng trước (%)",
                    Payed = Math.Round(totalPayment / listData[1].Payed * 100, 2, MidpointRounding.ToEven),
                };

                listData.Add(dataDifference);

                // Object báo cáo tăng giảm so với tháng trước (+/-)
                dataDifference = new ReportForTotalPayment()
                {
                    ReportID = "Tăng giảm so với tháng trước (+/-)",
                    Payed = Math.Round(totalPayment, 2, MidpointRounding.ToEven),
                };

                listData.Add(dataDifference);

                // Object báo cáo tăng giảm so với cùng kì năm trước (%)
                dataDifference = new ReportForTotalPayment()
                {
                    ReportID = "Tăng giảm so với cùng kì năm trước (%)",
                    Payed = Math.Round(totalPaymentLastYear / listData[2].Payed * 100, 2, MidpointRounding.ToEven),
                };

                listData.Add(dataDifference);

                // Object báo cáo tăng giảm so với cùng kì năm trước (+/-)
                dataDifference = new ReportForTotalPayment()
                {
                    ReportID = "Tăng giảm so với cùng kì năm trước (+/-)",
                    Payed = Math.Round(totalPaymentLastYear, 2, MidpointRounding.ToEven),
                };

                listData.Add(dataDifference);
            }
            else
            {
                listData = new List<ReportForTotalPayment>();
            }

            return Json(listData.ToDataSourceResult(request), JsonRequestBehavior.AllowGet);
        }


        /// <summary>
        /// Get data cho việc vẽ biểu đồ cột cho so sánh giai đoạn
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult SearchColumnsChartCompareForMonth([DataSourceRequest]DataSourceRequest request, int month, int year, string reportTypeID)
        {
            List<ReportForTotalPayment> listData = new HSReportBL().SearchReportTotalHSForCompareMonth(year, month, reportTypeID);

            GradationCompare[] arrayGradation = null;
            if (listData.Count.Equals(3))
            {
                arrayGradation = new GradationCompare[3];
                int count = 0;
                foreach (ReportForTotalPayment item in listData)
                {
                    // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                    arrayGradation[count] = new GradationCompare()
                    {
                        NameGradationCompare = "Tổng hồ sơ",
                        amount = item.Payed,
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

    }
}
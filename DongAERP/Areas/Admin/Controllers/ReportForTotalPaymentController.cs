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
    public class ReportForTotalPaymentController : Controller
    {
        DateTime now = DateTime.Now;
        // GET: Admin/ReportForTotalPayment
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult ReportDay(DateTime? fromDay, DateTime? toDay, string reportTypeID)
        {
            string nameUrl = "Doanh số/Tổng doanh số chi trả/Chi tiết/Theo Ngày";
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

        public ActionResult ReportMonth(DateTime? fromDate, DateTime? toDate, string reportTypeID)
        {
            string nameUrl = "Doanh số/Tổng doanh số chi trả/Chi tiết/Theo Tháng";
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

        public ActionResult ReportYear(DateTime? fromDate, DateTime? toDate, string reportTypeID)
        {
            string nameUrl = "Doanh số/Tổng doanh số chi trả/Chi tiết/Theo Năm";
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

        public ActionResult ReportGradationCompare(int gradation, int year, string reportTypeID)
        {
            string nameUrl = "Doanh số/Tổng doanh số chi trả/So sánh/Theo giai đoạn";
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

        public ActionResult ReportCompareForMonth(int? month, int? year, string reportTypeID)
        {
            string nameUrl = "Doanh số/Tổng doanh số chi trả/So sánh/Theo tháng";
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
        /// Search báo cáo theo ngày của tổng doanh số chi trả
        /// </summary>
        /// <param name="request"></param>
        /// <param name="fromDay"></param>
        /// <param name="toDay"></param>
        /// <returns></returns>
        [HttpPost]
        public ActionResult ReportDay([DataSourceRequest]DataSourceRequest request, DateTime fromDay, DateTime toDay, string reportTypeID)
        {
            List<ReportForTotalPayment> listData = new ReportBL().SearchReportTPForDay(fromDay, toDay, reportTypeID);
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
        public ActionResult SearchLineChartTotalPaymentForReport([DataSourceRequest]DataSourceRequest request, DateTime fromDay, DateTime toDay, string reportTypeID)
        {
            List<ReportForTotalPayment> listData = new ReportBL().SearchReportTPForDay(fromDay, toDay, reportTypeID);
            foreach (ReportForTotalPayment item in listData)
            {
                item.ReportID = string.Concat("Ngày ", item.CreatedDate.Day, "/", item.CreatedDate.Month, "/", item.CreatedDate.Year);
                item.Type = 0;
            }

            return Json(listData);
        }

        ///// <summary>
        ///// Hiển thị bảng dữ liệu
        ///// </summary>
        ///// <param name="request"></param>
        ///// <returns></returns>
        //[HttpPost]
        //public ActionResult ReportMonth([DataSourceRequest]DataSourceRequest request, string reportTypeID)
        //{
        //    List<ReportForTotalPayment> listData = new ReportBL().DataReportTPForMonth(reportTypeID);
        //    foreach (ReportForTotalPayment item in listData)
        //    {
        //        item.ReportID = string.Concat(item.Month, "/", item.Year);
        //        item.Type = 0;
        //    }

        //    ReportForTotalPayment dataItem = new ReportForTotalPayment()
        //    {
        //        ReportID = "Tổng",
        //        Payed = listData.Sum(x => x.Payed)
        //    };

        //    listData.Add(dataItem);

        //    return Json(listData.ToDataSourceResult(request), JsonRequestBehavior.AllowGet);
        //}


        /// <summary>
        /// Search báo cáo theo ngày của thị trường
        /// </summary>
        /// <param name="request"></param>
        /// <param name="fromDay"></param>
        /// <param name="toDay"></param>
        /// <returns></returns>
        public ActionResult SearchReportTotalPaymentForMonth([DataSourceRequest]DataSourceRequest request, DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            List<ReportForTotalPayment> listData = new ReportBL().SearchReportTPForMonth(fromDate, toDate, reportTypeID);
            foreach (ReportForTotalPayment item in listData)
            {
                item.ReportID = string.Concat(item.Month, "/", item.Year);
                item.Type = 0;
            }

            ReportForTotalPayment dataItem = new ReportForTotalPayment()
            {
                ReportID = "Tổng",
                Payed = listData.Sum(x => x.Payed)
            };

            listData.Add(dataItem);

            return Json(listData.ToDataSourceResult(request), JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// Search report month theo ngày nhập vào
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult SearchLineChartReportTotalPaymentForMonth([DataSourceRequest]DataSourceRequest request, DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            List<ReportForTotalPayment> listData = new ReportBL().SearchReportTPForMonth(fromDate, toDate, reportTypeID);
            foreach (ReportForTotalPayment item in listData)
            {
                item.ReportID = string.Concat("Tháng ", item.Month, "/", item.Year);
                item.Type = 0;
            }
            return Json(listData);
        }

        ///// <summary>
        ///// Hiển thị bảng dữ liệu
        ///// </summary>
        ///// <param name="request"></param>
        ///// <returns></returns>
        //[HttpPost]
        //public ActionResult ReportYear([DataSourceRequest]DataSourceRequest request, string reportTypeID)
        //{
        //    List<ReportForTotalPayment> listData = new ReportBL().DataReportTPForYear(reportTypeID);
        //    // Khởi tạo datatable
        //    DataTable table = new DataTable();
        //    // Tạo các cột cho datatable
        //    table.Columns.Add("ReportID", typeof(String));
        //    table.Columns.Add("year1", typeof(double));
        //    table.Columns.Add("Year", typeof(string));

        //    if (listData.Count > 0)
        //    {
        //        // Tạo vòng for cho 12 tháng
        //        for (int i = 1; i <= 12; i++)
        //        {
        //            List<ReportForTotalPayment> listMonths = listData.Where(x => x.Month == i.ToString()).ToList();
        //            // số thứ tự theo năm
        //            int count = 1;
        //            DataRow dr = table.NewRow();
        //            dr["ReportID"] = string.Concat("Tháng ", i);
        //            foreach (ReportForTotalPayment item in listMonths)
        //            {
        //                dr[string.Concat("year", count)] = item.Payed;
        //                dr["Year"] = item.Year;
        //                count++;
        //            }
        //            table.Rows.Add(dr);
        //        }
        //        // Add row tổng
        //        DataRow drows = table.NewRow();
        //        drows["ReportID"] = "Tổng";
        //        object sum = table.Compute("SUM(year1)", string.Empty);
        //        drows["year1"] = Convert.ToDouble(string.IsNullOrEmpty(sum.ToString()) ? 0 : sum);
        //        table.Rows.Add(drows);
        //    }

        //    return Json(table.ToDataSourceResult(request), JsonRequestBehavior.AllowGet);
        //}

        /// <summary>
        /// Search báo cáo theo ngày của thị trường
        /// </summary>
        /// <param name="request"></param>
        /// <param name="fromDay"></param>
        /// <param name="toDay"></param>
        /// <returns></returns>
        public ActionResult SearchReportTotalPaymentForYear([DataSourceRequest]DataSourceRequest request, DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            List<ReportForTotalPayment> listData = new ReportBL().SearchReportTPForYear(fromDate, toDate, reportTypeID);
            // Số năm chênh lệch
            // Số 1 là số ngày cộng
            int countYear = toDate.Year - fromDate.Year + 1;
            // Khởi tạo datatable
            DataTable table = new DataTable();
            // Tạo các cột cho datatable
            table.Columns.Add("ReportID", typeof(String));
            for (int i = 1; i <= countYear; i++)
            {
                table.Columns.Add(string.Concat("year", i), typeof(double));
            }
            table.Columns.Add("Year", typeof(string));

            if (listData.Count > 0)
            {
                // Tạo vòng for cho 12 tháng
                for (int i = 1; i <= 12; i++)
                {
                    List<ReportForTotalPayment> listMonths = listData.Where(x => x.Month == i.ToString()).ToList();

                    List<string> listYear = new List<string>();

                    // số thứ tự theo năm
                    int count = 1;
                    DataRow dr = table.NewRow();
                    dr["ReportID"] = string.Concat("Tháng ", i);

                    foreach (ReportForTotalPayment item in listMonths)
                    {
                        dr[string.Concat("year", count)] = item.Payed;
                        listYear.Add(item.Year);
                        count++;
                    }
                    // add list year
                    dr["Year"] = string.Join("_", listYear).ToString();
                    table.Rows.Add(dr);
                }

                // add dòng tổng
                DataRow drows = table.NewRow();
                drows["ReportID"] = "Tổng";
                for (int i = 1; i <= countYear; i++)
                {
                    object sum = table.Compute(string.Format("SUM({0})", string.Concat("year", i)), string.Empty);
                    drows[string.Concat("year", i)] = Convert.ToDouble(string.IsNullOrEmpty(sum.ToString()) ? 0 : sum);
                }
                table.Rows.Add(drows);
            }

            return Json(table.ToDataSourceResult(request), JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// Line chart mặt định cho báo cáo khi vừa vào màn hình , get ngày hiện tại
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult LineChartTotalPaymentReportForYear(string reportTypeID)
        {
            List<ReportForTotalPayment> listData = new ReportBL().DataReportTPForYear(reportTypeID);
            GradationCompare[] arrayGradation = null;
            if (listData.Count > 0)
            {
                int countList = listData.Count;
                arrayGradation = new GradationCompare[countList];
                int count = 0;
                foreach (ReportForTotalPayment item in listData)
                {
                    // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                    arrayGradation[count] = new GradationCompare()
                    {
                        NameGradationCompare = item.Year,
                        amount = item.Payed,
                        NameType = string.Concat("Tháng ", item.Month)
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
        /// Search line char cho báo cáo năm của tổng báo cáo
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult SearchLineChartTotalPaymentReportForYear([DataSourceRequest]DataSourceRequest request, DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            List<ReportForTotalPayment> listData = new ReportBL().SearchReportTPForYear(fromDate, toDate, reportTypeID);

            GradationCompare[] arrayGradation = null;
            if (listData.Count > 0)
            {
                int countList = listData.Count;
                arrayGradation = new GradationCompare[countList];
                int count = 0;
                foreach (ReportForTotalPayment item in listData)
                {
                    // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                    arrayGradation[count] = new GradationCompare()
                    {
                        NameGradationCompare = item.Year,
                        amount = item.Payed,
                        NameType = string.Concat("Tháng ", item.Month)
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

        ///// <summary>
        ///// Get data cho báo cáo so sánh giai đoạn
        ///// </summary>
        ///// <returns></returns>
        ///// <history>
        /////     [Truong Lam]   Created [10/06/2020]
        ///// </history>
        //[HttpPost]
        //public ActionResult ReportGradationCompare([DataSourceRequest]DataSourceRequest request, string reportTypeID)
        //{
        //    // Danh sach của data gradation gồm key và value

        //    string[] ArrayData = { "1", "3 tháng đầu năm" };
        //    // typeID trong sql
        //    int typeID = 1;
        //    int toYear = DateTime.Today.Year;
        //    List<ReportForTotalPayment> listData = new ReportBL().DataReportTPForGradationCompare(toYear, typeID, reportTypeID);

        //    if (listData.Count.Equals(2))
        //    {
        //        bool check = true;

        //        foreach (ReportForTotalPayment item in listData)
        //        {
        //            item.ReportID = string.Concat("Lũy kế ", ArrayData[1], " ", toYear);
        //            if (!check)
        //            {
        //                item.ReportID = string.Concat("Lũy kế ", ArrayData[1], " ", toYear - 1);
        //            }
        //            item.Type = 0;
        //            // Set lại giá trị cho check để lấy giá trị của năm trước
        //            check = false;
        //        }

        //        double totalPaymentPercent = listData[0].Payed - listData[1].Payed;

        //        // Object báo cáo tăng giảm so với cùng kỳ (%)
        //        ReportForTotalPayment dataDifferencePercent = new ReportForTotalPayment()
        //        {
        //            ReportID = string.Format("Tăng giảm so với cùng kì {0} (%)", toYear - 1),
        //            Payed = Math.Round(totalPaymentPercent / listData[1].Payed * 100, 2, MidpointRounding.ToEven),
        //        };

        //        listData.Add(dataDifferencePercent);

        //        // Object báo cáo tăng giảm so với cùng kỳ (+/-)
        //        ReportForTotalPayment dataDifference = new ReportForTotalPayment()
        //        {
        //            ReportID = string.Format("Tăng giảm so với cùng kì {0} (+/-)", toYear - 1),
        //            Payed = Math.Round(totalPaymentPercent, 2, MidpointRounding.ToEven),
        //        };
        //        listData.Add(dataDifference);
        //    }

        //    return Json(listData.ToDataSourceResult(request), JsonRequestBehavior.AllowGet);
        //}

        /// <summary>
        /// Search báo cáo theo ngày của thị trường
        /// </summary>
        /// <param name="request"></param>
        /// <param name="fromDay"></param>
        /// <param name="toDay"></param>
        /// <returns></returns>
        public ActionResult SearchReportGradationCompare([DataSourceRequest]DataSourceRequest request, int gradation, int year, string reportTypeID)
        {
            List<ReportForTotalPayment> listData = new ReportBL().DataReportTPForGradationCompare(year, gradation, reportTypeID);

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
            List<ReportForTotalPayment> listData = new ReportBL().DataReportTPForGradationCompare(year, gradation, reportTypeID);

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
                        NameGradationCompare = string.Format("Lũy kế {0} {1}", text, item.Year),
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

        ///// <summary>
        ///// Get data cho báo cáo so sánh giai đoạn
        ///// </summary>
        ///// <returns></returns>
        ///// <history>
        /////     [Truong Lam]   Created [10/06/2020]
        ///// </history>
        //[HttpPost]
        //public ActionResult ReportCompareForMonth([DataSourceRequest]DataSourceRequest request, string reportTypeID)
        //{
        //    // Danh sach của data gradation gồm key và value

        //    string[] ArrayData = { "1", "3 tháng đầu năm" };
        //    List<ReportForTotalPayment> listData = new ReportBL().DataReportTPCompareForMonth(now.Year, now.Month, reportTypeID);

        //    if (listData.Count.Equals(3))
        //    {
        //        foreach (ReportForTotalPayment item in listData)
        //        {
        //            item.ReportID = string.Format("Tháng {0}/{1}", item.Month, item.Year);
        //        }

        //        double totalPayment = listData[0].Payed - listData[1].Payed;
        //        double totalPaymentLastYear = listData[0].Payed - listData[2].Payed;

        //        // Object báo cáo tăng giảm so với tháng trước (%)
        //        ReportForTotalPayment dataDifference = null;
        //        dataDifference = new ReportForTotalPayment()
        //        {
        //            ReportID = "Tăng giảm so với tháng trước (%)",
        //            Payed = Math.Round(totalPayment / listData[1].Payed * 100, 2, MidpointRounding.ToEven),
        //        };

        //        listData.Add(dataDifference);

        //        // Object báo cáo tăng giảm so với tháng trước (+/-)
        //        dataDifference = new ReportForTotalPayment()
        //        {
        //            ReportID = "Tăng giảm so với tháng trước (+/-)",
        //            Payed = Math.Round(totalPayment, 2, MidpointRounding.ToEven),
        //        };

        //        listData.Add(dataDifference);

        //        // Object báo cáo tăng giảm so với cùng kì năm trước (%)
        //        dataDifference = new ReportForTotalPayment()
        //        {
        //            ReportID = "Tăng giảm so với cùng kì năm trước (%)",
        //            Payed = Math.Round(totalPaymentLastYear / listData[2].Payed * 100, 2, MidpointRounding.ToEven),
        //        };

        //        listData.Add(dataDifference);

        //        // Object báo cáo tăng giảm so với cùng kì năm trước (+/-)
        //        dataDifference = new ReportForTotalPayment()
        //        {
        //            ReportID = "Tăng giảm so với cùng kì năm trước (+/-)",
        //            Payed = Math.Round(totalPaymentLastYear, 2, MidpointRounding.ToEven),
        //        };

        //        listData.Add(dataDifference);
        //    }
        //    else
        //    {
        //        listData = new List<ReportForTotalPayment>();
        //    }

        //    return Json(listData.ToDataSourceResult(request), JsonRequestBehavior.AllowGet);
        //}

        /// <summary>
        /// Get data cho việc vẽ biểu đồ cột cho so sánh giai đoạn
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult ColumnsChartCompareForMonth(string reportTypeID)
        {
            List<ReportForTotalPayment> listData = new ReportBL().DataReportTPCompareForMonth(now.Year, now.Month, reportTypeID);

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
                        NameGradationCompare = string.Format("Tháng {0}/{1}", item.Month, item.Year),
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

        /// <summary>
        /// Get data init cho màn hình so sánh theo tháng
        /// </summary>
        /// <param name="request"></param>
        /// <returns></returns>
        [HttpPost]
        public ActionResult SearchReportCompareForMonth([DataSourceRequest]DataSourceRequest request, int month, int year, string reportTypeID)
        {
            List<ReportForTotalPayment> listData = new ReportBL().DataReportTPCompareForMonth(year, month, reportTypeID);

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
            List<ReportForTotalPayment> listData = new ReportBL().DataReportTPCompareForMonth(year, month, reportTypeID);

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
                        NameGradationCompare = string.Format("Tháng {0}/{1}", item.Month, item.Year),
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
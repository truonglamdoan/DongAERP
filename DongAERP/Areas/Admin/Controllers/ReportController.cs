﻿using DongA.Bussiness;
using DongA.Entities;
using DongAERP.Areas.Admin.Models;
using Kendo.Mvc.Extensions;
using Kendo.Mvc.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace DongAERP.Areas.Admin.Controllers
{
    public class ReportController : Controller
    {

        DateTime now = DateTime.Now;
        // GET: Admin/Report
        /// <summary>
        /// ReportDay
        /// </summary>
        /// <returns></returns>
        public ActionResult Index()
        {
            return View();
        }

        /// <summary>
        /// Màn hình báo cáo cho ngày
        /// </summary>
        /// <returns></returns>
        public ActionResult ReportDay(DateTime? fromDay, DateTime? toDay, string reportTypeID)
        {
            string nameUrl = "Doanh số/Loại hình dịch vụ/Chi tiết/";
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
        public ActionResult ReportMonth(DateTime? fromDate, DateTime? toDate, string reportTypeID)
        {
            string nameUrl = "Doanh số/Loại hình dịch vụ/Chi tiết/Theo tháng";
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
        public ActionResult ReportYear(DateTime? fromDate, DateTime? toDate, string reportTypeID)
        {
            string nameUrl = "Doanh số/Loại hình dịch vụ/Chi tiết/Theo Năm";
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

        public ActionResult ReportGradationCompare(string gradation, int? year, string reportTypeID)
        {
            string nameUrl = "Doanh số/Loại hình dịch vụ/So sánh/Theo giai đoạn";
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

        public ActionResult ReportCompareForMonth(int? month, int? year, string reportTypeID)
        {
            string nameUrl = "Doanh số/Loại hình dịch vụ/So sánh/Theo tháng";
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
        /// Bảng hiển thị thông tin các giao dịch qua các ngày
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult ReportDay([DataSourceRequest]DataSourceRequest request, string reportTypeID)
        {
            List<Report> listData = new ReportBL().GetListReport(reportTypeID);
            foreach(Report item in listData)
            {
                item.ReportID = item.CreatedDate.ToString("dd/MM/yyyy");
                item.TongDS = item.DSChiNha + item.DSChiQuay + item.DSCK;
                item.Type = 0;
            }

            Report dataItem = new Report()
            {
                ReportID = "Tổng",
                DSChiQuay = listData.Sum(x => x.DSChiQuay),
                DSChiNha = listData.Sum(x => x.DSChiNha),
                DSCK = listData.Sum(x => x.DSCK),
                TongDS = listData.Sum(x => x.TongDS)
            };

            listData.Add(dataItem);

            return Json(listData.ToDataSourceResult(request), JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// Biểu đồ hiển thị thông tin các giao dịch qua các ngày
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>

        [HttpPost]
        public ActionResult LineChartReport([DataSourceRequest]DataSourceRequest request, string reportTypeID)
        {
            List<Report> listData = new ReportBL().GetListReport(reportTypeID);
            foreach (Report item in listData)
            {
                item.ReportID = item.CreatedDate.ToString("dd/MM/yyyy");
            }

            return Json(listData);
        }

        /// <summary>
        /// Search report day theo ngày nhập vào
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public ActionResult SearchReportDay([DataSourceRequest]DataSourceRequest request, DateTime fromDay, DateTime toDay, string reportTypeID)
        {
            List<Report> listData = new ReportBL().SearchDay(fromDay, toDay, reportTypeID);
            foreach (Report item in listData)
            {
                item.ReportID = item.CreatedDate.ToString("dd/MM/yyyy");
                item.TongDS = item.DSChiNha + item.DSChiQuay + item.DSCK;
                item.Type = 0;
            }

            Report dataItem = new Report()
            {
                ReportID = "Tổng",
                DSChiQuay = listData.Sum(x => x.DSChiQuay),
                DSChiNha = listData.Sum(x => x.DSChiNha),
                DSCK = listData.Sum(x => x.DSCK),
                TongDS = listData.Sum(x => x.TongDS)
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
        public ActionResult SearchLineChartReport([DataSourceRequest]DataSourceRequest request, DateTime fromDay, DateTime toDay, string reportTypeID)
        {
            List<Report> listData = new ReportBL().SearchDay(fromDay, toDay, reportTypeID);
            foreach (Report item in listData)
            {
                item.ReportID = item.CreatedDate.ToString("dd/MM/yyyy");
            }


            return Json(listData);
        }

        /// <summary>
        /// Bảng hiển thị thông tin các giao dịch qua các ngày
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult ReportMonth([DataSourceRequest]DataSourceRequest request, string reportTypeID)
        {
            List<Report> listData = new ReportBL().GetListReportMonth(now, reportTypeID);
            foreach (Report item in listData)
            {
                item.ReportID = string.Concat("Tháng ", item.Month, "/", item.Year);
                item.TongDS = item.DSChiNha + item.DSChiQuay + item.DSCK;
                item.Type = 0;
            }

            Report dataItem = new Report()
            {
                ReportID = "Tổng",
                DSChiQuay = listData.Sum(x => x.DSChiQuay),
                DSChiNha = listData.Sum(x => x.DSChiNha),
                DSCK = listData.Sum(x => x.DSCK),
                TongDS = listData.Sum(x => x.TongDS)
            };

            listData.Add(dataItem);

            return Json(listData.ToDataSourceResult(request), JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// Search report tháng theo ngày nhập vào
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult SearchReportMonth([DataSourceRequest]DataSourceRequest request, DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            List<Report> listData = new ReportBL().SearchMonth(fromDate, toDate, reportTypeID);
            foreach (Report item in listData)
            {
                item.ReportID = string.Concat("Tháng ", item.Month, "/", item.Year);
                item.TongDS = item.DSChiNha + item.DSChiQuay + item.DSCK;
                item.Type = 0;
            }

            Report dataItem = new Report()
            {
                ReportID = "Tổng",
                DSChiQuay = listData.Sum(x => x.DSChiQuay),
                DSChiNha = listData.Sum(x => x.DSChiNha),
                DSCK = listData.Sum(x => x.DSCK),
                TongDS = listData.Sum(x => x.TongDS)
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
        public ActionResult SearchLineChartReportMonth([DataSourceRequest]DataSourceRequest request, DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            List<Report> listData = new ReportBL().SearchMonth(fromDate, toDate, reportTypeID);
            foreach (Report item in listData)
            {
                item.ReportID = string.Concat("Tháng ", item.Month, "/", item.Year);
                item.TongDS = item.DSChiNha + item.DSChiQuay + item.DSCK;
            }
            return Json(listData);
        }
        
        /// <summary>
        /// Search report năm theo năm nhập vào
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult SearchReportYear([DataSourceRequest]DataSourceRequest request, DateTime fromYear, DateTime toYear, string reportTypeID)
        {
            List<Report> listData = new ReportBL().SearchYear(fromYear, toYear, reportTypeID);
            foreach (Report item in listData)
            {
                item.ReportID = item.Year;
                item.TongDS = item.DSChiNha + item.DSChiQuay + item.DSCK;
                item.Type = 0;
            }

            Report dataItem = new Report()
            {
                ReportID = "Tổng",
                DSChiQuay = listData.Sum(x => x.DSChiQuay),
                DSChiNha = listData.Sum(x => x.DSChiNha),
                DSCK = listData.Sum(x => x.DSCK),
                TongDS = listData.Sum(x => x.TongDS)
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
        public ActionResult SearchLineChartReportYear([DataSourceRequest]DataSourceRequest request, DateTime fromYear, DateTime toYear, string reportTypeID)
        {
            List<Report> listData = new ReportBL().SearchYear(fromYear, toYear, reportTypeID);
            foreach (Report item in listData)
            {
                item.Year = item.Year;
                item.TongDS = item.DSChiNha + item.DSChiQuay + item.DSCK;
            }
            return Json(listData);
        }
        
        /// <summary>
        /// Get data cho báo cáo so sánh giai đoạn
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult SearchGradationCompare([DataSourceRequest]DataSourceRequest request, string gradation, int? year, string reportTypeID)
        {
            // Việc chọn gradation xử lý sau
            // Set giá trị cho chọn báo cáo theo giai đoạn so sánh
            List<Report> listDataGradation = new ReportBL().ListDataGradationCompare(year.Value, int.Parse(gradation), reportTypeID);
            bool check = true;
            string text = " tháng đầu năm";
            switch (gradation)
            {
                case "1":
                    text = string.Concat("3", text);
                    break;
                case "2":
                    text = string.Concat("6", text);
                    break;
                case "3":
                    text = string.Concat("9", text);
                    break;
                default:
                    text = string.Concat("12", text);
                    break;
            }

            if (listDataGradation.Count.Equals(2))
            {
                foreach (Report item in listDataGradation)
                {
                    item.GradationID = string.Concat("Lũy kế ", text, " ", year);
                    if (!check)
                    {
                        item.GradationID = string.Concat("Lũy kế ", text, " ", year - 1);
                    }
                    item.TongDS = item.DSChiNha + item.DSChiQuay + item.DSCK;
                    item.Type = 0;
                    // Set lại giá trị cho check để lấy giá trị của năm trước
                    check = false;
                }

                // Object báo cáo tăng giảm so với cùng kỳ (+/-)
                Report dataDifference = new Report()
                {
                    GradationID = string.Format("Tăng giảm so với cùng kì {0} (+/-)", year - 1),
                    DSChiQuay = listDataGradation[1].DSChiQuay == 0 ? 0 : Math.Round(listDataGradation[0].DSChiQuay - listDataGradation[1].DSChiQuay, 2, MidpointRounding.ToEven),
                    DSChiNha = listDataGradation[1].DSChiNha == 0 ? 0 : Math.Round(listDataGradation[0].DSChiNha - listDataGradation[1].DSChiNha, 2, MidpointRounding.ToEven),
                    DSCK = listDataGradation[1].DSCK == 0 ? 0 : Math.Round(listDataGradation[0].DSCK - listDataGradation[1].DSCK, 2, MidpointRounding.ToEven),
                    TongDS = listDataGradation[1].TongDS == 0 ? 0 : Math.Round(listDataGradation[0].TongDS - listDataGradation[1].TongDS, 2, MidpointRounding.ToEven),
                };

                listDataGradation.Add(dataDifference);

                double dsChiQuayPercent = dataDifference.DSChiQuay / listDataGradation[1].DSChiQuay * 100;
                double dsChiNhaPercent = dataDifference.DSChiNha / listDataGradation[1].DSChiNha * 100;
                double dsChiCKPercent = dataDifference.DSCK / listDataGradation[1].DSCK * 100;
                double TongDSyPercent = dataDifference.TongDS / listDataGradation[1].TongDS * 100;
                // Object báo cáo tăng giảm so với cùng kỳ (%)
                Report dataDifferencePercent = new Report()
                {
                    GradationID = string.Format("Tăng giảm so với cùng kì {0} (%)", year - 1),
                    DSChiQuay = Math.Round(dsChiQuayPercent, 2, MidpointRounding.ToEven),
                    DSChiNha = Math.Round(dsChiNhaPercent, 2, MidpointRounding.ToEven),
                    DSCK = Math.Round(dsChiCKPercent, 2, MidpointRounding.ToEven),
                    TongDS = Math.Round(TongDSyPercent, 2, MidpointRounding.ToEven),
                };

                listDataGradation.Add(dataDifferencePercent);
            }
            else
            {
                listDataGradation = new List<Report>();
            }

            return Json(listDataGradation.ToDataSourceResult(request), JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// Get data cho việc vẽ biểu đồ cột cho so sánh giai đoạn
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult GetDataGradationCompare(string reportTypeID)
        {
            int nowYear = DateTime.Today.Year;
            int typeID = 1;
            List<Report> listData = new ReportBL().ListDataGradationCompare(nowYear, typeID, reportTypeID);
            int count = 0;
            int year = nowYear;

            // # dòng record
            GradationCompare[] arrayGradation = null;

            if (listData.Count == 0)
            {
                arrayGradation = new GradationCompare[1];
                arrayGradation[0] = new GradationCompare()
                {
                    NameGradationCompare = "1",
                    amount = 0,
                    NameType = ""

                };
            }
            else
            {
                // tạo mảng gồm 8 object
                arrayGradation = new GradationCompare[6];
            }

            foreach (Report item in listData)
            {
                // tổng doanh số
                item.TongDS = item.DSChiNha + item.DSChiQuay + item.DSCK;

                // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                arrayGradation[count] = new GradationCompare()
                {
                    NameGradationCompare = string.Concat("Lũy kế 3 tháng năm ", year),
                    amount = item.DSChiQuay,
                    NameType = "Doanh số \n chi quầy"
                };

                count++;
                arrayGradation[count] = new GradationCompare()
                {
                    NameGradationCompare = string.Concat("Lũy kế 3 tháng năm ", year),
                    amount = item.DSChiNha,
                    NameType = "Doanh số \n chi nhà"
                };

                count++;
                arrayGradation[count] = new GradationCompare()
                {
                    NameGradationCompare = string.Concat("Lũy kế 3 tháng năm ", year),
                    amount = item.DSCK,
                    NameType = "Doanh số \n chuyển khoản"
                };

                //count++;
                //arrayGradation[count] = new GradationCompare()
                //{
                //    NameGradationCompare = string.Concat("Lũy kế 3 tháng năm ", year),
                //    amount = item.TongDS,
                //    NameType = "Tổng doanh số"
                //};
                // Tăng count lên 1 đơn vị
                count++;
                year = year - 1;
            }
            
            return Json(arrayGradation);
        }
        

        /// <summary>
        /// Search báo cáo so sánh giai đoạn
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult SearchGradationComparePercentGrid([DataSourceRequest]DataSourceRequest request, string gradation, int? year, string reportTypeID)
        {
            List<Report> listDataGradation = new ReportBL().ListDataGradationComparePercent(int.Parse(gradation), year.Value, reportTypeID);
            bool check = true;

            // Tháng hiện tại và cùng kì năm ngoái
            if (listDataGradation.Count == 2)
            {
                string text = " tháng đầu năm";
                switch (gradation)
                {
                    case "1":
                        text = string.Concat("3", text);
                        break;
                    case "2":
                        text = string.Concat("6", text);
                        break;
                    case "3":
                        text = string.Concat("9", text);
                        break;
                    default:
                        text = string.Concat("12", text);
                        break;
                }

                foreach (Report item in listDataGradation)
                {
                    item.GradationID = string.Concat("Lũy kế ", text, " ", year);
                    if (!check)
                    {
                        item.GradationID = string.Concat("Lũy kế ", text, " ", year - 1);
                    }
                    item.TongDS = item.DSChiNha + item.DSChiQuay + item.DSCK;
                    item.Type = 0;
                    // Set lại giá trị cho check để lấy giá trị của năm trước
                    check = false;
                }

                // Object báo cáo tăng giảm so với cùng kỳ (%)
                double sumDSChiQuay = listDataGradation[0].DSChiQuay - listDataGradation[1].DSChiQuay;
                double sumDSChiNha = listDataGradation[0].DSChiNha - listDataGradation[1].DSChiNha;
                double sumDSCK = listDataGradation[0].DSCK - listDataGradation[1].DSCK;
                Report dataDifferencePercent = new Report()
                {
                    GradationID = string.Format("Tăng giảm so với cùng kì {0} (%)", year - 1),
                    DSChiQuay = listDataGradation[1].DSChiQuay == 0 ? 0 : Math.Round(sumDSChiQuay / listDataGradation[1].DSChiQuay * 100, 2, MidpointRounding.ToEven),
                    DSChiNha = listDataGradation[1].DSChiNha == 0 ? 0 : Math.Round(sumDSChiNha / listDataGradation[1].DSChiNha * 100, 2, MidpointRounding.ToEven),
                    DSCK = listDataGradation[1].DSCK == 0 ? 0 : Math.Round(sumDSCK / listDataGradation[1].DSCK * 100, 2, MidpointRounding.ToEven),
                };

                listDataGradation.Add(dataDifferencePercent);
            }
            else
            {
                listDataGradation = new List<Report>();
            }

            return Json(listDataGradation.ToDataSourceResult(request), JsonRequestBehavior.AllowGet);
        }
        
        /// <summary>
        /// Get data cho báo cáo so sánh giai đoạn
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult SearchDataGridCompareMonth([DataSourceRequest]DataSourceRequest request, int? month, int? year, string reportTypeID)
        {
            List<Report> listDataGradation = new ReportBL().ListDataMonthCompareGrid(year.Value, month.Value, reportTypeID);

            foreach (Report item in listDataGradation)
            {
                item.GradationID = string.Format("Tháng {0}/{1}", item.Month, item.Year);
                item.TongDS = item.DSChiQuay + item.DSChiNha + item.DSCK;
            }

            if (listDataGradation.Count.Equals(3))
            {
                // ds so với tháng trước
                double dsChiQuaySum = (listDataGradation[0].DSChiQuay - listDataGradation[1].DSChiQuay);
                double dsChiNhaSum = (listDataGradation[0].DSChiNha - listDataGradation[1].DSChiNha);
                double dsCKSum = (listDataGradation[0].DSCK - listDataGradation[1].DSCK);
                double tongDSSum = (listDataGradation[0].TongDS - listDataGradation[1].TongDS);

                // ds so với cùng kì năm trước
                double dsChiQuaySumlastYear = (listDataGradation[0].DSChiQuay - listDataGradation[2].DSChiQuay);
                double dsChiNhaSumlastYear = (listDataGradation[0].DSChiNha - listDataGradation[2].DSChiNha);
                double dsCKSumlastYear = (listDataGradation[0].DSCK - listDataGradation[2].DSCK);
                double tongDSSumlastYear = (listDataGradation[0].TongDS - listDataGradation[2].TongDS);

                List<Report> listReport = new List<Report>()
                {
                    // So sánh Tháng trước theo %
                    new Report()
                    {
                        GradationID = "Tăng giảm so với tháng trước (%)",
                        DSChiQuay = listDataGradation[1].DSChiQuay == 0 ? 0 : Math.Round(dsChiQuaySum / listDataGradation[1].DSChiQuay *100, 2, MidpointRounding.ToEven),
                        DSChiNha = listDataGradation[1].DSChiNha == 0 ? 0 : Math.Round(dsChiNhaSum / listDataGradation[1].DSChiNha *100, 2, MidpointRounding.ToEven),
                        DSCK = listDataGradation[1].DSCK == 0 ? 0 : Math.Round(dsCKSum / listDataGradation[1].DSCK *100, 2, MidpointRounding.ToEven),
                        TongDS = listDataGradation[1].TongDS == 0 ? 0 : Math.Round(tongDSSum / listDataGradation[1].TongDS *100, 2, MidpointRounding.ToEven)
                    },

                    // So sánh Tháng trước theo tăng giảm
                    new Report()
                    {
                        GradationID = "Tăng giảm so với tháng trước (+/-)",
                        DSChiQuay = Math.Round(dsChiQuaySum, 2, MidpointRounding.ToEven),
                        DSChiNha = Math.Round(dsChiNhaSum, 2, MidpointRounding.ToEven),
                        DSCK = Math.Round(dsCKSum, 2,MidpointRounding.ToEven),
                        TongDS = Math.Round(tongDSSum, 2,MidpointRounding.ToEven),
                    },

                    // So sánh cùng kỳ năm trước %
                    new Report()
                    {
                        GradationID = "Tăng giảm so với cùng kỳ năm trước (%)",
                        DSChiQuay = listDataGradation[2].DSChiQuay == 0 ? 0 : Math.Round(dsChiQuaySumlastYear / listDataGradation[2].DSChiQuay *100, 2, MidpointRounding.ToEven),
                        DSChiNha = listDataGradation[2].DSChiNha == 0 ? 0 : Math.Round(dsChiNhaSumlastYear / listDataGradation[2].DSChiNha *100, 2, MidpointRounding.ToEven),
                        DSCK = listDataGradation[2].DSCK == 0 ? 0 : Math.Round(dsCKSumlastYear / listDataGradation[2].DSCK *100, 2, MidpointRounding.ToEven),
                        TongDS = listDataGradation[2].TongDS == 0 ? 0 : Math.Round(tongDSSumlastYear / listDataGradation[2].TongDS *100, 2, MidpointRounding.ToEven)
                    },

                    // So sánh cùng kỳ năm trước theo tăng giảm
                    new Report()
                    {
                        GradationID = "Tăng giảm so với cùng kỳ năm trước (+/-)",
                        DSChiQuay = Math.Round(dsChiQuaySumlastYear, 2, MidpointRounding.ToEven),
                        DSChiNha = Math.Round(dsChiNhaSumlastYear, 2, MidpointRounding.ToEven),
                        DSCK = Math.Round(dsCKSumlastYear, 2,MidpointRounding.ToEven),
                        TongDS = Math.Round(tongDSSumlastYear, 2,MidpointRounding.ToEven),
                    },
                };

                // Merger
                listDataGradation.AddRange(listReport);
            }
            else
            {
                listDataGradation = new List<Report>();
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
        public ActionResult SearchDataChartCompareMonth([DataSourceRequest]DataSourceRequest request, int? month, int? year, string reportTypeID)
        {
            List<Report> listDataGradation = new ReportBL().ListDataMonthCompareGrid(year.Value, month.Value, reportTypeID);

            // # dòng record
            GradationCompare[] arrayGradation = null;

            if (listDataGradation.Count.Equals(3))
            {
                // tạo mảng gồm 8 object
                arrayGradation = new GradationCompare[9];
                int count = 0;
                foreach (Report item in listDataGradation)
                {
                    // tổng doanh số
                    item.TongDS = item.DSChiNha + item.DSChiQuay + item.DSCK;

                    // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                    arrayGradation[count] = new GradationCompare()
                    {
                        NameGradationCompare = string.Format("Tháng {0}/{1}", item.Month, item.Year),
                        amount = item.DSChiQuay,
                        NameType = "Doanh số \n chi quầy"
                    };

                    count++;
                    arrayGradation[count] = new GradationCompare()
                    {
                        NameGradationCompare = string.Format("Tháng {0}/{1}", item.Month, item.Year),
                        amount = item.DSChiNha,
                        NameType = "Doanh số \n chi nhà"
                    };

                    count++;
                    arrayGradation[count] = new GradationCompare()
                    {
                        NameGradationCompare = string.Format("Tháng {0}/{1}", item.Month, item.Year),
                        amount = item.DSCK,
                        NameType = "Doanh số \n chuyển khoản"
                    };

                    //count++;
                    //arrayGradation[count] = new GradationCompare()
                    //{
                    //    NameGradationCompare = string.Format("Tháng {0}/{1}", item.Month, item.Year),
                    //    amount = item.TongDS,
                    //    NameType = "Tổng doanh số"
                    //};
                    // Tăng count lên 1 đơn vị
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
        /// Get data cho việc vẽ biểu đồ cột cho so sánh theo tháng và cùng kì năm trước
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult SearchGridMonthCompareProportion([DataSourceRequest]DataSourceRequest request, int? month, int? year, string reportTypeID)
        {
            List<Report> listDataGradation = new ReportBL().ListDataLastMonthCompareProportionPercent(year.Value, month.Value, reportTypeID);

            foreach (Report item in listDataGradation)
            {
                item.GradationID = string.Format("Tháng {0}/{1}", item.Month, item.Year);
                item.TongDS = item.DSChiQuay + item.DSChiNha + item.DSCK;
            }

            // Phải có đủ dữ liệu tháng hiện tại, tháng trước và cùng kì năm trước
            if (listDataGradation.Count == 3)
            {
                // ds so với tháng trước
                double dsChiQuaySum = (listDataGradation[0].DSChiQuay - listDataGradation[1].DSChiQuay);
                double dsChiNhaSum = (listDataGradation[0].DSChiNha - listDataGradation[1].DSChiNha);
                double dsCKSum = (listDataGradation[0].DSCK - listDataGradation[1].DSCK);
                double tongDSSum = (listDataGradation[0].TongDS - listDataGradation[1].TongDS);

                // ds so với cùng kì năm trước
                double dsChiQuaySumlastYear = (listDataGradation[0].DSChiQuay - listDataGradation[2].DSChiQuay);
                double dsChiNhaSumlastYear = (listDataGradation[0].DSChiNha - listDataGradation[2].DSChiNha);
                double dsCKSumlastYear = (listDataGradation[0].DSCK - listDataGradation[2].DSCK);
                double tongDSSumlastYear = (listDataGradation[0].TongDS - listDataGradation[2].TongDS);

                List<Report> listReport = new List<Report>()
                {
                    // So sánh Tháng trước theo %
                    new Report()
                    {
                        GradationID = "Tăng giảm so với tháng trước (%)",
                        DSChiQuay = listDataGradation[1].DSChiQuay == 0 ? 0 : Math.Round(dsChiQuaySum / listDataGradation[1].DSChiQuay *100, 2, MidpointRounding.ToEven),
                        DSChiNha = listDataGradation[1].DSChiNha == 0 ? 0 : Math.Round(dsChiNhaSum / listDataGradation[1].DSChiNha *100, 2, MidpointRounding.ToEven),
                        DSCK = listDataGradation[1].DSCK == 0 ? 0 : Math.Round(dsCKSum / listDataGradation[1].DSCK *100, 2, MidpointRounding.ToEven),
                        TongDS = listDataGradation[1].TongDS == 0 ? 0 : Math.Round(tongDSSum / listDataGradation[1].TongDS *100, 2, MidpointRounding.ToEven)
                    },

                    // So sánh cùng kỳ năm trước %
                    new Report()
                    {
                        GradationID = "Tăng giảm so với cùng kỳ năm trước (%)",
                        DSChiQuay = listDataGradation[2].DSChiQuay == 0 ? 0 : Math.Round(dsChiQuaySumlastYear / listDataGradation[2].DSChiQuay *100, 2, MidpointRounding.ToEven),
                        DSChiNha = listDataGradation[2].DSChiNha == 0 ? 0 :  Math.Round(dsChiNhaSumlastYear / listDataGradation[2].DSChiNha *100, 2, MidpointRounding.ToEven),
                        DSCK = listDataGradation[2].DSCK == 0 ? 0 :  Math.Round(dsCKSumlastYear / listDataGradation[2].DSCK *100, 2, MidpointRounding.ToEven),
                        TongDS = listDataGradation[2].TongDS == 0 ? 0 :  Math.Round(tongDSSumlastYear / listDataGradation[2].TongDS *100, 2, MidpointRounding.ToEven)
                    }
                };

                // Merger
                listDataGradation.AddRange(listReport);
            }
            else
            {
                listDataGradation = new List<Report>();
            }

            return Json(listDataGradation.ToDataSourceResult(request), JsonRequestBehavior.AllowGet);
        }
        

        /// <summary>
        /// search data cho biểu đồ của tháng hiện tại
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult SearchPieMonthCompareProportion([DataSourceRequest]DataSourceRequest request, int? month, int? year, string reportTypeID)
        {
            List<Report> listDataGradation = new ReportBL().ListDataLastMonthCompareProportionPercent(year.Value, month.Value, reportTypeID);
            // tháng hiện tại
            Report listDataGradationMonth = listDataGradation.Find(x => x.Month == month.ToString() && x.Year == year.ToString());
            // # dòng record
            GradationChartPie[] arrayGradation = null;

            int count = 0;
            if (listDataGradationMonth != null)
            {
                arrayGradation = new GradationChartPie[3];
                // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                arrayGradation[count] = new GradationChartPie()
                {
                    category = "Doanh số \n chi quầy",
                    value = listDataGradationMonth.DSChiQuay,
                    color = "#FFBF00"
                };

                count++;
                arrayGradation[count] = new GradationChartPie()
                {
                    category = "Doanh số \n chi nhà",
                    value = listDataGradationMonth.DSChiNha,
                    color = "#40FF00"
                };

                count++;
                arrayGradation[count] = new GradationChartPie()
                {
                    category = "Doanh số \n chuyển khoản",
                    value = listDataGradationMonth.DSCK,
                    color = "#2ECCFA"
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
        /// search data cho biểu đồ của tháng trước
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult SearchPieLastMonthCompareProportion([DataSourceRequest]DataSourceRequest request, int? month, int? year, string reportTypeID)
        {
            List<Report> listDataGradation = new ReportBL().ListDataLastMonthCompareProportionPercent(year.Value, month.Value, reportTypeID);
            // tháng trước
            Report listDataGradationMonth = null;

            // Trường hợp là tháng 1
            if (month == 1)
            {
                listDataGradationMonth = listDataGradation.Find(x => x.Month == "12" && x.Year == (year -1).ToString());
            }
            else
            {
                listDataGradationMonth = listDataGradation.Find(x => x.Month == (month - 1).ToString() && x.Year == year.ToString());
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
                    category = "Doanh số \n chi quầy",
                    value = listDataGradationMonth.DSChiQuay,
                    color = "#FFBF00"
                };

                count++;
                arrayGradation[count] = new GradationChartPie()
                {
                    category = "Doanh số \n chi nhà",
                    value = listDataGradationMonth.DSChiNha,
                    color = "#40FF00"
                };

                count++;
                arrayGradation[count] = new GradationChartPie()
                {
                    category = "Doanh số \n chuyển khoản",
                    value = listDataGradationMonth.DSCK,
                    color = "#2ECCFA"
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
        /// search data cho biểu đồ của tháng trước
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult SearchPieMonthLastYearCompareProportion([DataSourceRequest]DataSourceRequest request, int? month, int? year, string reportTypeID)
        {
            List<Report> listDataGradation = new ReportBL().ListDataLastMonthCompareProportionPercent(year.Value, month.Value, reportTypeID);
            // Cùng kì năm trước
            Report listDataGradationMonth = listDataGradation.Find(x => x.Month == month.ToString() && x.Year == (year - 1).ToString());
            // # dòng record
            GradationChartPie[] arrayGradation = null;

            int count = 0;
            if (listDataGradationMonth != null)
            {
                arrayGradation = new GradationChartPie[3];
                // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                arrayGradation[count] = new GradationChartPie()
                {
                    category = "Doanh số \n chi quầy",
                    value = listDataGradationMonth.DSChiQuay,
                    color = "#FFBF00"
                };

                count++;
                arrayGradation[count] = new GradationChartPie()
                {
                    category = "Doanh số \n chi nhà",
                    value = listDataGradationMonth.DSChiNha,
                    color = "#40FF00"
                };

                count++;
                arrayGradation[count] = new GradationChartPie()
                {
                    category = "Doanh số \n chuyển khoản",
                    value = listDataGradationMonth.DSCK,
                    color = "#2ECCFA"
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
        /// Get data cho việc vẽ biểu đồ cột cho so sánh theo tháng và cùng kì năm trước
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult SearchDataChartCompare([DataSourceRequest]DataSourceRequest request, string gradation, int? year, string reportTypeID)
        {
            List<Report> listData = new ReportBL().ListDataGradationCompare(year.Value, int.Parse(gradation), reportTypeID);
            GradationCompare[] arrayGradation = null;

            if (listData.Any(x => x.Year == year.ToString()) && listData.Any(x => x.Year == (year - 1).ToString()))
            {

                arrayGradation = new GradationCompare[6];
                string text = " tháng năm";
                switch (gradation)
                {
                    case "1":
                        text = string.Concat("3", text);
                        break;
                    case "2":
                        text = string.Concat("6", text);
                        break;
                    case "3":
                        text = string.Concat("9", text);
                        break;
                    default:
                        text = string.Concat("12", text);
                        break;
                }

                int count = 0;
                foreach (Report item in listData)
                {
                    // tổng doanh số
                    item.TongDS = item.DSChiNha + item.DSChiQuay + item.DSCK;

                    // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                    arrayGradation[count] = new GradationCompare()
                    {
                        NameGradationCompare = string.Format(" Lũy kế {0} {1}", text, year),
                        amount = item.DSChiQuay,
                        NameType = "Doanh số \n chi quầy"
                    };

                    count++;
                    arrayGradation[count] = new GradationCompare()
                    {
                        NameGradationCompare = string.Format(" Lũy kế {0} {1}", text, year),
                        amount = item.DSChiNha,
                        NameType = "Doanh số \n chi nhà"
                    };

                    count++;
                    arrayGradation[count] = new GradationCompare()
                    {
                        NameGradationCompare = string.Format(" Lũy kế {0} {1}", text, year),
                        amount = item.DSCK,
                        NameType = "Doanh số \n chuyển khoản"
                    };

                    //count++;
                    //arrayGradation[count] = new GradationCompare()
                    //{
                    //    NameGradationCompare = string.Format(" Lũy kế {0} {1}", text, year),
                    //    amount = item.TongDS,
                    //    NameType = "Tổng doanh số"
                    //};
                    // Tăng count lên 1 đơn vị
                    count++;
                    year = year - 1;
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
        /// search data cho biểu đồ của tháng hiện tại
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult SearchPieCompareProportionYear([DataSourceRequest]DataSourceRequest request, string gradation, int? year, string reportTypeID)
        {
            List<Report> listDataGradation = new ReportBL().ListDataGradationComparePercent(int.Parse(gradation), year.Value, reportTypeID);

            GradationChartPie[] arrayGradation = null;
            if (listDataGradation.Find(x =>x.Year == year.ToString()) != null)
            {
                arrayGradation = new GradationChartPie[3];
                int count = 0;
                foreach (Report item in listDataGradation)
                {
                    if (item.Year.Equals(year.ToString()))
                    {
                        // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                        arrayGradation[count] = new GradationChartPie()
                        {
                            category = "Doanh số \n chi quầy",
                            value = item.DSChiQuay,
                            color = "#FFBF00"
                        };

                        count++;
                        arrayGradation[count] = new GradationChartPie()
                        {
                            category = "Doanh số \n chi nhà",
                            value = item.DSChiNha,
                            color = "#40FF00"
                        };

                        count++;
                        arrayGradation[count] = new GradationChartPie()
                        {
                            category = "Doanh số \n chuyển khoản",
                            value = item.DSCK,
                            color = "#2ECCFA"
                        };
                        break;
                    }
                }
            }
            else
            {
                arrayGradation = new GradationChartPie[1];
                arrayGradation[0] = new GradationChartPie()
                {
                    category = "",
                    value = 0,
                    color = "#9de219"

                };
            }

            return Json(arrayGradation);
        }

        /// <summary>
        /// search data cho biểu đồ của tháng hiện tại
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult SearchPieCompareProportionLastYear([DataSourceRequest]DataSourceRequest request, string gradation, int? year, string reportTypeID)
        {
            int lastYear = year.Value - 1;
            List<Report> listDataGradation = new ReportBL().ListDataGradationComparePercent(int.Parse(gradation), lastYear, reportTypeID);

            GradationChartPie[] arrayGradation = null;
            if (listDataGradation.Find(x => x.Year == lastYear.ToString()) != null)
            {
                arrayGradation = new GradationChartPie[3];
                int count = 0;
                foreach (Report item in listDataGradation)
                {
                    if (item.Year.Equals(lastYear.ToString()))
                    {
                        // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                        arrayGradation[count] = new GradationChartPie()
                        {
                            category = "Doanh số \n chi quầy",
                            value = item.DSChiQuay,
                            color = "#FFBF00"
                        };

                        count++;
                        arrayGradation[count] = new GradationChartPie()
                        {
                            category = "Doanh số \n chi nhà",
                            value = item.DSChiNha,
                            color = "#40FF00"
                        };

                        count++;
                        arrayGradation[count] = new GradationChartPie()
                        {
                            category = "Doanh số \n chuyển khoản",
                            value = item.DSCK,
                            color = "#2ECCFA"
                        };
                        break;
                    }
                }
            }
            else
            {
                arrayGradation = new GradationChartPie[1];
                arrayGradation[0] = new GradationChartPie()
                {
                    category = "",
                    value = 0,
                    color = "#9de219"

                };
            }

            return Json(arrayGradation);
        }
    }
}
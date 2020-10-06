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
    public class ReportHSMoneyTypeController : Controller
    {
        // GET: Admin/ReportHSMoneyType
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult ReportDay(DateTime? fromDate, DateTime? toDate, string reportTypeID)
        {
            string nameUrl = "Hồ sơ/Loại tiền/Chi tiết/Theo Ngày";
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
        /// Search báo cáo theo ngày của thị trường
        /// </summary>
        /// <param name="request"></param>
        /// <param name="fromDate"></param>
        /// <param name="toDay"></param>
        /// <returns></returns>
        public ActionResult SearchReportTotalMoneyTypeForDay([DataSourceRequest]DataSourceRequest request, DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            List<ReportForTotalMoneyType> listData = new HSReportBL().SearchReportHSTotalMoneyTypeForDay(fromDate, toDate, reportTypeID);
            foreach (ReportForTotalMoneyType item in listData)
            {
                double sumTong = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;
                item.ReportID = item.CreatedDate.ToString("dd/MM/yyyy");
                item.TongDS = Math.Round(sumTong, 2, MidpointRounding.ToEven);
                item.Type = 0;
            }


            // Add dòng tổng
            ReportForTotalMoneyType dataItem = new ReportForTotalMoneyType()
            {
                ReportID = "Tổng",
                VND = listData.Sum(x => x.VND),
                USD = listData.Sum(x => x.USD),
                EUR = listData.Sum(x => x.EUR),
                CAD = listData.Sum(x => x.CAD),
                AUD = listData.Sum(x => x.AUD),
                GBP = listData.Sum(x => x.GBP),
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
        public ActionResult SearchLineChartTotalMoneyTypeForDay([DataSourceRequest]DataSourceRequest request, DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            List<ReportForTotalMoneyType> listData = new HSReportBL().SearchReportHSTotalMoneyTypeForDay(fromDate, toDate, reportTypeID);
            foreach (ReportForTotalMoneyType item in listData)
            {
                double sumTong = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;
                item.ReportID = item.CreatedDate.ToString("dd/MM/yyyy");
                item.TongDS = Math.Round(sumTong, 2, MidpointRounding.ToEven);
                item.Type = 0;
            }

            return Json(listData);
        }


        public ActionResult ReportMonth(DateTime? fromDate, DateTime? toDate, string reportTypeID)
        {
            string nameUrl = "Doanh số/Loại tiền/Chi tiết/Theo Tháng";
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
        /// Search báo cáo theo ngày của thị trường
        /// </summary>
        /// <param name="request"></param>
        /// <param name="fromDate"></param>
        /// <param name="toDay"></param>
        /// <returns></returns>
        public ActionResult SearchReportTotalMoneyTypeForMonth([DataSourceRequest]DataSourceRequest request, DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            List<ReportForTotalMoneyType> listData = new HSReportBL().SearchReportHSTotalMoneyTypeForMonth(fromDate, toDate, reportTypeID);
            foreach (ReportForTotalMoneyType item in listData)
            {
                double sumTong = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;
                item.ReportID = string.Format("Tháng {0}/{1}", item.Month, item.Year);
                item.TongDS = Math.Round(sumTong, 2, MidpointRounding.ToEven);
                item.Type = 0;
            }


            // Add dòng tổng
            ReportForTotalMoneyType dataItem = new ReportForTotalMoneyType()
            {
                ReportID = "Tổng",
                VND = listData.Sum(x => x.VND),
                USD = listData.Sum(x => x.USD),
                EUR = listData.Sum(x => x.EUR),
                CAD = listData.Sum(x => x.CAD),
                AUD = listData.Sum(x => x.AUD),
                GBP = listData.Sum(x => x.GBP),
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
        public ActionResult SearchLineChartTotalMoneyTypeForMonth([DataSourceRequest]DataSourceRequest request, DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            List<ReportForTotalMoneyType> listData = new HSReportBL().SearchReportHSTotalMoneyTypeForMonth(fromDate, toDate, reportTypeID);
            foreach (ReportForTotalMoneyType item in listData)
            {
                double sumTong = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;
                item.ReportID = string.Format("Tháng {0}/{1}", item.Month, item.Year);
                item.TongDS = Math.Round(sumTong, 2, MidpointRounding.ToEven);
                item.Type = 0;
            }

            return Json(listData);
        }


        public ActionResult ReportYear(DateTime? fromDate, DateTime? toDate, string reportTypeID)
        {
            string nameUrl = "Hồ sơ/Loại tiền/Chi tiết/Theo Năm";
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
        /// Search báo cáo theo ngày của thị trường
        /// </summary>
        /// <param name="request"></param>
        /// <param name="fromDate"></param>
        /// <param name="toDay"></param>
        /// <returns></returns>
        public ActionResult SearchReportTotalMoneyTypeForYear([DataSourceRequest]DataSourceRequest request, DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            List<ReportForTotalMoneyType> listData = new HSReportBL().SearchReportHSTotalMoneyTypeForYear(fromDate, toDate, reportTypeID);
            foreach (ReportForTotalMoneyType item in listData)
            {
                double sumTong = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;
                item.ReportID = string.Format("Năm {0}", item.Year);
                item.TongDS = Math.Round(sumTong, 2, MidpointRounding.ToEven);
                item.Type = 0;
            }


            // Add dòng tổng
            ReportForTotalMoneyType dataItem = new ReportForTotalMoneyType()
            {
                ReportID = "Tổng",
                VND = listData.Sum(x => x.VND),
                USD = listData.Sum(x => x.USD),
                EUR = listData.Sum(x => x.EUR),
                CAD = listData.Sum(x => x.CAD),
                AUD = listData.Sum(x => x.AUD),
                GBP = listData.Sum(x => x.GBP),
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
        public ActionResult SearchLineChartTotalMoneyTypeForYear([DataSourceRequest]DataSourceRequest request, DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            List<ReportForTotalMoneyType> listData = new HSReportBL().SearchReportHSTotalMoneyTypeForYear(fromDate, toDate, reportTypeID);
            foreach (ReportForTotalMoneyType item in listData)
            {
                double sumTong = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;
                item.ReportID = string.Format("Năm {0}", item.Year);
                item.TongDS = Math.Round(sumTong, 2, MidpointRounding.ToEven);
                item.Type = 0;
            }

            return Json(listData);
        }

        public ActionResult ReportGradationCompare(string gradation, int? year, string reportTypeID)
        {
            string nameUrl = "Doanh số/Loại tiền/So sánh/Theo giai đoạn";
            ViewBag.NameURL = nameUrl;

            //int typeID = 7;
            //int year = DateTime.Today.Year;
            //List<ReportForTotalMoneyType> listData = new ReportBL().ReportTotalMoneyType(year, typeID);

            DataTable table = new DataTable();
            table.Columns.Add("ReportID", typeof(String));
            table.Columns.Add("AccumulateID1", typeof(double));
            table.Columns.Add("AccumulateID2", typeof(double));
            table.Columns.Add("CompareToIDPercent", typeof(double));
            table.Columns.Add("CompareToID", typeof(double));
            table.PrimaryKey = new DataColumn[] { table.Columns["ReportID"] };

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

            return View(table);
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
            List<ReportForTotalMoneyType> listData = new HSReportBL().SearchReportHSTotalMoneyTypeForGradationCompare(year, gradation, reportTypeID);

            // Khởi tạo datatable
            DataTable table = new DataTable();
            // Tạo các cột cho datatable
            table.Columns.Add("ReportID", typeof(String));
            table.Columns.Add("AccumulateID1", typeof(double));
            table.Columns.Add("AccumulateID2", typeof(double));
            table.Columns.Add("CompareToIDPercent", typeof(double));
            table.Columns.Add("CompareToID", typeof(double));

            string[] str = { "VND", "USD", "EUR", "CAD", "AUD", "GBP" };

            if (listData.Count.Equals(2))
            {
                double VNDCompare = listData[0].VND - listData[1].VND;
                double USDCompare = listData[0].USD - listData[1].USD;
                double EURCompare = listData[0].EUR - listData[1].EUR;
                double CADCompare = listData[0].CAD - listData[1].CAD;
                double AUDCompare = listData[0].AUD - listData[1].AUD;
                double GBPCompare = listData[0].GBP - listData[1].GBP;

                // add row vào table
                table.Rows.Add(str[0], listData[0].VND, listData[1].VND, Math.Round(VNDCompare / listData[1].VND * 100, 2, MidpointRounding.ToEven), VNDCompare);
                table.Rows.Add(str[1], listData[0].USD, listData[1].USD, Math.Round(USDCompare / listData[1].USD * 100, 2, MidpointRounding.ToEven), USDCompare);
                table.Rows.Add(str[2], listData[0].EUR, listData[1].EUR, Math.Round(EURCompare / listData[1].EUR * 100, 2, MidpointRounding.ToEven), EURCompare);
                table.Rows.Add(str[3], listData[0].CAD, listData[1].CAD, Math.Round(CADCompare / listData[1].CAD * 100, 2, MidpointRounding.ToEven), CADCompare);
                table.Rows.Add(str[4], listData[0].AUD, listData[1].AUD, Math.Round(AUDCompare / listData[1].AUD * 100, 2, MidpointRounding.ToEven), AUDCompare);
                table.Rows.Add(str[5], listData[0].GBP, listData[1].GBP, Math.Round(GBPCompare / listData[1].GBP * 100, 2, MidpointRounding.ToEven), GBPCompare);

                DataRow row = table.NewRow();
                row["ReportID"] = "Tổng";
                row["AccumulateID1"] = table.Compute("Sum(AccumulateID1)", "");
                row["AccumulateID2"] = table.Compute("Sum(AccumulateID2)", "");
                row["CompareToIDPercent"] = Math.Round((double)table.Compute("AVG(CompareToIDPercent)", ""), 2, MidpointRounding.ToEven);
                row["CompareToID"] = table.Compute("Sum(CompareToID)", "");
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
        public ActionResult SearchColumnChartReportForGradation([DataSourceRequest]DataSourceRequest request, int gradation, int year, string reportTypeID)
        {
            List<ReportForTotalMoneyType> listData = new HSReportBL().SearchReportHSTotalMoneyTypeForGradationCompare(year, gradation, reportTypeID);

            GradationCompare[] arrayGradation = null;

            if (listData.Count.Equals(2))
            {
                // Số record của mảng
                int countArray = 6;
                arrayGradation = new GradationCompare[countArray * listData.Count];
                int count = 0;
                foreach (ReportForTotalMoneyType item in listData)
                {
                    // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                    arrayGradation[count] = new GradationCompare()
                    {
                        NameGradationCompare = string.Concat("Lũy kế 3 tháng năm ", year),
                        amount = item.VND,
                        NameType = "VND"
                    };

                    count++;
                    arrayGradation[count] = new GradationCompare()
                    {
                        NameGradationCompare = string.Concat("Lũy kế 3 tháng năm ", year),
                        amount = item.USD,
                        NameType = "USD"
                    };

                    count++;
                    arrayGradation[count] = new GradationCompare()
                    {
                        NameGradationCompare = string.Concat("Lũy kế 3 tháng năm ", year),
                        amount = item.EUR,
                        NameType = "EUR"
                    };

                    count++;
                    arrayGradation[count] = new GradationCompare()
                    {
                        NameGradationCompare = string.Concat("Lũy kế 3 tháng năm ", year),
                        amount = item.CAD,
                        NameType = "CAD"
                    };

                    count++;
                    arrayGradation[count] = new GradationCompare()
                    {
                        NameGradationCompare = string.Concat("Lũy kế 3 tháng năm ", year),
                        amount = item.AUD,
                        NameType = "AUD"
                    };

                    count++;
                    arrayGradation[count] = new GradationCompare()
                    {
                        NameGradationCompare = string.Concat("Lũy kế 3 tháng năm ", year),
                        amount = item.GBP,
                        NameType = "GBP"
                    };

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
        public ActionResult SearchGradationComparePieForYear([DataSourceRequest]DataSourceRequest request, int gradation, int year, string reportTypeID)
        {

            List<ReportForTotalMoneyType> listData = new HSReportBL().SearchReportHSTotalMoneyTypeForGradationComparePercent(year, gradation, reportTypeID);

            List<ReportForTotalMoneyType> listDataForYear = listData.Where(x => x.Year == year.ToString()).ToList();
            // # dòng record
            GradationChartPie[] arrayGradation = null;

            if (listDataForYear.Count == 0)
            {
                arrayGradation = new GradationChartPie[1];
                arrayGradation[0] = new GradationChartPie()
                {
                    category = "1",
                    value = 0,
                    color = "#9de219"

                };
            }
            else
            {
                // tạo mảng gồm 8 object
                arrayGradation = new GradationChartPie[6];
            }

            int count = 0;
            foreach (ReportForTotalMoneyType item in listDataForYear)
            {
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
        public ActionResult SearchGradationComparePieForLastYear([DataSourceRequest]DataSourceRequest request, int gradation, int year, string reportTypeID)
        {
            List<ReportForTotalMoneyType> listData = new HSReportBL().SearchReportHSTotalMoneyTypeForGradationComparePercent(year, gradation, reportTypeID);

            List<ReportForTotalMoneyType> listDataForYear = listData.Where(x => x.Year == (year - 1).ToString()).ToList();
            // # dòng record
            GradationChartPie[] arrayGradation = null;

            if (listDataForYear.Count == 0)
            {
                arrayGradation = new GradationChartPie[1];
                arrayGradation[0] = new GradationChartPie()
                {
                    category = "1",
                    value = 0,
                    color = "#9de219"

                };
            }
            else
            {
                // tạo mảng gồm 8 object
                arrayGradation = new GradationChartPie[6];
            }

            int count = 0;
            foreach (ReportForTotalMoneyType item in listDataForYear)
            {
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
            }

            return Json(arrayGradation);
        }


        /// <summary>
        /// get data default cho Grid
        /// </summary>
        /// <param name="request"></param>
        /// <returns></returns>
        [HttpPost]
        public ActionResult SearchReportGradationComparePercent([DataSourceRequest]DataSourceRequest request, int gradation, int year, string reportTypeID)
        {
            List<ReportForTotalMoneyType> listData = new HSReportBL().SearchReportHSTotalMoneyTypeForGradationComparePercent(year, gradation, reportTypeID);

            // Khởi tạo datatable
            DataTable table = new DataTable();
            // Tạo các cột cho datatable
            table.Columns.Add("ReportID", typeof(String));
            table.Columns.Add("AccumulateID1", typeof(double));
            table.Columns.Add("AccumulateID2", typeof(double));
            table.Columns.Add("CompareToIDPercent", typeof(double));

            string[] str = { "VND", "USD", "EUR", "CAD", "AUD", "GBP" };

            if (listData.Count.Equals(2))
            {
                double VNDCompare = listData[0].VND - listData[1].VND;
                double USDCompare = listData[0].USD - listData[1].USD;
                double EURCompare = listData[0].EUR - listData[1].EUR;
                double CADCompare = listData[0].CAD - listData[1].CAD;
                double AUDCompare = listData[0].AUD - listData[1].AUD;
                double GBPCompare = listData[0].GBP - listData[1].GBP;

                // add row vào table
                table.Rows.Add(str[0], listData[0].VND, listData[1].VND, VNDCompare);
                table.Rows.Add(str[1], listData[0].USD, listData[1].USD, USDCompare);
                table.Rows.Add(str[2], listData[0].EUR, listData[1].EUR, EURCompare);
                table.Rows.Add(str[3], listData[0].CAD, listData[1].CAD, CADCompare);
                table.Rows.Add(str[4], listData[0].AUD, listData[1].AUD, AUDCompare);
                table.Rows.Add(str[5], listData[0].GBP, listData[1].GBP, GBPCompare);

                DataRow row = table.NewRow();
                row["ReportID"] = "Tổng";
                row["AccumulateID1"] = 100;
                row["AccumulateID2"] = 100;
                row["CompareToIDPercent"] = 0;
                table.Rows.Add(row);
            }
            return Json(table.ToDataSourceResult(request), JsonRequestBehavior.AllowGet);
        }


        public ActionResult ReportCompareForMonth(int? month, int? year, string reportTypeID)
        {
            string nameUrl = "Doanh số/Loại tiền/So sánh/Theo tháng";
            ViewBag.NameURL = nameUrl;

            DataTable table = new DataTable();
            table.Columns.Add("ReportID", typeof(String));
            table.Columns.Add("AccumulateID1", typeof(double));
            table.Columns.Add("AccumulateID2", typeof(double));
            table.Columns.Add("AccumulateID3", typeof(double));
            table.Columns.Add("CompareToMonth", typeof(double));
            table.Columns.Add("CompareToMonthPercent", typeof(double));
            table.Columns.Add("CompareToMonthLastYear", typeof(double));
            table.Columns.Add("CompareToMonthLastYearPercent", typeof(double));
            table.PrimaryKey = new DataColumn[] { table.Columns["ReportID"] };

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
        /// Get data init cho màn hình so sánh theo tháng
        /// </summary>
        /// <param name="request"></param>
        /// <returns></returns>
        [HttpPost]
        public ActionResult SearchReportCompareForMonth([DataSourceRequest]DataSourceRequest request, int month, int year, string reportTypeID)
        {
            List<ReportForTotalMoneyType> listData = new HSReportBL().SearchReportHSTotalMoneyTypeForCompareMonth(year, month, reportTypeID);

            DataTable table = new DataTable();
            table.Columns.Add("ReportID", typeof(String));
            table.Columns.Add("AccumulateID1", typeof(double));
            table.Columns.Add("AccumulateID2", typeof(double));
            table.Columns.Add("AccumulateID3", typeof(double));
            table.Columns.Add("CompareToMonth", typeof(double));
            table.Columns.Add("CompareToMonthPercent", typeof(double));
            table.Columns.Add("CompareToMonthLastYear", typeof(double));
            table.Columns.Add("CompareToMonthLastYearPercent", typeof(double));
            table.PrimaryKey = new DataColumn[] { table.Columns["ReportID"] };

            string[] str = { "VND", "USD", "EUR", "CAD", "AUD", "GBP" };

            if (listData.Count.Equals(3))
            {
                // tháng hiện tại
                double VNDCompareMonth = listData[0].VND - listData[1].VND;
                double USDCompareMonth = listData[0].USD - listData[1].USD;
                double EURCompareMonth = listData[0].EUR - listData[1].EUR;
                double CADCompareMonth = listData[0].CAD - listData[1].CAD;
                double AUDCompareMonth = listData[0].AUD - listData[1].AUD;
                double GBPCompareMonth = listData[0].GBP - listData[1].GBP;

                // tháng cùng kì năm trước
                double VNDCompareLastMonth = listData[0].VND - listData[2].VND;
                double USDCompareLastMonth = listData[0].USD - listData[2].USD;
                double EURCompareLastMonth = listData[0].EUR - listData[2].EUR;
                double CADCompareLastMonth = listData[0].CAD - listData[2].CAD;
                double AUDCompareLastMonth = listData[0].AUD - listData[2].AUD;
                double GBPCompareLastMonth = listData[0].GBP - listData[2].GBP;

                // add row vào table
                table.Rows.Add(str[0], listData[0].VND, listData[1].VND, listData[2].VND
                    , listData[1].VND, Math.Round(VNDCompareMonth / listData[1].VND * 100, 2, MidpointRounding.ToEven)
                    , listData[2].VND, Math.Round(VNDCompareLastMonth / listData[2].VND * 100, 2, MidpointRounding.ToEven));

                table.Rows.Add(str[1], listData[0].USD, listData[1].USD, listData[2].USD
                    , listData[1].USD, Math.Round(USDCompareMonth / listData[1].USD * 100, 2, MidpointRounding.ToEven)
                    , listData[2].USD, Math.Round(USDCompareLastMonth / listData[2].USD * 100, 2, MidpointRounding.ToEven));

                table.Rows.Add(str[2], listData[0].EUR, listData[1].EUR, listData[2].EUR
                    , listData[1].EUR, Math.Round(EURCompareMonth / listData[1].EUR * 100, 2, MidpointRounding.ToEven)
                    , listData[2].EUR, Math.Round(EURCompareLastMonth / listData[2].EUR * 100, 2, MidpointRounding.ToEven));

                table.Rows.Add(str[3], listData[0].CAD, listData[1].CAD, listData[2].CAD
                    , listData[1].CAD, Math.Round(CADCompareMonth / listData[1].CAD * 100, 2, MidpointRounding.ToEven)
                    , listData[2].CAD, Math.Round(CADCompareLastMonth / listData[2].CAD * 100, 2, MidpointRounding.ToEven));

                table.Rows.Add(str[4], listData[0].AUD, listData[1].AUD, listData[2].AUD
                    , listData[1].AUD, Math.Round(AUDCompareMonth / listData[1].AUD * 100, 2, MidpointRounding.ToEven)
                    , listData[2].AUD, Math.Round(AUDCompareLastMonth / listData[2].AUD * 100, 2, MidpointRounding.ToEven));

                table.Rows.Add(str[5], listData[0].GBP, listData[1].GBP, listData[2].GBP
                    , listData[1].GBP, Math.Round(GBPCompareMonth / listData[1].GBP * 100, 2, MidpointRounding.ToEven)
                    , listData[2].GBP, Math.Round(GBPCompareLastMonth / listData[2].GBP * 100, 2, MidpointRounding.ToEven));

            }

            DataRow row = table.NewRow();
            row["ReportID"] = "Tổng";
            row["AccumulateID1"] = table.Compute("Sum(AccumulateID1)", "");
            row["AccumulateID2"] = table.Compute("Sum(AccumulateID2)", "");
            row["AccumulateID3"] = table.Compute("Sum(AccumulateID3)", "");

            row["CompareToMonth"] = table.Compute("Sum(CompareToMonth)", "");
            row["CompareToMonthPercent"] = table.Compute("Sum(CompareToMonthPercent)", "");
            row["CompareToMonthLastYear"] = table.Compute("Sum(CompareToMonthLastYear)", "");
            row["CompareToMonthLastYearPercent"] = table.Compute("Sum(CompareToMonthLastYearPercent)", "");
            
            table.Rows.Add(row);
            return Json(table.ToDataSourceResult(request), JsonRequestBehavior.AllowGet);
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
            List<ReportForTotalMoneyType> listData = new HSReportBL().SearchReportHSTotalMoneyTypeForCompareMonth(year, month, reportTypeID);

            GradationCompare[] arrayGradation = null;

            if (listData.Count.Equals(3))
            {
                // Tổng số cột là 18, 3 cột tổng
                arrayGradation = new GradationCompare[18];
                int count = 0;

                foreach (ReportForTotalMoneyType item in listData)
                {
                    // Tính tổng doanh số
                    item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;
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
        public ActionResult SearchGradationComparePieForMonth([DataSourceRequest]DataSourceRequest request, int month, int year, string reportTypeID)
        {
            List<ReportForTotalMoneyType> listData = new HSReportBL().SearchReportHSTotalMoneyTypeForCompareMonthPercent(year, month, reportTypeID);
            ReportForTotalMoneyType dataPercent = listData.Find(x => x.Month == month.ToString() && x.Year == year.ToString());
            // # dòng record
            GradationChartPie[] arrayGradation = null;

            if (dataPercent != null)
            {
                // tạo mảng gồm 8 object
                arrayGradation = new GradationChartPie[6];

                int count = 0;
                // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                arrayGradation[count] = new GradationChartPie()
                {
                    category = "VND",
                    value = dataPercent.VND,
                    color = "#FFBF00"
                };

                count++;
                arrayGradation[count] = new GradationChartPie()
                {
                    category = "USD",
                    value = dataPercent.USD,
                    color = "#40FF00"
                };

                count++;
                arrayGradation[count] = new GradationChartPie()
                {
                    category = "EUR",
                    value = dataPercent.EUR,
                    color = "#2ECCFA"
                };

                count++;
                arrayGradation[count] = new GradationChartPie()
                {
                    category = "CAD",
                    value = dataPercent.CAD,
                    color = "#9A2EFE"
                };

                count++;
                arrayGradation[count] = new GradationChartPie()
                {
                    category = "AUD",
                    value = dataPercent.AUD,
                    color = "#FE2EF7"
                };

                count++;
                arrayGradation[count] = new GradationChartPie()
                {
                    category = "GBP",
                    value = dataPercent.GBP,
                    color = "#0000FF"
                };
            }
            else
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
        public ActionResult SearchGradationComparePieForMonthLastMonth([DataSourceRequest]DataSourceRequest request, int month, int year, string reportTypeID)
        {
            List<ReportForTotalMoneyType> listData = new HSReportBL().SearchReportHSTotalMoneyTypeForCompareMonthPercent(year, month, reportTypeID);

            ReportForTotalMoneyType dataItem = null;

            if (month == 1)
            {
                dataItem = listData.Find(x => x.Month == "12" && x.Year == (year - 1).ToString());
            }
            else
            {
                dataItem = listData.Find(x => x.Month == (month - 1).ToString() && x.Year == year.ToString());
            }

            // # dòng record
            GradationChartPie[] arrayGradation = null;

            if (dataItem != null)
            {
                // tạo mảng gồm 8 object
                arrayGradation = new GradationChartPie[6];

                int count = 0;
                // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                arrayGradation[count] = new GradationChartPie()
                {
                    category = "VND",
                    value = dataItem.VND,
                    color = "#FFBF00"
                };

                count++;
                arrayGradation[count] = new GradationChartPie()
                {
                    category = "USD",
                    value = dataItem.USD,
                    color = "#40FF00"
                };

                count++;
                arrayGradation[count] = new GradationChartPie()
                {
                    category = "EUR",
                    value = dataItem.EUR,
                    color = "#2ECCFA"
                };

                count++;
                arrayGradation[count] = new GradationChartPie()
                {
                    category = "CAD",
                    value = dataItem.CAD,
                    color = "#9A2EFE"
                };

                count++;
                arrayGradation[count] = new GradationChartPie()
                {
                    category = "AUD",
                    value = dataItem.AUD,
                    color = "#FE2EF7"
                };

                count++;
                arrayGradation[count] = new GradationChartPie()
                {
                    category = "GBP",
                    value = dataItem.GBP,
                    color = "#0000FF"
                };
            }
            else
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
        /// search data cho biểu đồ của tháng hiện tại
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult SearchGradationComparePieForMonthLastYear([DataSourceRequest]DataSourceRequest request, int month, int year, string reportTypeID)
        {
            List<ReportForTotalMoneyType> listData = new HSReportBL().SearchReportHSTotalMoneyTypeForCompareMonthPercent(year, month, reportTypeID);
            ReportForTotalMoneyType dataItem = listData.Find(x => x.Month == month.ToString() && x.Year == (year - 1).ToString());
            // # dòng record
            GradationChartPie[] arrayGradation = null;

            if (dataItem != null)
            {
                // tạo mảng gồm 8 object
                arrayGradation = new GradationChartPie[6];

                int count = 0;
                // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                arrayGradation[count] = new GradationChartPie()
                {
                    category = "VND",
                    value = dataItem.VND,
                    color = "#FFBF00"
                };

                count++;
                arrayGradation[count] = new GradationChartPie()
                {
                    category = "USD",
                    value = dataItem.USD,
                    color = "#40FF00"
                };

                count++;
                arrayGradation[count] = new GradationChartPie()
                {
                    category = "EUR",
                    value = dataItem.EUR,
                    color = "#2ECCFA"
                };

                count++;
                arrayGradation[count] = new GradationChartPie()
                {
                    category = "CAD",
                    value = dataItem.CAD,
                    color = "#9A2EFE"
                };

                count++;
                arrayGradation[count] = new GradationChartPie()
                {
                    category = "AUD",
                    value = dataItem.AUD,
                    color = "#FE2EF7"
                };

                count++;
                arrayGradation[count] = new GradationChartPie()
                {
                    category = "GBP",
                    value = dataItem.GBP,
                    color = "#0000FF"
                };
            }
            else
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
        /// Search dữ liệu cho báo cáo so sánh theo tháng hiện tại với tháng trước và cùng kì năm ngoái (%)
        /// </summary>
        /// <param name="request"></param>
        /// <param name="month"></param>
        /// <param name="year"></param>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult SearchReportCompareForMonthPercent([DataSourceRequest]DataSourceRequest request, int month, int year, string reportTypeID)
        {
            List<ReportForTotalMoneyType> listData = new HSReportBL().SearchReportHSTotalMoneyTypeForCompareMonthPercent(year, month, reportTypeID);


            // Khởi tạo datatable
            DataTable table = new DataTable();
            // Tạo các cột cho datatable
            table.Columns.Add("ReportID", typeof(String));
            table.Columns.Add("AccumulateID1", typeof(double));
            table.Columns.Add("AccumulateID2", typeof(double));
            table.Columns.Add("AccumulateID3", typeof(double));
            table.Columns.Add("CompareToMonthPercent", typeof(double));
            table.Columns.Add("CompareToMonthLastYearPercent", typeof(double));
            table.PrimaryKey = new DataColumn[] { table.Columns["ReportID"] };

            string[] str = { "VND", "USD", "EUR", "CAD", "AUD", "GBP" };

            if (listData.Count.Equals(3))
            {
                // tháng hiện tại
                double VNDCompareMonth = listData[0].VND - listData[1].VND;
                double USDCompareMonth = listData[0].USD - listData[1].USD;
                double EURCompareMonth = listData[0].EUR - listData[1].EUR;
                double CADCompareMonth = listData[0].CAD - listData[1].CAD;
                double AUDCompareMonth = listData[0].AUD - listData[1].AUD;
                double GBPCompareMonth = listData[0].GBP - listData[1].GBP;

                // tháng cùng kì năm trước
                double VNDCompareMonthLastYear = listData[0].VND - listData[2].VND;
                double USDCompareMonthLastYear = listData[0].USD - listData[2].USD;
                double EURCompareMonthLastYear = listData[0].EUR - listData[2].EUR;
                double CADCompareMonthLastYear = listData[0].CAD - listData[2].CAD;
                double AUDCompareMonthLastYear = listData[0].AUD - listData[2].AUD;
                double GBPCompareMonthLastYear = listData[0].GBP - listData[2].GBP;

                // add row vào table
                table.Rows.Add(str[0], listData[0].VND, listData[1].VND, listData[2].VND
                    , listData[1].VND.Equals(0) ? 0 : Math.Round(VNDCompareMonth / listData[1].VND * 100, 2, MidpointRounding.ToEven)
                    , listData[2].VND.Equals(0) ? 0 : Math.Round(VNDCompareMonthLastYear / listData[2].VND * 100, 2, MidpointRounding.ToEven));

                table.Rows.Add(str[1], listData[0].USD, listData[1].USD, listData[2].USD
                    , listData[1].USD.Equals(0) ? 0 : Math.Round(USDCompareMonth / listData[1].USD * 100, 2, MidpointRounding.ToEven)
                    , listData[2].USD.Equals(0) ? 0 : Math.Round(USDCompareMonthLastYear / listData[2].USD * 100, 2, MidpointRounding.ToEven));

                table.Rows.Add(str[2], listData[0].EUR, listData[1].EUR, listData[2].EUR
                    , listData[1].EUR.Equals(0) ? 0 : Math.Round(EURCompareMonth / listData[1].EUR * 100, 2, MidpointRounding.ToEven)
                    , listData[2].EUR.Equals(0) ? 0 : Math.Round(EURCompareMonthLastYear / listData[2].EUR * 100, 2, MidpointRounding.ToEven));

                table.Rows.Add(str[3], listData[0].CAD, listData[1].CAD, listData[2].CAD
                    , listData[1].CAD.Equals(0) ? 0 : Math.Round(CADCompareMonth / listData[1].CAD * 100, 2, MidpointRounding.ToEven)
                    , listData[2].CAD.Equals(0) ? 0 : Math.Round(CADCompareMonthLastYear / listData[2].CAD * 100, 2, MidpointRounding.ToEven));

                table.Rows.Add(str[4], listData[0].AUD, listData[1].AUD, listData[2].AUD
                    , listData[1].AUD.Equals(0) ? 0 : Math.Round(AUDCompareMonth / listData[1].AUD * 100, 2, MidpointRounding.ToEven)
                    , listData[2].AUD.Equals(0) ? 0 : Math.Round(AUDCompareMonthLastYear / listData[2].AUD * 100, 2, MidpointRounding.ToEven));

                table.Rows.Add(str[5], listData[0].GBP, listData[1].GBP, listData[2].GBP
                    , listData[1].GBP.Equals(0) ? 0 : Math.Round(GBPCompareMonth / listData[1].GBP * 100, 2, MidpointRounding.ToEven)
                    , listData[2].GBP.Equals(0) ? 0 : Math.Round(GBPCompareMonthLastYear / listData[2].GBP * 100, 2, MidpointRounding.ToEven));

                DataRow row = table.NewRow();
                row["ReportID"] = "Tổng";
                row["AccumulateID1"] = 100;
                row["AccumulateID2"] = 100;
                row["AccumulateID3"] = 100;

                // Sum row tổng compare month
                row["CompareToMonthPercent"] = 0;

                row["CompareToMonthLastYearPercent"] = 0;
                table.Rows.Add(row);
            }
            return Json(table.ToDataSourceResult(request), JsonRequestBehavior.AllowGet);
        }
    }
}
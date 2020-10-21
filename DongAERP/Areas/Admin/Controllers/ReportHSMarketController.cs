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
    public class ReportHSMarketController : Controller
    {
        // GET: Admin/ReportHSMarket
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult ReportDay(DateTime? fromDate, DateTime? toDate, string reportTypeID)
        {
            string nameUrl = "Hồ sơ/Theo thị trường/Chi tiết/Theo Ngày";
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
        /// <param name="toDayte"></param>
        /// <returns></returns>
        public ActionResult SearchReportMaketForDay([DataSourceRequest]DataSourceRequest request, DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            List<ReportForMaket> listData = new HSReportBL().SearchReportMaketForDay(fromDate, toDate, reportTypeID);
            foreach (ReportForMaket item in listData)
            {
                double sumTong = item.American + item.Asia + item.Global + item.Europe + item.Canada + item.Australia;
                item.ReportID = item.CreatedDate.ToString("dd/MM/yyyy");
                item.TongDS = Math.Round(sumTong, 2, MidpointRounding.ToEven);
                item.Type = 0;
            }

            ReportForMaket dataItem = new ReportForMaket()
            {
                ReportID = "Tổng",
                American = listData.Sum(x => x.American),
                Asia = listData.Sum(x => x.Asia),
                Global = listData.Sum(x => x.Global),
                Europe = listData.Sum(x => x.Europe),
                Canada = listData.Sum(x => x.Canada),
                Australia = listData.Sum(x => x.Australia),
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
        public ActionResult SearchLineChartMaketForDay([DataSourceRequest]DataSourceRequest request, DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            List<ReportForMaket> listData = new HSReportBL().SearchReportMaketForDay(fromDate, toDate, reportTypeID);
            foreach (ReportForMaket item in listData)
            {
                double sumTong = item.American + item.Asia + item.Global + item.Europe + item.Canada + item.Australia;
                item.ReportID = item.CreatedDate.ToString("dd/MM/yyyy");
                item.TongDS = Math.Round(sumTong, 2, MidpointRounding.ToEven);
                item.Type = 0;
            }

            return Json(listData);
        }

        public ActionResult ReportMonth(DateTime? fromDate, DateTime? toDate, string reportTypeID)
        {
            string nameUrl = "Hồ sơ/Theo thị trường/Chi tiết/Theo Tháng";
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
        /// <param name="toDayte"></param>
        /// <returns></returns>
        public ActionResult SearchReportMaketForMonth([DataSourceRequest]DataSourceRequest request, DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            List<ReportForMaket> listData = new HSReportBL().SearchReportMaketForMonth(fromDate, toDate, reportTypeID);
            foreach (ReportForMaket item in listData)
            {
                double sumTong = item.American + item.Asia + item.Global + item.Europe + item.Canada + item.Australia;
                item.ReportID = string.Format("Tháng {0}/{1}", item.Month, item.Year);
                item.TongDS = Math.Round(sumTong, 2, MidpointRounding.ToEven);
                item.Type = 0;
            }

            ReportForMaket dataItem = new ReportForMaket()
            {
                ReportID = "Tổng",
                American = listData.Sum(x => x.American),
                Asia = listData.Sum(x => x.Asia),
                Global = listData.Sum(x => x.Global),
                Europe = listData.Sum(x => x.Europe),
                Canada = listData.Sum(x => x.Canada),
                Australia = listData.Sum(x => x.Australia),
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
        public ActionResult SearchLineChartMaketForMonth([DataSourceRequest]DataSourceRequest request, DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            List<ReportForMaket> listData = new HSReportBL().SearchReportMaketForMonth(fromDate, toDate, reportTypeID);
            foreach (ReportForMaket item in listData)
            {
                double sumTong = item.American + item.Asia + item.Global + item.Europe + item.Canada + item.Australia;
                item.ReportID = string.Format("Tháng {0}/{1}", item.Month, item.Year);
                item.TongDS = Math.Round(sumTong, 2, MidpointRounding.ToEven);
                item.Type = 0;
            }

            return Json(listData);
        }


        public ActionResult ReportYear(DateTime? fromDate, DateTime? toDate, string reportTypeID)
        {
            string nameUrl = "Hồ sơ/Theo thị trường/Chi tiết/Theo Năm";
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
        /// <param name="toDayte"></param>
        /// <returns></returns>
        public ActionResult SearchReportMaketForYear([DataSourceRequest]DataSourceRequest request, DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            List<ReportForMaket> listData = new HSReportBL().SearchReportMaketForYear(fromDate, toDate, reportTypeID);
            foreach (ReportForMaket item in listData)
            {
                double sumTong = item.American + item.Asia + item.Global + item.Europe + item.Canada + item.Australia;
                item.ReportID = string.Format("Năm {0}", item.Year);
                item.TongDS = Math.Round(sumTong, 2, MidpointRounding.ToEven);
                item.Type = 0;
            }

            ReportForMaket dataItem = new ReportForMaket()
            {
                ReportID = "Tổng",
                American = listData.Sum(x => x.American),
                Asia = listData.Sum(x => x.Asia),
                Global = listData.Sum(x => x.Global),
                Europe = listData.Sum(x => x.Europe),
                Canada = listData.Sum(x => x.Canada),
                Australia = listData.Sum(x => x.Australia),
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
        public ActionResult SearchLineChartMaketForYear([DataSourceRequest]DataSourceRequest request, DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            List<ReportForMaket> listData = new HSReportBL().SearchReportMaketForYear(fromDate, toDate, reportTypeID);
            foreach (ReportForMaket item in listData)
            {
                double sumTong = item.American + item.Asia + item.Global + item.Europe + item.Canada + item.Australia;
                item.ReportID = string.Format("Năm {0}", item.Year);
                item.TongDS = Math.Round(sumTong, 2, MidpointRounding.ToEven);
                item.Type = 0;
            }

            return Json(listData);
        }

        public ActionResult ReportGradationCompare(string gradation, int? year, string reportTypeID)
        {
            string nameUrl = "Hồ sơ/Theo thị trường/So sánh/Theo giai đoạn";
            ViewBag.NameURL = nameUrl;

            //int typeID = 4;
            //int year = DateTime.Today.Year;
            //List<ReportForMaket> listData = new ReportBL().DataReportMaketForGradationCompare(year, typeID);

            DataTable table = new DataTable();
            table.Columns.Add("MaketID", typeof(String));
            table.Columns.Add("AccumulateID1", typeof(double));
            table.Columns.Add("AccumulateID2", typeof(double));
            table.Columns.Add("CompareToID", typeof(double));
            table.Columns.Add("CompareToIDPercent", typeof(double));
            table.PrimaryKey = new DataColumn[] { table.Columns["MaketID"] };

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
        public ActionResult SearchReportGradationCompare([DataSourceRequest]DataSourceRequest request, string gradation, int year, string reportTypeID)
        {
            List<ReportForMaket> listData = new HSReportBL().SearchReportMaketForGradationCompare(year, int.Parse(gradation), reportTypeID);

            // Khởi tạo datatable
            DataTable table = new DataTable();
            // Tạo các cột cho datatable
            table.Columns.Add("MaketID", typeof(String));
            table.Columns.Add("AccumulateID1", typeof(double));
            table.Columns.Add("AccumulateID2", typeof(double));
            table.Columns.Add("CompareToIDPercent", typeof(double));
            table.Columns.Add("CompareToID", typeof(double));

            string[] str = { "Mỹ", "Châu Á", "Toàn cầu", "Châu Âu", "Canada", "Úc" };

            if (listData.Count.Equals(2))
            {
                double americanCompare = listData[0].American - listData[1].American;
                double asiaCompare = listData[0].Asia - listData[1].Asia;
                double globalCompare = listData[0].Global - listData[1].Global;
                double europeCompare = listData[0].Europe - listData[1].Europe;
                double canadaCompare = listData[0].Canada - listData[1].Canada;
                double australiaCompare = listData[0].Australia - listData[1].Australia;

                // add row vào table
                table.Rows.Add(str[0], listData[0].American, listData[1].American, listData[1].American == 0 ? 0 :
                    Math.Round(americanCompare / listData[1].American * 100, 2, MidpointRounding.ToEven), americanCompare);
                table.Rows.Add(str[1], listData[0].Asia, listData[1].Asia, listData[1].Asia == 0 ? 0 :
                    Math.Round(asiaCompare / listData[1].Asia * 100, 2, MidpointRounding.ToEven), asiaCompare);
                table.Rows.Add(str[2], listData[0].Global, listData[1].Global, listData[1].Global == 0 ? 0 :
                    Math.Round(globalCompare / listData[1].Global * 100, 2, MidpointRounding.ToEven), globalCompare);
                table.Rows.Add(str[3], listData[0].Europe, listData[1].Europe, listData[1].Europe == 0 ? 0 :
                    Math.Round(europeCompare / listData[1].Europe * 100, 2, MidpointRounding.ToEven), europeCompare);
                table.Rows.Add(str[4], listData[0].Canada, listData[1].Canada, listData[1].Canada == 0 ? 0 :
                    Math.Round(canadaCompare / listData[1].Canada * 100, 2, MidpointRounding.ToEven), canadaCompare);
                table.Rows.Add(str[5], listData[0].Australia, listData[1].Australia, listData[1].Australia == 0 ? 0 :
                    Math.Round(australiaCompare / listData[1].Australia * 100, 2, MidpointRounding.ToEven), australiaCompare);

                DataRow row = table.NewRow();
                row["MaketID"] = "Tổng";
                double sumAccumulateID1 = 0;
                double sumAccumulateID2 = 0;

                if (!string.IsNullOrEmpty(table.Compute("Sum(AccumulateID1)", "").ToString()))
                {
                    sumAccumulateID1 = Convert.ToDouble(table.Compute("Sum(AccumulateID1)", ""));
                }

                if (!string.IsNullOrEmpty(table.Compute("Sum(AccumulateID2)", "").ToString()))
                {
                    sumAccumulateID2 = Convert.ToDouble(table.Compute("Sum(AccumulateID2)", ""));
                }

                row["AccumulateID1"] = sumAccumulateID1;
                row["AccumulateID2"] = sumAccumulateID2;
                row["CompareToID"] = Math.Round(sumAccumulateID1 - sumAccumulateID2, 2, MidpointRounding.ToEven);
                row["CompareToIDPercent"] = sumAccumulateID2 == 0 ? 0 : Math.Round((sumAccumulateID1 - sumAccumulateID2) / sumAccumulateID2 * 100, 2, MidpointRounding.ToEven);
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
        public ActionResult SearchColumnChartMaketReportForGradation([DataSourceRequest]DataSourceRequest request, int gradation, int year, string reportTypeID)
        {
            List<ReportForMaket> listData = new HSReportBL().SearchReportMaketForGradationCompare(year, gradation, reportTypeID);
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

                arrayGradation = new GradationCompare[12];
                int count = 0;

                foreach (ReportForMaket item in listData)
                {
                    // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                    arrayGradation[count] = new GradationCompare()
                    {
                        NameGradationCompare = string.Format("Lũy kế {0} {1}", text, item.Year),
                        amount = item.American,
                        NameType = "Mỹ"
                    };

                    count++;
                    arrayGradation[count] = new GradationCompare()
                    {
                        NameGradationCompare = string.Format("Lũy kế {0} {1}", text, item.Year),
                        amount = item.Asia,
                        NameType = "Châu Á"
                    };

                    count++;
                    arrayGradation[count] = new GradationCompare()
                    {
                        NameGradationCompare = string.Format("Lũy kế {0} {1}", text, item.Year),
                        amount = item.Global,
                        NameType = "Toàn Cầu"
                    };

                    count++;
                    arrayGradation[count] = new GradationCompare()
                    {
                        NameGradationCompare = string.Format("Lũy kế {0} {1}", text, item.Year),
                        amount = item.Europe,
                        NameType = "Châu Âu"
                    };

                    count++;
                    arrayGradation[count] = new GradationCompare()
                    {
                        NameGradationCompare = string.Format("Lũy kế {0} {1}", text, item.Year),
                        amount = item.Canada,
                        NameType = "Canada"
                    };

                    count++;
                    arrayGradation[count] = new GradationCompare()
                    {
                        NameGradationCompare = string.Format("Lũy kế {0} {1}", text, item.Year),
                        amount = item.Australia,
                        NameType = "Úc"
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
        public ActionResult SearchGradationComparePieForYear([DataSourceRequest]DataSourceRequest request, string gradation, int year, string reportTypeID)
        {
            List<ReportForMaket> listData = new HSReportBL().SearchReportMaketForGradationComparePercent(year, int.Parse(gradation), reportTypeID);

            // # dòng record
            GradationChartPie[] arrayGradation = new GradationChartPie[1];
            arrayGradation[0] = new GradationChartPie()
            {
                category = "1",
                value = 0,
                color = "#9de219"

            };

            int count = 0;
            foreach (ReportForMaket item in listData)
            {
                if (item.Year == year.ToString())
                {
                    // tạo mảng gồm 8 object
                    arrayGradation = new GradationChartPie[6];

                    // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                    arrayGradation[count] = new GradationChartPie()
                    {
                        category = "Mỹ",
                        value = item.American,
                        color = "#FFBF00"
                    };

                    count++;
                    arrayGradation[count] = new GradationChartPie()
                    {
                        category = "Châu Á",
                        value = item.Asia,
                        color = "#40FF00"
                    };

                    count++;
                    arrayGradation[count] = new GradationChartPie()
                    {
                        category = "Toàn cầu",
                        value = item.Global,
                        color = "#2ECCFA"
                    };

                    count++;
                    arrayGradation[count] = new GradationChartPie()
                    {
                        category = "Châu Âu",
                        value = item.Europe,
                        color = "#9A2EFE"
                    };

                    count++;
                    arrayGradation[count] = new GradationChartPie()
                    {
                        category = "Canada",
                        value = item.Canada,
                        color = "#FE2EF7"
                    };

                    count++;
                    arrayGradation[count] = new GradationChartPie()
                    {
                        category = "Úc",
                        value = item.Australia,
                        color = "#0000FF"
                    };
                }
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
        public ActionResult SearchGradationComparePieForLastYear([DataSourceRequest]DataSourceRequest request, string gradation, int year, string reportTypeID)
        {
            List<ReportForMaket> listData = new HSReportBL().SearchReportMaketForGradationComparePercent(year, int.Parse(gradation), reportTypeID);

            // # dòng record
            GradationChartPie[] arrayGradation = new GradationChartPie[1];
            arrayGradation[0] = new GradationChartPie()
            {
                category = "1",
                value = 0,
                color = "#9de219"

            };

            int count = 0;
            foreach (ReportForMaket item in listData)
            {
                if (item.Year == (year - 1).ToString())
                {
                    // tạo mảng gồm 8 object
                    arrayGradation = new GradationChartPie[6];

                    // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                    arrayGradation[count] = new GradationChartPie()
                    {
                        category = "Mỹ",
                        value = item.American,
                        color = "#FFBF00"
                    };

                    count++;
                    arrayGradation[count] = new GradationChartPie()
                    {
                        category = "Châu Á",
                        value = item.Asia,
                        color = "#40FF00"
                    };

                    count++;
                    arrayGradation[count] = new GradationChartPie()
                    {
                        category = "Toàn cầu",
                        value = item.Global,
                        color = "#2ECCFA"
                    };

                    count++;
                    arrayGradation[count] = new GradationChartPie()
                    {
                        category = "Châu Âu",
                        value = item.Europe,
                        color = "#9A2EFE"
                    };

                    count++;
                    arrayGradation[count] = new GradationChartPie()
                    {
                        category = "Canada",
                        value = item.Canada,
                        color = "#FE2EF7"
                    };

                    count++;
                    arrayGradation[count] = new GradationChartPie()
                    {
                        category = "Úc",
                        value = item.Australia,
                        color = "#0000FF"
                    };
                }
            }

            return Json(arrayGradation);
        }


        /// <summary>
        /// get data default cho Grid
        /// </summary>
        /// <param name="request"></param>
        /// <returns></returns>
        [HttpPost]
        public ActionResult SearchReportGradationComparePercent([DataSourceRequest]DataSourceRequest request, string gradation, int year, string reportTypeID)
        {
            List<ReportForMaket> listData = new HSReportBL().SearchReportMaketForGradationComparePercent(year, int.Parse(gradation), reportTypeID);

            // Khởi tạo datatable
            DataTable table = new DataTable();
            // Tạo các cột cho datatable
            table.Columns.Add("MaketID", typeof(String));
            table.Columns.Add("AccumulateID1", typeof(double));
            table.Columns.Add("AccumulateID2", typeof(double));
            table.Columns.Add("CompareToIDPercent", typeof(double));

            string[] str = { "Mỹ", "Châu Á", "Toàn cầu", "Châu Âu", "Canada", "Úc" };

            if (listData.Count.Equals(2))
            {
                ReportForMaket dataYear = listData.Find(x => x.Year == year.ToString());
                ReportForMaket dataLastYear = listData.Find(x => x.Year == (year - 1).ToString());

                double americanCompare = dataYear.American - dataLastYear.American;
                double asiaCompare = dataYear.Asia - dataLastYear.Asia;
                double globalCompare = dataYear.Global - dataLastYear.Global;
                double europeCompare = dataYear.Europe - dataLastYear.Europe;
                double canadaCompare = dataYear.Canada - dataLastYear.Canada;
                double australiaCompare = dataYear.Australia - dataLastYear.Australia;

                // add row vào table
                table.Rows.Add(str[0], dataYear.American, dataLastYear.American, dataLastYear.American == 0 ? 0 : Math.Round((americanCompare / dataLastYear.American) * 100, 2, MidpointRounding.ToEven));
                table.Rows.Add(str[1], dataYear.Asia, dataLastYear.Asia, dataLastYear.Asia == 0 ? 0 : Math.Round((asiaCompare / dataLastYear.Asia) * 100, 2, MidpointRounding.ToEven));
                table.Rows.Add(str[2], dataYear.Global, dataLastYear.Global, dataLastYear.Global == 0 ? 0 : Math.Round((globalCompare / dataLastYear.Global) * 100, 2, MidpointRounding.ToEven));
                table.Rows.Add(str[3], dataYear.Europe, dataLastYear.Europe, dataLastYear.Europe == 0 ? 0 : Math.Round((europeCompare / dataLastYear.Europe) * 100, 2, MidpointRounding.ToEven));
                table.Rows.Add(str[4], dataYear.Canada, dataLastYear.Canada, dataLastYear.Canada == 0 ? 0 : Math.Round((canadaCompare / dataLastYear.Canada) * 100, 2, MidpointRounding.ToEven));
                table.Rows.Add(str[5], dataYear.Australia, dataLastYear.Australia, dataLastYear.Australia == 0 ? 0 : Math.Round((australiaCompare / dataLastYear.Australia) * 100, 2, MidpointRounding.ToEven));

                DataRow row = table.NewRow();
                row["MaketID"] = "Tổng";
                row["AccumulateID1"] = 100;
                row["AccumulateID2"] = 100;
                row["CompareToIDPercent"] = 0;
                table.Rows.Add(row);
            }
            return Json(table.ToDataSourceResult(request), JsonRequestBehavior.AllowGet);
        }


        public ActionResult ReportCompareForMonth(int? month, int? year, string reportTypeID)
        {
            string nameUrl = "Hồ sơ/Theo thị trường/So sánh/Theo tháng";
            ViewBag.NameURL = nameUrl;

            DataTable table = new DataTable();
            table.Columns.Add("MaketID", typeof(String));
            table.Columns.Add("AccumulateID1", typeof(double));
            table.Columns.Add("AccumulateID2", typeof(double));
            table.Columns.Add("AccumulateID3", typeof(double));
            table.Columns.Add("CompareToMonth", typeof(double));
            table.Columns.Add("CompareToMonthPercent", typeof(double));
            table.Columns.Add("CompareToMonthLastYear", typeof(double));
            table.Columns.Add("CompareToMonthLastYearPercent", typeof(double));
            table.PrimaryKey = new DataColumn[] { table.Columns["MaketID"] };

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
            List<ReportForMaket> listData = new HSReportBL().SearchReportMaketForCompareForMonth(year, month, reportTypeID);

            DataTable table = new DataTable();
            table.Columns.Add("MaketID", typeof(String));
            table.Columns.Add("AccumulateID1", typeof(double));
            table.Columns.Add("AccumulateID2", typeof(double));
            table.Columns.Add("AccumulateID3", typeof(double));
            table.Columns.Add("CompareToMonth", typeof(double));
            table.Columns.Add("CompareToMonthPercent", typeof(double));
            table.Columns.Add("CompareToMonthLastYear", typeof(double));
            table.Columns.Add("CompareToMonthLastYearPercent", typeof(double));
            table.PrimaryKey = new DataColumn[] { table.Columns["MaketID"] };

            string[] str = { "Mỹ", "Châu Á", "Toàn cầu", "Châu Âu", "Canada", "Úc" };

            if (listData.Count.Equals(3))
            {
                // tháng hiện tại
                double americanCompareMonth = listData[0].American - listData[1].American;
                double asiaCompareMonth = listData[0].Asia - listData[1].Asia;
                double globalCompareMonth = listData[0].Global - listData[1].Global;
                double europeCompareMonth = listData[0].Europe - listData[1].Europe;
                double canadaCompareMonth = listData[0].Canada - listData[1].Canada;
                double australiaCompareMonth = listData[0].Australia - listData[1].Australia;

                // tháng cùng kì năm trước
                double americanCompareLastMonth = listData[0].American - listData[2].American;
                double asiaCompareLastMonth = listData[0].Asia - listData[2].Asia;
                double globalCompareLastMonth = listData[0].Global - listData[2].Global;
                double europeCompareLastMonth = listData[0].Europe - listData[2].Europe;
                double canadaCompareLastMonth = listData[0].Canada - listData[2].Canada;
                double australiaCompareLastMonth = listData[0].Australia - listData[2].Australia;

                // add row vào table
                table.Rows.Add(str[0], listData[0].American, listData[1].American, listData[2].American
                    , americanCompareMonth, listData[1].American == 0 ? 0 : Math.Round((americanCompareMonth / listData[1].American) * 100, 2, MidpointRounding.ToEven)
                    , americanCompareLastMonth, listData[2].American == 0 ? 0 : Math.Round((americanCompareLastMonth / listData[2].American) * 100, 2, MidpointRounding.ToEven));

                table.Rows.Add(str[1], listData[0].Asia, listData[1].Asia, listData[2].Asia
                    , asiaCompareMonth, listData[1].Asia == 0 ? 0 : Math.Round((asiaCompareMonth / listData[1].Asia) * 100, 2, MidpointRounding.ToEven)
                    , asiaCompareLastMonth, listData[2].Asia == 0 ? 0 : Math.Round((asiaCompareLastMonth / listData[2].Asia) * 100, 2, MidpointRounding.ToEven));

                table.Rows.Add(str[2], listData[0].Global, listData[1].Global, listData[2].Global
                    , globalCompareMonth, listData[1].Global == 0 ? 0 : Math.Round((globalCompareMonth / listData[1].Global) * 100, 2, MidpointRounding.ToEven)
                    , globalCompareLastMonth, listData[2].Global == 0 ? 0 : Math.Round((globalCompareLastMonth / listData[2].Global) * 100, 2, MidpointRounding.ToEven));

                table.Rows.Add(str[3], listData[0].Europe, listData[1].Europe, listData[2].Europe
                    , europeCompareMonth, listData[1].Europe == 0 ? 0 : Math.Round((europeCompareMonth / listData[1].Europe) * 100, 2, MidpointRounding.ToEven)
                    , europeCompareLastMonth, listData[2].Europe == 0 ? 0 : Math.Round((europeCompareLastMonth / listData[2].Europe) * 100, 2, MidpointRounding.ToEven));

                table.Rows.Add(str[4], listData[0].Canada, listData[1].Canada, listData[2].Canada
                    , canadaCompareMonth, listData[1].Canada == 0 ? 0 : Math.Round((canadaCompareMonth / listData[1].Canada) * 100, 2, MidpointRounding.ToEven)
                    , canadaCompareLastMonth, listData[2].Canada == 0 ? 0 : Math.Round((canadaCompareLastMonth / listData[2].Canada) * 100, 2, MidpointRounding.ToEven));

                table.Rows.Add(str[5], listData[0].Australia, listData[1].Australia, listData[2].Australia
                    , australiaCompareMonth, listData[1].Australia == 0 ? 0 : Math.Round((australiaCompareMonth / listData[1].Australia) * 100, 2, MidpointRounding.ToEven)
                    , australiaCompareLastMonth, listData[2].Australia == 0 ? 0 : Math.Round((australiaCompareLastMonth / listData[2].Australia) * 100, 2, MidpointRounding.ToEven));

                DataRow row = table.NewRow();
                row["MaketID"] = "Tổng";
                row["AccumulateID1"] = table.Compute("Sum(AccumulateID1)", "");
                row["AccumulateID2"] = table.Compute("Sum(AccumulateID2)", "");
                row["AccumulateID3"] = table.Compute("Sum(AccumulateID3)", "");

                // Sum row tổng compare month
                double sumCompareMonth = (double)row["AccumulateID1"] - (double)row["AccumulateID2"];
                row["CompareToMonth"] = sumCompareMonth;
                row["CompareToMonthPercent"] = Math.Round(sumCompareMonth / (double)row["AccumulateID2"] * 100, 2, MidpointRounding.ToEven);

                // Sum row tổng compare month
                double sumCompareMonthLastYear = (double)row["AccumulateID1"] - (double)row["AccumulateID3"];

                row["CompareToMonthLastYear"] = sumCompareMonthLastYear;
                row["CompareToMonthLastYearPercent"] = Math.Round(sumCompareMonthLastYear / (double)row["AccumulateID3"] * 100, 2, MidpointRounding.ToEven);
                table.Rows.Add(row);
            }
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
            List<ReportForMaket> listData = new HSReportBL().SearchReportMaketForCompareForMonth(year, month, reportTypeID);

            GradationCompare[] arrayGradation = null;

            if (listData.Count.Equals(3))
            {
                // Tổng số cột là 18, 3 cột tổng
                arrayGradation = new GradationCompare[18];
                int count = 0;

                foreach (ReportForMaket item in listData)
                {
                    // Tính tổng doanh số
                    item.TongDS = item.American + item.Asia + item.Global + item.Europe + item.Canada + item.Australia;
                    // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                    arrayGradation[count] = new GradationCompare()
                    {
                        NameGradationCompare = string.Format("Tháng {0}/{1}", item.Month, item.Year),
                        amount = item.American,
                        NameType = "Mỹ"
                    };

                    count++;
                    arrayGradation[count] = new GradationCompare()
                    {
                        NameGradationCompare = string.Format("Tháng {0}/{1}", item.Month, item.Year),
                        amount = item.Asia,
                        NameType = "Châu Á"
                    };

                    count++;
                    arrayGradation[count] = new GradationCompare()
                    {
                        NameGradationCompare = string.Format("Tháng {0}/{1}", item.Month, item.Year),
                        amount = item.Global,
                        NameType = "Toàn Cầu"
                    };

                    count++;
                    arrayGradation[count] = new GradationCompare()
                    {
                        NameGradationCompare = string.Format("Tháng {0}/{1}", item.Month, item.Year),
                        amount = item.Europe,
                        NameType = "Châu Âu"
                    };

                    count++;
                    arrayGradation[count] = new GradationCompare()
                    {
                        NameGradationCompare = string.Format("Tháng {0}/{1}", item.Month, item.Year),
                        amount = item.Canada,
                        NameType = "Canada"
                    };

                    count++;
                    arrayGradation[count] = new GradationCompare()
                    {
                        NameGradationCompare = string.Format("Tháng {0}/{1}", item.Month, item.Year),
                        amount = item.Australia,
                        NameType = "Úc"
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
            List<ReportForMaket> listData = new ReportBL().DataReportMarketCompareForMonthPercent(year, month, reportTypeID);

            // # dòng record
            GradationChartPie[] arrayGradation = null;

            if (listData.Find(x => x.Month == month.ToString() && x.Year == year.ToString()) != null)
            {
                // tạo mảng gồm 8 object
                arrayGradation = new GradationChartPie[6];

                int count = 0;
                foreach (ReportForMaket item in listData)
                {
                    if (item.Year == year.ToString() && item.Month == month.ToString())
                    {
                        // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                        arrayGradation[count] = new GradationChartPie()
                        {
                            category = "Mỹ",
                            value = item.American,
                            color = "#FFBF00"
                        };

                        count++;
                        arrayGradation[count] = new GradationChartPie()
                        {
                            category = "Châu Á",
                            value = item.Asia,
                            color = "#40FF00"
                        };

                        count++;
                        arrayGradation[count] = new GradationChartPie()
                        {
                            category = "Toàn cầu",
                            value = item.Global,
                            color = "#2ECCFA"
                        };

                        count++;
                        arrayGradation[count] = new GradationChartPie()
                        {
                            category = "Châu Âu",
                            value = item.Europe,
                            color = "#9A2EFE"
                        };

                        count++;
                        arrayGradation[count] = new GradationChartPie()
                        {
                            category = "Canada",
                            value = item.Canada,
                            color = "#FE2EF7"
                        };

                        count++;
                        arrayGradation[count] = new GradationChartPie()
                        {
                            category = "Úc",
                            value = item.Australia,
                            color = "#0000FF"
                        };
                    }
                }
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
        /// Search dữ liệu cho biểu đồ tròn với tháng trước
        /// </summary>
        /// <param name="request"></param>
        /// <param name="month"></param>
        /// <param name="year"></param>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult SearchGradationComparePieForLastMonth([DataSourceRequest]DataSourceRequest request, int month, int year, string reportTypeID)
        {
            List<ReportForMaket> listData = new HSReportBL().SearchReportMaketForCompareMonthPercent(year, month, reportTypeID);

            // # dòng record
            GradationChartPie[] arrayGradation = null;
            // Trường hợp với tháng là 1
            int lastMonth = month - 1;
            int lastyear = year;

            if (lastMonth == 0)
            {
                lastMonth = 12;
                lastyear = year - 1;
            }

            if (listData.Find(x => x.Month == lastMonth.ToString() && x.Year == lastyear.ToString()) != null)
            {
                // tạo mảng gồm 8 object
                arrayGradation = new GradationChartPie[6];

                int count = 0;
                foreach (ReportForMaket item in listData)
                {
                    if (item.Year == lastyear.ToString() && item.Month == lastMonth.ToString())
                    {
                        // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                        arrayGradation[count] = new GradationChartPie()
                        {
                            category = "Mỹ",
                            value = item.American,
                            color = "#FFBF00"
                        };

                        count++;
                        arrayGradation[count] = new GradationChartPie()
                        {
                            category = "Châu Á",
                            value = item.Asia,
                            color = "#40FF00"
                        };

                        count++;
                        arrayGradation[count] = new GradationChartPie()
                        {
                            category = "Toàn cầu",
                            value = item.Global,
                            color = "#2ECCFA"
                        };

                        count++;
                        arrayGradation[count] = new GradationChartPie()
                        {
                            category = "Châu Âu",
                            value = item.Europe,
                            color = "#9A2EFE"
                        };

                        count++;
                        arrayGradation[count] = new GradationChartPie()
                        {
                            category = "Canada",
                            value = item.Canada,
                            color = "#FE2EF7"
                        };

                        count++;
                        arrayGradation[count] = new GradationChartPie()
                        {
                            category = "Úc",
                            value = item.Australia,
                            color = "#0000FF"
                        };
                    }
                }
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
        /// Search dữ liệu cho biểu đồ tròn với tháng cùng kì năm ngoái
        /// </summary>
        /// <param name="request"></param>
        /// <param name="month"></param>
        /// <param name="year"></param>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult SearchGradationComparePieForMonthLastYear([DataSourceRequest]DataSourceRequest request, int month, int year, string reportTypeID)
        {
            List<ReportForMaket> listData = new HSReportBL().SearchReportMaketForCompareMonthPercent(year, month, reportTypeID);

            // # dòng record
            GradationChartPie[] arrayGradation = null;

            if (listData.Find(x => x.Month == month.ToString() && x.Year == (year - 1).ToString()) != null)
            {
                // tạo mảng gồm 8 object
                arrayGradation = new GradationChartPie[6];

                int count = 0;
                foreach (ReportForMaket item in listData)
                {
                    if (item.Year == (year - 1).ToString() && item.Month == month.ToString())
                    {
                        // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                        arrayGradation[count] = new GradationChartPie()
                        {
                            category = "Mỹ",
                            value = item.American,
                            color = "#FFBF00"
                        };

                        count++;
                        arrayGradation[count] = new GradationChartPie()
                        {
                            category = "Châu Á",
                            value = item.Asia,
                            color = "#40FF00"
                        };

                        count++;
                        arrayGradation[count] = new GradationChartPie()
                        {
                            category = "Toàn cầu",
                            value = item.Global,
                            color = "#2ECCFA"
                        };

                        count++;
                        arrayGradation[count] = new GradationChartPie()
                        {
                            category = "Châu Âu",
                            value = item.Europe,
                            color = "#9A2EFE"
                        };

                        count++;
                        arrayGradation[count] = new GradationChartPie()
                        {
                            category = "Canada",
                            value = item.Canada,
                            color = "#FE2EF7"
                        };

                        count++;
                        arrayGradation[count] = new GradationChartPie()
                        {
                            category = "Úc",
                            value = item.Australia,
                            color = "#0000FF"
                        };
                    }
                }
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
            List<ReportForMaket> listData = new HSReportBL().SearchReportMaketForCompareMonthPercent(year, month, reportTypeID);

            // Khởi tạo datatable
            DataTable table = new DataTable();
            // Tạo các cột cho datatable
            table.Columns.Add("MaketID", typeof(String));
            table.Columns.Add("AccumulateID1", typeof(double));
            table.Columns.Add("AccumulateID2", typeof(double));
            table.Columns.Add("AccumulateID3", typeof(double));
            table.Columns.Add("CompareToMonthPercent", typeof(double));
            table.Columns.Add("CompareToMonthLastYearPercent", typeof(double));
            table.PrimaryKey = new DataColumn[] { table.Columns["MaketID"] };

            string[] str = { "Mỹ", "Châu Á", "Toàn cầu", "Châu Âu", "Canada", "Úc" };

            if (listData.Count.Equals(3))
            {
                // tháng hiện tại
                double americanCompareMonth = listData[0].American - listData[1].American;
                double asiaCompareMonth = listData[0].Asia - listData[1].Asia;
                double globalCompareMonth = listData[0].Global - listData[1].Global;
                double europeCompareMonth = listData[0].Europe - listData[1].Europe;
                double canadaCompareMonth = listData[0].Canada - listData[1].Canada;
                double australiaCompareMonth = listData[0].Australia - listData[1].Australia;

                // tháng cùng kì năm trước
                double americanCompareMonthLastYear = listData[0].American - listData[2].American;
                double asiaCompareMonthLastYear = listData[0].Asia - listData[2].Asia;
                double globalCompareMonthLastYear = listData[0].Global - listData[2].Global;
                double europeCompareMonthLastYear = listData[0].Europe - listData[2].Europe;
                double canadaCompareMonthLastYear = listData[0].Canada - listData[2].Canada;
                double australiaCompareMonthLastYear = listData[0].Australia - listData[2].Australia;

                // add row vào table
                table.Rows.Add(str[0], listData[0].American, listData[1].American, listData[2].American
                    , listData[1].American == 0 ? 0 : Math.Round((americanCompareMonth / listData[1].American) * 100, 2, MidpointRounding.ToEven)
                    , listData[2].American == 0 ? 0 : Math.Round((americanCompareMonthLastYear / listData[2].American) * 100, 2, MidpointRounding.ToEven));

                table.Rows.Add(str[1], listData[0].Asia, listData[1].Asia, listData[2].Asia
                    , listData[1].Asia == 0 ? 0 : Math.Round((asiaCompareMonth / listData[1].Asia) * 100, 2, MidpointRounding.ToEven)
                    , listData[2].Asia == 0 ? 0 : Math.Round((asiaCompareMonthLastYear / listData[2].Asia) * 100, 2, MidpointRounding.ToEven));

                table.Rows.Add(str[2], listData[0].Global, listData[1].Global, listData[2].Global
                    , listData[1].Global == 0 ? 0 : Math.Round((globalCompareMonth / listData[1].Global) * 100, 2, MidpointRounding.ToEven)
                    , listData[2].Global == 0 ? 0 : Math.Round((globalCompareMonthLastYear / listData[2].Global) * 100, 2, MidpointRounding.ToEven));

                table.Rows.Add(str[3], listData[0].Europe, listData[1].Europe, listData[2].Europe
                    , listData[1].Europe == 0 ? 0 : Math.Round((europeCompareMonth / listData[1].Europe) * 100, 2, MidpointRounding.ToEven)
                    , listData[2].Europe == 0 ? 0 : Math.Round((europeCompareMonthLastYear / listData[2].Europe) * 100, 2, MidpointRounding.ToEven));

                table.Rows.Add(str[4], listData[0].Canada, listData[1].Canada, listData[2].Canada
                    , listData[1].Canada == 0 ? 0 : Math.Round((canadaCompareMonth / listData[1].Canada) * 100, 2, MidpointRounding.ToEven)
                    , listData[2].Canada == 0 ? 0 : Math.Round((canadaCompareMonthLastYear / listData[2].Canada) * 100, 2, MidpointRounding.ToEven));

                table.Rows.Add(str[5], listData[0].Australia, listData[1].Australia, listData[2].Australia
                    , listData[1].Australia == 0 ? 0 : Math.Round((australiaCompareMonth / listData[1].Australia) * 100, 2, MidpointRounding.ToEven)
                    , listData[2].Australia == 0 ? 0 : Math.Round((australiaCompareMonthLastYear / listData[2].Australia) * 100, 2, MidpointRounding.ToEven));

                DataRow row = table.NewRow();
                row["MaketID"] = "Tổng";
                row["AccumulateID1"] = 100;
                row["AccumulateID2"] = 100;
                row["AccumulateID3"] = 100;

                row["CompareToMonthPercent"] = 0;
                row["CompareToMonthLastYearPercent"] = 0;

                table.Rows.Add(row);
            }
            return Json(table.ToDataSourceResult(request), JsonRequestBehavior.AllowGet);
        }
    }
}
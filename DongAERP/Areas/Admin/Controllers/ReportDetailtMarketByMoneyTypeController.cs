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
    public class ReportDetailtMarketByMoneyTypeController : Controller
    {
        // GET: Admin/ReportDetailtMarketByMoneyType
        public ActionResult Index()
        {
            return View();
        }

        /// <summary>
        /// Báo cáo thị trường theo loại tiền chi trả tất cả
        /// </summary>
        /// <returns></returns>
        public ActionResult MarketForTotal()
        {
            string nameUrl = "Báo cáo chi tiết/Theo doanh số chi trả/Theo thị trường/Loại tiền chi trả/Tất cả";
            ViewBag.NameURL = nameUrl;
            return View();
        }

        /// <summary>
        /// Báo cáo thị trường theo loại tiền chi trả từng thị trường
        /// </summary>
        /// <returns></returns>
        public ActionResult MarketForOne()
        {
            string nameUrl = "Báo cáo chi tiết/Theo doanh số chi trả/Theo thị trường/Loại tiền chi trả/Từng thị trường";
            ViewBag.NameURL = nameUrl;
            return View();
        }

        public ActionResult ReportDetailtGradationCompare()
        {
            string nameUrl = "Báo cáo chi tiết/Theo doanh số chi trả/Theo thị trường/Loại tiền chi trả/So sánh - Giai đoạn - Tất cả thị trường";
            ViewBag.NameURL = nameUrl;
            // Khởi tạo datatable
            DataTable table = new DataTable();
            // Tạo các cột cho datatable
            table.Columns.Add("MarketName", typeof(String));
            table.Columns.Add("VND1", typeof(double));
            table.Columns.Add("USD1", typeof(double));
            table.Columns.Add("EUR1", typeof(double));
            table.Columns.Add("CAD1", typeof(double));
            table.Columns.Add("AUD1", typeof(double));
            table.Columns.Add("GBP1", typeof(double));
            table.Columns.Add("TDS1", typeof(double));

            table.Columns.Add("VND2", typeof(double));
            table.Columns.Add("USD2", typeof(double));
            table.Columns.Add("EUR2", typeof(double));
            table.Columns.Add("CAD2", typeof(double));
            table.Columns.Add("AUD2", typeof(double));
            table.Columns.Add("GBP2", typeof(double));
            table.Columns.Add("TDS2", typeof(double));
            return View(table);
        }


        public ActionResult ReportDetailtGradationCompareForOne()
        {
            string nameUrl = "Báo cáo chi tiết/Theo doanh số chi trả/Theo thị trường/Loại tiền chi trả/So sánh - Giai đoạn - Từng thị trường";
            ViewBag.NameURL = nameUrl;
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
            table.Columns.Add("TDS1", typeof(double));

            table.Columns.Add("VND2", typeof(double));
            table.Columns.Add("USD2", typeof(double));
            table.Columns.Add("EUR2", typeof(double));
            table.Columns.Add("CAD2", typeof(double));
            table.Columns.Add("AUD2", typeof(double));
            table.Columns.Add("GBP2", typeof(double));
            table.Columns.Add("TDS2", typeof(double));


            table.Columns.Add("MarketName", typeof(String));
            return View(table);
        }

        public ActionResult ReportDetailtCompareMonthForAll()
        {
            string nameUrl = "Báo cáo chi tiết/Theo doanh số chi trả/Theo thị trường/Loại tiền chi trả/So sánh - Theo tháng - tất cả thị trường";
            ViewBag.NameURL = nameUrl;
            // Khởi tạo datatable
            DataTable table = new DataTable();
            // Tạo các cột cho datatable
            table.Columns.Add("MarketName", typeof(String));
            table.Columns.Add("VND1", typeof(double));
            table.Columns.Add("USD1", typeof(double));
            table.Columns.Add("EUR1", typeof(double));
            table.Columns.Add("CAD1", typeof(double));
            table.Columns.Add("AUD1", typeof(double));
            table.Columns.Add("GBP1", typeof(double));
            table.Columns.Add("TDS1", typeof(double));

            // tháng trước
            table.Columns.Add("VND2", typeof(double));
            table.Columns.Add("USD2", typeof(double));
            table.Columns.Add("EUR2", typeof(double));
            table.Columns.Add("CAD2", typeof(double));
            table.Columns.Add("AUD2", typeof(double));
            table.Columns.Add("GBP2", typeof(double));
            table.Columns.Add("TDS2", typeof(double));

            // Năm hiện tại
            table.Columns.Add("VND3", typeof(double));
            table.Columns.Add("USD3", typeof(double));
            table.Columns.Add("EUR3", typeof(double));
            table.Columns.Add("CAD3", typeof(double));
            table.Columns.Add("AUD3", typeof(double));
            table.Columns.Add("GBP3", typeof(double));
            table.Columns.Add("TDS3", typeof(double));

            table.Columns.Add("MarketName", typeof(String));
            return View(table);
        }


        public ActionResult ReportDetailtCompareMonthForOne()
        {
            string nameUrl = "Báo cáo chi tiết/Theo doanh số chi trả/Theo thị trường/Loại tiền chi trả/So sánh - Theo tháng - Từng thị trường";
            ViewBag.NameURL = nameUrl;
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
            table.Columns.Add("TDS1", typeof(double));

            // tháng trước
            table.Columns.Add("VND2", typeof(double));
            table.Columns.Add("USD2", typeof(double));
            table.Columns.Add("EUR2", typeof(double));
            table.Columns.Add("CAD2", typeof(double));
            table.Columns.Add("AUD2", typeof(double));
            table.Columns.Add("GBP2", typeof(double));
            table.Columns.Add("TDS2", typeof(double));

            // Năm hiện tại
            table.Columns.Add("VND3", typeof(double));
            table.Columns.Add("USD3", typeof(double));
            table.Columns.Add("EUR3", typeof(double));
            table.Columns.Add("CAD3", typeof(double));
            table.Columns.Add("AUD3", typeof(double));
            table.Columns.Add("GBP3", typeof(double));
            table.Columns.Add("TDS3", typeof(double));

            table.Columns.Add("MarketName", typeof(String));
            return View(table);
        }

        /// <summary>
        /// Bảng hiển thị thông tin các giao dịch qua các ngày
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult MarketForTotal([DataSourceRequest]DataSourceRequest request, string reportTypeID)
        {
            List<ReportDetailtForTotalMoneyType> listData = new ReportBL().ListReportDetailtMTForAll(reportTypeID);
            int count = 1;
            foreach (ReportDetailtForTotalMoneyType item in listData)
            {
                item.STT = (count++).ToString();
                item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;
            }

            ReportDetailtForTotalMoneyType dataItem = new ReportDetailtForTotalMoneyType()
            {
                MarketName = "Tổng",
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
        public ActionResult SearchMarketForTotal([DataSourceRequest]DataSourceRequest request, DateTime fromDay, DateTime toDay, string reportTypeID, string marketID)
        {
            List<ReportDetailtForTotalMoneyType> listData = new ReportBL().SearchReportDetailtMTForAll(fromDay, toDay, reportTypeID, marketID);

            foreach (ReportDetailtForTotalMoneyType item in listData)
            {
                item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;
            }

            ReportDetailtForTotalMoneyType dataItem = new ReportDetailtForTotalMoneyType()
            {
                MarketName = "Tổng",
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
        /// Bảng hiển thị thông tin các giao dịch qua các ngày
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult MarketForTotalConvert([DataSourceRequest]DataSourceRequest request, string reportTypeID)
        {
            List<ReportDetailtForTotalMoneyType> listData = new ReportBL().ListReportDetailtMTForAllConvert(reportTypeID);
            int count = 1;
            foreach (ReportDetailtForTotalMoneyType item in listData)
            {
                item.STT = (count++).ToString();
                item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;
            }

            ReportDetailtForTotalMoneyType dataItem = new ReportDetailtForTotalMoneyType()
            {
                MarketName = "Tổng",
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
        public ActionResult SearchMarketForTotalConvert([DataSourceRequest]DataSourceRequest request, DateTime fromDay, DateTime toDay, string reportTypeID, string marketID)
        {
            List<ReportDetailtForTotalMoneyType> listData = new ReportBL().SearchReportDetailtMTForAllConvert(fromDay, toDay, reportTypeID, marketID);

            foreach (ReportDetailtForTotalMoneyType item in listData)
            {
                item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;
            }

            ReportDetailtForTotalMoneyType dataItem = new ReportDetailtForTotalMoneyType()
            {
                MarketName = "Tổng",
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
        /// Bảng hiển thị thông tin các giao dịch qua các ngày
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult MarketForOne([DataSourceRequest]DataSourceRequest request, string reportTypeID)
        {
            List<ReportDetailtForTotalMoneyType> listData = new ReportBL().ListReportDetailtMTForOne(reportTypeID);
            int count = 1;
            foreach (ReportDetailtForTotalMoneyType item in listData)
            {
                item.STT = (count++).ToString();
                item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;
            }

            return Json(listData.ToDataSourceResult(request), JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// Search report day theo ngày nhập vào
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public ActionResult SearchMarketForOne([DataSourceRequest]DataSourceRequest request, DateTime fromDay, DateTime toDay, string reportTypeID, string marketID)
        {
            List<ReportDetailtForTotalMoneyType> listData = new ReportBL().SearchReportDetailtMTForOne(fromDay, toDay, reportTypeID, marketID);
            int count = 1;
            foreach (ReportDetailtForTotalMoneyType item in listData)
            {
                item.STT = (count++).ToString();
                item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;
            }

            return Json(listData.ToDataSourceResult(request), JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// Bảng hiển thị thông tin các giao dịch qua các ngày
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult MarketForOneConvert([DataSourceRequest]DataSourceRequest request, string reportTypeID)
        {
            List<ReportDetailtForTotalMoneyType> listData = new ReportBL().ListReportDetailtMTForOneConvert(reportTypeID);
            int count = 1;
            foreach (ReportDetailtForTotalMoneyType item in listData)
            {
                item.STT = (count++).ToString();
                item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;
            }

            return Json(listData.ToDataSourceResult(request), JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// Search report day theo ngày nhập vào
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public ActionResult SearchMarketForOneConvert([DataSourceRequest]DataSourceRequest request, DateTime fromDay, DateTime toDay, string reportTypeID, string marketID)
        {
            List<ReportDetailtForTotalMoneyType> listData = new ReportBL().SearchReportDetailtMTForOneConvert(fromDay, toDay, reportTypeID, marketID);
            int count = 1;
            foreach (ReportDetailtForTotalMoneyType item in listData)
            {
                item.STT = (count++).ToString();
                item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;
            }
            return Json(listData.ToDataSourceResult(request), JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// Get data cho báo cáo so sánh giai đoạn
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult ReportDetailtGradationCompareForAllConvert([DataSourceRequest]DataSourceRequest request, string reportTypeID)
        {
            // Danh sach của data gradation gồm key và value

            string[] ArrayData = { "1", "3 tháng đầu năm" };

            int toYear = DateTime.Now.Year;
            List<ReportDetailtForTotalMoneyType> listDataGradation = new ReportBL().ReportDetailtMTGradationCompareForAllConvert(toYear, int.Parse(ArrayData[0]), reportTypeID);

            foreach (ReportDetailtForTotalMoneyType item in listDataGradation)
            {
                item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;
            }

            // Danh sách mã thị trường của Tất cả
            List<string> listMarket = new List<string>() { "003", "005", "001", "002", "014", "004" };

            // Khởi tạo datatable
            DataTable table = new DataTable();
            // Tạo các cột cho datatable
            table.Columns.Add("MarketName", typeof(String));
            table.Columns.Add("VND1", typeof(double));
            table.Columns.Add("USD1", typeof(double));
            table.Columns.Add("EUR1", typeof(double));
            table.Columns.Add("CAD1", typeof(double));
            table.Columns.Add("AUD1", typeof(double));
            table.Columns.Add("GBP1", typeof(double));
            table.Columns.Add("TDS1", typeof(double));

            table.Columns.Add("VND2", typeof(double));
            table.Columns.Add("USD2", typeof(double));
            table.Columns.Add("EUR2", typeof(double));
            table.Columns.Add("CAD2", typeof(double));
            table.Columns.Add("AUD2", typeof(double));
            table.Columns.Add("GBP2", typeof(double));
            table.Columns.Add("TDS2", typeof(double));

            foreach (string item in listMarket)
            {
                // Cùng kì
                ReportDetailtForTotalMoneyType dataItemLastYear = listDataGradation.Find(x => x.MarketCode == item && x.Year == (toYear - 1).ToString());
                ReportDetailtForTotalMoneyType dataItemYear = listDataGradation.Find(x => x.MarketCode == item && x.Year == toYear.ToString());

                // add item vào table
                table.Rows.Add(dataItemLastYear.MarketName
                    , dataItemLastYear.VND, dataItemLastYear.USD, dataItemLastYear.EUR, dataItemLastYear.CAD, dataItemLastYear.AUD, dataItemLastYear.GBP, dataItemLastYear.TongDS
                    , dataItemYear.VND, dataItemYear.USD, dataItemYear.EUR, dataItemYear.CAD, dataItemYear.AUD, dataItemYear.GBP, dataItemYear.TongDS
                    );
            }

            DataRow row = table.NewRow();
            row["MarketName"] = "Tổng";
            row["VND1"] = table.Compute("Sum(VND1)", "");
            row["USD1"] = table.Compute("Sum(USD1)", "");
            row["EUR1"] = table.Compute("Sum(EUR1)", "");
            row["CAD1"] = table.Compute("Sum(CAD1)", "");
            row["AUD1"] = table.Compute("Sum(AUD1)", "");
            row["GBP1"] = table.Compute("Sum(GBP1)", "");
            row["TDS1"] = table.Compute("Sum(TDS1)", "");

            row["VND2"] = table.Compute("Sum(VND2)", "");
            row["USD2"] = table.Compute("Sum(USD2)", "");
            row["EUR2"] = table.Compute("Sum(EUR2)", "");
            row["CAD2"] = table.Compute("Sum(CAD2)", "");
            row["AUD2"] = table.Compute("Sum(AUD2)", "");
            row["GBP2"] = table.Compute("Sum(GBP2)", "");
            row["TDS2"] = table.Compute("Sum(TDS2)", "");
            table.Rows.Add(row);

            //ViewBag.ListDataGradation = listDataGradation;
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
            List<ReportDetailtForTotalMoneyType> listDataGradation = new ReportBL().ReportDetailtMTGradationCompareForAllConvert(year, gradation, reportTypeID);

            foreach (ReportDetailtForTotalMoneyType item in listDataGradation)
            {
                item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;
            }

            // Danh sách mã thị trường của Tất cả
            List<string> listMarket = new List<string>() { "003", "005", "001", "002", "014", "004" };

            // Khởi tạo datatable
            DataTable table = new DataTable();
            // Tạo các cột cho datatable
            table.Columns.Add("MarketName", typeof(String));
            table.Columns.Add("VND1", typeof(double));
            table.Columns.Add("USD1", typeof(double));
            table.Columns.Add("EUR1", typeof(double));
            table.Columns.Add("CAD1", typeof(double));
            table.Columns.Add("AUD1", typeof(double));
            table.Columns.Add("GBP1", typeof(double));
            table.Columns.Add("TDS1", typeof(double));

            table.Columns.Add("VND2", typeof(double));
            table.Columns.Add("USD2", typeof(double));
            table.Columns.Add("EUR2", typeof(double));
            table.Columns.Add("CAD2", typeof(double));
            table.Columns.Add("AUD2", typeof(double));
            table.Columns.Add("GBP2", typeof(double));
            table.Columns.Add("TDS2", typeof(double));

            foreach (string item in listMarket)
            {
                // Cùng kì
                ReportDetailtForTotalMoneyType dataItemLastYear = listDataGradation.Find(x => x.MarketCode == item && x.Year == (year - 1).ToString());
                ReportDetailtForTotalMoneyType dataItemYear = listDataGradation.Find(x => x.MarketCode == item && x.Year == year.ToString());

                // add item vào table
                table.Rows.Add(dataItemLastYear.MarketName
                    , dataItemLastYear.VND, dataItemLastYear.USD, dataItemLastYear.EUR, dataItemLastYear.CAD, dataItemLastYear.AUD, dataItemLastYear.GBP, dataItemLastYear.TongDS
                    , dataItemYear.VND, dataItemYear.USD, dataItemYear.EUR, dataItemYear.CAD, dataItemYear.AUD, dataItemYear.GBP, dataItemYear.TongDS
                    );
            }

            DataRow row = table.NewRow();
            row["MarketName"] = "Tổng";
            row["VND1"] = table.Compute("Sum(VND1)", "");
            row["USD1"] = table.Compute("Sum(USD1)", "");
            row["EUR1"] = table.Compute("Sum(EUR1)", "");
            row["CAD1"] = table.Compute("Sum(CAD1)", "");
            row["AUD1"] = table.Compute("Sum(AUD1)", "");
            row["GBP1"] = table.Compute("Sum(GBP1)", "");
            row["TDS1"] = table.Compute("Sum(TDS1)", "");

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
        /// Get data cho báo cáo so sánh giai đoạn
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult GradationCompareForAllConvertCompare([DataSourceRequest]DataSourceRequest request, string reportTypeID)
        {
            // Danh sach của data gradation gồm key và value

            string[] ArrayData = { "1", "3 tháng đầu năm" };

            int toYear = DateTime.Now.Year;
            List<ReportDetailtForTotalMoneyType> listDataGradation = new ReportBL().ReportDetailtMTGradationCompareForAllConvert(toYear, int.Parse(ArrayData[0]), reportTypeID);

            foreach (ReportDetailtForTotalMoneyType item in listDataGradation)
            {
                item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;
            }

            // Danh sách mã thị trường của Tất cả
            List<string> listMarket = new List<string>() { "003", "005", "001", "002", "014", "004" };

            // Khởi tạo datatable
            DataTable table = new DataTable();
            // Tạo các cột cho datatable
            table.Columns.Add("MarketName", typeof(String));
            table.Columns.Add("VND1", typeof(double));
            table.Columns.Add("USD1", typeof(double));
            table.Columns.Add("EUR1", typeof(double));
            table.Columns.Add("CAD1", typeof(double));
            table.Columns.Add("AUD1", typeof(double));
            table.Columns.Add("GBP1", typeof(double));
            table.Columns.Add("TDS1", typeof(double));

            foreach (string item in listMarket)
            {
                // Cùng kì
                ReportDetailtForTotalMoneyType dataItemLastYear = listDataGradation.Find(x => x.MarketCode == item && x.Year == (toYear - 1).ToString());
                ReportDetailtForTotalMoneyType dataItemYear = listDataGradation.Find(x => x.MarketCode == item && x.Year == toYear.ToString());

                // Trường hợp năm trước có đối tác và năm nay không có
                if (dataItemLastYear != null && dataItemYear == null)
                {
                    dataItemYear = new ReportDetailtForTotalMoneyType();
                    dataItemYear.MarketCode = dataItemLastYear.MarketCode;
                    dataItemYear.MarketName = dataItemLastYear.MarketName;
                    dataItemYear.VND = 0;
                    dataItemYear.USD = 0;
                    dataItemYear.EUR = 0;
                    dataItemYear.CAD = 0;
                    dataItemYear.AUD = 0;
                    dataItemYear.GBP = 0;
                    dataItemYear.Year = toYear.ToString();
                }

                // Trường hợp năm trước không có đối tác và năm nay có đối tác
                if (dataItemYear != null && dataItemLastYear == null)
                {
                    dataItemLastYear = new ReportDetailtForTotalMoneyType();
                    dataItemLastYear.MarketCode = dataItemYear.MarketCode;
                    dataItemLastYear.MarketName = dataItemYear.MarketName;
                    dataItemLastYear.VND = 0;
                    dataItemLastYear.USD = 0;
                    dataItemLastYear.EUR = 0;
                    dataItemLastYear.CAD = 0;
                    dataItemLastYear.AUD = 0;
                    dataItemLastYear.GBP = 0;
                    dataItemLastYear.Year = (toYear - 1).ToString();
                }

                double sumVND = dataItemYear.VND - dataItemLastYear.VND;
                double sumUSD = dataItemYear.USD - dataItemLastYear.USD;
                double sumEUR = dataItemYear.EUR - dataItemLastYear.EUR;
                double sumCAD = dataItemYear.CAD - dataItemLastYear.CAD;
                double sumAUD = dataItemYear.AUD - dataItemLastYear.AUD;
                double sumGBP = dataItemYear.GBP - dataItemLastYear.GBP;
                double sumTongDS = sumVND + sumUSD + sumEUR + sumCAD + sumAUD + sumGBP;

                // add item vào table
                table.Rows.Add(dataItemLastYear.MarketName
                    , Math.Round(sumVND, 2, MidpointRounding.ToEven)
                    , Math.Round(sumUSD, 2, MidpointRounding.ToEven)
                    , Math.Round(sumEUR, 2, MidpointRounding.ToEven)
                    , Math.Round(sumCAD, 2, MidpointRounding.ToEven)
                    , Math.Round(sumAUD, 2, MidpointRounding.ToEven)
                    , Math.Round(sumGBP, 2, MidpointRounding.ToEven)
                    , Math.Round(sumTongDS, 2, MidpointRounding.ToEven)
                    );
            }

            DataRow row = table.NewRow();
            row["MarketName"] = "Tổng";
            row["VND1"] = table.Compute("Sum(VND1)", "");
            row["USD1"] = table.Compute("Sum(USD1)", "");
            row["EUR1"] = table.Compute("Sum(EUR1)", "");
            row["CAD1"] = table.Compute("Sum(CAD1)", "");
            row["AUD1"] = table.Compute("Sum(AUD1)", "");
            row["GBP1"] = table.Compute("Sum(GBP1)", "");
            row["TDS1"] = table.Compute("Sum(TDS1)", "");
            table.Rows.Add(row);

            //ViewBag.ListDataGradation = listDataGradation;
            return Json(table.ToDataSourceResult(request), JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// Get data cho việc vẽ biểu đồ cột chồng cho so sánh giai đoạn- Tất cả - Loại tiền chi trả
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult ColumnChartGradationCompareStack(string reportTypeID)
        {
            DateTime now = DateTime.Now;
            // Với mặt định chọn giai đoạn 3 tháng đầu năm
            int gradationID = 1;
            List<ReportDetailtForTotalMoneyType> listDataGradation = new ReportBL().ReportDetailtMTGradationCompareForAllPercent(now.Year, gradationID, reportTypeID);

            // Số mảng cần tạo
            int arrayCount = 6;
            GradationCharColumn[] arrayGradation = new GradationCharColumn[listDataGradation.Count * arrayCount];
            int count = 0;
            // group theo nhóm
            int tooltip = 1;
            foreach (ReportDetailtForTotalMoneyType item in listDataGradation)
            {
                // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                arrayGradation[count] = new GradationCharColumn()
                {
                    Serie = "VND",
                    Segmento = string.Format("{0} {1}", item.MarketName, item.Year),
                    Valor1 = item.VND,
                    Tooltip = tooltip
                };

                count++;
                arrayGradation[count] = new GradationCharColumn()
                {
                    Serie = "USD",
                    Segmento = string.Format("{0} {1}", item.MarketName, item.Year),
                    Valor1 = item.USD,
                    Tooltip = tooltip
                };

                count++;
                arrayGradation[count] = new GradationCharColumn()
                {
                    Serie = "EUR",
                    Segmento = string.Format("{0} {1}", item.MarketName, item.Year),
                    Valor1 = item.EUR,
                    Tooltip = tooltip
                };
                count++;

                arrayGradation[count] = new GradationCharColumn()
                {
                    Serie = "CAD",
                    Segmento = string.Format("{0} {1}", item.MarketName, item.Year),
                    Valor1 = item.CAD,
                    Tooltip = tooltip
                };
                count++;

                arrayGradation[count] = new GradationCharColumn()
                {
                    Serie = "AUD",
                    Segmento = string.Format("{0} {1}", item.MarketName, item.Year),
                    Valor1 = item.AUD,
                    Tooltip = tooltip
                };
                count++;

                arrayGradation[count] = new GradationCharColumn()
                {
                    Serie = "GBP",
                    Segmento = string.Format("{0} {1}", item.MarketName, item.Year),
                    Valor1 = item.GBP,
                    Tooltip = tooltip
                };
                count++;
                tooltip++;
            }

            if (listDataGradation.Count == 0)
            {
                arrayGradation = new GradationCharColumn[1];
                arrayGradation[0] = new GradationCharColumn()
                {
                    Serie = "1",
                    Segmento = "1",
                    Valor1 = 0,
                    Tooltip = 1

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
        public ActionResult SearchColumnChartReportForGradationPercent([DataSourceRequest]DataSourceRequest request, int gradation, int year, string reportTypeID)
        {
            List<ReportDetailtForTotalMoneyType> listDataGradation = new ReportBL().ReportDetailtMTGradationCompareForAllPercent(year, gradation, reportTypeID);

            // Số mảng cần tạo
            int arrayCount = 6;
            GradationCharColumn[] arrayGradation = new GradationCharColumn[listDataGradation.Count * arrayCount];
            int count = 0;
            // group theo nhóm
            int tooltip = 1;
            foreach (ReportDetailtForTotalMoneyType item in listDataGradation)
            {
                // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                arrayGradation[count] = new GradationCharColumn()
                {
                    Serie = "VND",
                    Segmento = string.Format("{0} {1}", item.MarketName, item.Year),
                    Valor1 = item.VND,
                    Tooltip = tooltip
                };

                count++;
                arrayGradation[count] = new GradationCharColumn()
                {
                    Serie = "USD",
                    Segmento = string.Format("{0} {1}", item.MarketName, item.Year),
                    Valor1 = item.USD,
                    Tooltip = tooltip
                };

                count++;
                arrayGradation[count] = new GradationCharColumn()
                {
                    Serie = "EUR",
                    Segmento = string.Format("{0} {1}", item.MarketName, item.Year),
                    Valor1 = item.EUR,
                    Tooltip = tooltip
                };
                count++;

                arrayGradation[count] = new GradationCharColumn()
                {
                    Serie = "CAD",
                    Segmento = string.Format("{0} {1}", item.MarketName, item.Year),
                    Valor1 = item.CAD,
                    Tooltip = tooltip
                };
                count++;

                arrayGradation[count] = new GradationCharColumn()
                {
                    Serie = "AUD",
                    Segmento = string.Format("{0} {1}", item.MarketName, item.Year),
                    Valor1 = item.AUD,
                    Tooltip = tooltip
                };
                count++;

                arrayGradation[count] = new GradationCharColumn()
                {
                    Serie = "GBP",
                    Segmento = string.Format("{0} {1}", item.MarketName, item.Year),
                    Valor1 = item.GBP,
                    Tooltip = tooltip
                };
                count++;
                tooltip++;
            }

            if (listDataGradation.Count == 0)
            {
                arrayGradation = new GradationCharColumn[1];
                arrayGradation[0] = new GradationCharColumn()
                {
                    Serie = "1",
                    Segmento = "1",
                    Valor1 = 0,
                    Tooltip = 1

                };
            }

            return Json(arrayGradation);
        }


        /// <summary>
        /// Get data cho việc vẽ biểu đồ cột cho so sánh giai đoạn- Tất cả thị trường
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult ColumnsChartGradationCompare(string reportTypeID)
        {
            int gradationID = 1;
            int year = DateTime.Today.Year;
            List<ReportDetailtForTotalMoneyType> listDataGradation = new ReportBL().ReportDetailtMTGradationCompareForAllConvert(year, gradationID, reportTypeID);

            // Số record của mảng
            int countArray = 6;
            GradationCompare[] arrayGradation = new GradationCompare[countArray * listDataGradation.Count];
            int count = 0;
            foreach (ReportDetailtForTotalMoneyType item in listDataGradation)
            {
                // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                arrayGradation[count] = new GradationCompare()
                {
                    NameGradationCompare = "VND",
                    amount = item.VND,
                    NameType = string.Format("{0} {1}", item.MarketName, item.Year),
                };

                count++;
                arrayGradation[count] = new GradationCompare()
                {
                    NameGradationCompare = "USD",
                    amount = item.USD,
                    NameType = string.Format("{0} {1}", item.MarketName, item.Year),
                };

                count++;
                arrayGradation[count] = new GradationCompare()
                {
                    NameGradationCompare = "EUR",
                    amount = item.EUR,
                    NameType = string.Format("{0} {1}", item.MarketName, item.Year),
                };

                count++;
                arrayGradation[count] = new GradationCompare()
                {
                    NameGradationCompare = "CAD",
                    amount = item.CAD,
                    NameType = string.Format("{0} {1}", item.MarketName, item.Year),
                };

                count++;
                arrayGradation[count] = new GradationCompare()
                {
                    NameGradationCompare = "AUD",
                    amount = item.AUD,
                    NameType = string.Format("{0} {1}", item.MarketName, item.Year),
                };

                count++;
                arrayGradation[count] = new GradationCompare()
                {
                    NameGradationCompare = "GBP",
                    amount = item.GBP,
                    NameType = string.Format("{0} {1}", item.MarketName, item.Year),
                };

                // Tăng count lên 1 đơn vị
                count++;
                year = year - 1;
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
        /// Get data cho việc vẽ biểu đồ cột cho so sánh theo giai đoạn - Tất cẳ thị trường
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult SearchColumnChartReportForGradation([DataSourceRequest]DataSourceRequest request, int gradation, int year, string reportTypeID)
        {
            List<ReportDetailtForTotalMoneyType> listDataGradation = new ReportBL().ReportDetailtMTGradationCompareForAllConvert(year, gradation, reportTypeID);

            // Số record của mảng
            int countArray = 6;
            GradationCompare[] arrayGradation = new GradationCompare[countArray * listDataGradation.Count];
            int count = 0;
            foreach (ReportDetailtForTotalMoneyType item in listDataGradation)
            {
                // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                arrayGradation[count] = new GradationCompare()
                {
                    NameGradationCompare = "VND",
                    amount = item.VND,
                    NameType = string.Format("{0} {1}", item.MarketName, item.Year),
                };

                count++;
                arrayGradation[count] = new GradationCompare()
                {
                    NameGradationCompare = "USD",
                    amount = item.USD,
                    NameType = string.Format("{0} {1}", item.MarketName, item.Year),
                };

                count++;
                arrayGradation[count] = new GradationCompare()
                {
                    NameGradationCompare = "EUR",
                    amount = item.EUR,
                    NameType = string.Format("{0} {1}", item.MarketName, item.Year),
                };

                count++;
                arrayGradation[count] = new GradationCompare()
                {
                    NameGradationCompare = "CAD",
                    amount = item.CAD,
                    NameType = string.Format("{0} {1}", item.MarketName, item.Year),
                };

                count++;
                arrayGradation[count] = new GradationCompare()
                {
                    NameGradationCompare = "AUD",
                    amount = item.AUD,
                    NameType = string.Format("{0} {1}", item.MarketName, item.Year),
                };

                count++;
                arrayGradation[count] = new GradationCompare()
                {
                    NameGradationCompare = "GBP",
                    amount = item.GBP,
                    NameType = string.Format("{0} {1}", item.MarketName, item.Year),
                };

                // Tăng count lên 1 đơn vị
                count++;
                year = year - 1;
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
        /// Get data cho báo cáo so sánh giai đoạn theo loại tiền chi trả- từng thị trường
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult ReportDetailtGradationCompareForOneConvert([DataSourceRequest]DataSourceRequest request, string reportTypeID)
        {
            // Danh sach của data gradation gồm key và value

            string[] ArrayData = { "1", "3 tháng đầu năm" };

            int toYear = DateTime.Now.Year;
            // Giá trị ban đầu
            string marketID = "001";
            List<ReportDetailtForTotalMoneyType> listDataGradation = new ReportBL().ReportDetailtMTGradationCompareForOneConvert(toYear, int.Parse(ArrayData[0]), reportTypeID, marketID);

            foreach (ReportDetailtForTotalMoneyType item in listDataGradation)
            {
                item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;
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
            table.Columns.Add("TDS1", typeof(double));

            table.Columns.Add("VND2", typeof(double));
            table.Columns.Add("USD2", typeof(double));
            table.Columns.Add("EUR2", typeof(double));
            table.Columns.Add("CAD2", typeof(double));
            table.Columns.Add("AUD2", typeof(double));
            table.Columns.Add("GBP2", typeof(double));
            table.Columns.Add("TDS2", typeof(double));


            table.Columns.Add("MarketName", typeof(String));


            foreach (ReportDetailtForTotalMoneyType item in listDataGradation)
            {

                // Cùng kì
                ReportDetailtForTotalMoneyType dataItemLastYear = listDataGradation.Find(x => x.PartnerCode == item.PartnerCode && x.Year == (toYear - 1).ToString());
                ReportDetailtForTotalMoneyType dataItemYear = listDataGradation.Find(x => x.PartnerCode == item.PartnerCode && x.Year == toYear.ToString());

                // Trường hợp năm trước có đối tác và năm nay không có
                if (dataItemLastYear != null && dataItemYear == null)
                {
                    dataItemYear = new ReportDetailtForTotalMoneyType();
                    dataItemYear.PartnerCode = dataItemLastYear.PartnerCode;
                    dataItemYear.PartnerName = dataItemLastYear.PartnerName;
                    dataItemYear.Year = toYear.ToString();
                }

                // Trường hợp năm trước không có đối tác và năm nay có đối tác
                if (dataItemYear != null && dataItemLastYear == null)
                {
                    dataItemLastYear = new ReportDetailtForTotalMoneyType();
                    dataItemLastYear.PartnerCode = dataItemYear.PartnerCode;
                    dataItemLastYear.PartnerName = dataItemYear.PartnerName;
                    dataItemLastYear.Year = (toYear - 1).ToString();
                }

                // Check tồn tại của item
                string value = string.Format("PartnerName='{0}'", item.PartnerName);
                DataRow[] foundRows = table.Select(value);
                if (dataItemLastYear != null && dataItemYear != null && foundRows.Count() == 0)
                {
                    // add item vào table
                    table.Rows.Add(item.PartnerName
                        , dataItemLastYear.VND, dataItemLastYear.USD, dataItemLastYear.EUR, dataItemLastYear.CAD, dataItemLastYear.AUD, dataItemLastYear.GBP, dataItemLastYear.TongDS
                        , dataItemYear.VND, dataItemYear.USD, dataItemYear.EUR, dataItemYear.CAD, dataItemYear.AUD, dataItemYear.GBP, dataItemYear.TongDS
                        , item.MarketName);
                }
            }


            DataRow row = table.NewRow();
            row["PartnerName"] = "Tổng";
            row["VND1"] = table.Compute("Sum(VND1)", "");
            row["USD1"] = table.Compute("Sum(USD1)", "");
            row["EUR1"] = table.Compute("Sum(EUR1)", "");
            row["CAD1"] = table.Compute("Sum(CAD1)", "");
            row["AUD1"] = table.Compute("Sum(AUD1)", "");
            row["GBP1"] = table.Compute("Sum(GBP1)", "");
            row["TDS1"] = table.Compute("Sum(TDS1)", "");

            row["VND2"] = table.Compute("Sum(VND2)", "");
            row["USD2"] = table.Compute("Sum(USD2)", "");
            row["EUR2"] = table.Compute("Sum(EUR2)", "");
            row["CAD2"] = table.Compute("Sum(CAD2)", "");
            row["AUD2"] = table.Compute("Sum(AUD2)", "");
            row["GBP2"] = table.Compute("Sum(GBP2)", "");
            row["TDS2"] = table.Compute("Sum(TDS2)", "");


            row["MarketName"] = "Châu Âu";
            table.Rows.Add(row);

            //ViewBag.ListDataGradation = listDataGradation;
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
        public ActionResult SearchGridReportForGradationForOne([DataSourceRequest]DataSourceRequest request, int gradation, int year, string reportTypeID, string marketID)
        {
            List<ReportDetailtForTotalMoneyType> listDataGradation = new ReportBL().ReportDetailtMTGradationCompareForOneConvert(year, gradation, reportTypeID, marketID);

            foreach (ReportDetailtForTotalMoneyType item in listDataGradation)
            {
                item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;
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
            table.Columns.Add("TDS1", typeof(double));

            table.Columns.Add("VND2", typeof(double));
            table.Columns.Add("USD2", typeof(double));
            table.Columns.Add("EUR2", typeof(double));
            table.Columns.Add("CAD2", typeof(double));
            table.Columns.Add("AUD2", typeof(double));
            table.Columns.Add("GBP2", typeof(double));
            table.Columns.Add("TDS2", typeof(double));

            table.Columns.Add("MarketName", typeof(String));

            foreach (ReportDetailtForTotalMoneyType item in listDataGradation)
            {

                // Cùng kì
                ReportDetailtForTotalMoneyType dataItemLastYear = listDataGradation.Find(x => x.PartnerCode == item.PartnerCode && x.Year == (year - 1).ToString());
                ReportDetailtForTotalMoneyType dataItemYear = listDataGradation.Find(x => x.PartnerCode == item.PartnerCode && x.Year == year.ToString());

                // Trường hợp năm trước có đối tác và năm nay không có
                if (dataItemLastYear != null && dataItemYear == null)
                {
                    dataItemYear = new ReportDetailtForTotalMoneyType();
                    dataItemYear.PartnerCode = dataItemLastYear.PartnerCode;
                    dataItemYear.PartnerName = dataItemLastYear.PartnerName;
                    dataItemYear.Year = year.ToString();
                }

                // Trường hợp năm trước không có đối tác và năm nay có đối tác
                if (dataItemYear != null && dataItemLastYear == null)
                {
                    dataItemLastYear = new ReportDetailtForTotalMoneyType();
                    dataItemLastYear.PartnerCode = dataItemYear.PartnerCode;
                    dataItemLastYear.PartnerName = dataItemYear.PartnerName;
                    dataItemLastYear.Year = (year - 1).ToString();
                }

                // Check tồn tại của item
                string value = string.Format("PartnerName='{0}'", item.PartnerName);
                DataRow[] foundRows = table.Select(value);
                if (dataItemLastYear != null && dataItemYear != null && foundRows.Count() == 0)
                {
                    // add item vào table
                    table.Rows.Add(item.PartnerName
                        , dataItemLastYear.VND, dataItemLastYear.USD, dataItemLastYear.EUR, dataItemLastYear.CAD, dataItemLastYear.AUD, dataItemLastYear.GBP, dataItemLastYear.TongDS
                        , dataItemYear.VND, dataItemYear.USD, dataItemYear.EUR, dataItemYear.CAD, dataItemYear.AUD, dataItemYear.GBP, dataItemYear.TongDS
                        , item.MarketName
                        );
                }
            }


            //DataRow row = table.NewRow();
            //row["PartnerName"] = "Tổng";
            //row["VND1"] = table.Compute("Sum(VND1)", "");
            //row["USD1"] = table.Compute("Sum(USD1)", "");
            //row["EUR1"] = table.Compute("Sum(EUR1)", "");
            //row["CAD1"] = table.Compute("Sum(CAD1)", "");
            //row["AUD1"] = table.Compute("Sum(AUD1)", "");
            //row["GBP1"] = table.Compute("Sum(GBP1)", "");
            //row["TDS1"] = table.Compute("Sum(TDS1)", "");

            //row["VND2"] = table.Compute("Sum(VND2)", "");
            //row["USD2"] = table.Compute("Sum(USD2)", "");
            //row["EUR2"] = table.Compute("Sum(EUR2)", "");
            //row["CAD2"] = table.Compute("Sum(CAD2)", "");
            //row["AUD2"] = table.Compute("Sum(AUD2)", "");
            //row["GBP2"] = table.Compute("Sum(GBP2)", "");
            //row["TDS2"] = table.Compute("Sum(TDS2)", "");

            //row["MarketName"] = "";
            //table.Rows.Add(row);

            return Json(table.ToDataSourceResult(request), JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// Get data cho báo cáo so sánh giai đoạn theo loại tiền chi trả- từng thị trường
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult ReportDetailtGradationCompareForOneConvertCompare([DataSourceRequest]DataSourceRequest request, string reportTypeID)
        {
            // Danh sach của data gradation gồm key và value

            string[] ArrayData = { "1", "3 tháng đầu năm" };

            int toYear = DateTime.Now.Year;
            // Giá trị ban đầu
            string marketID = "001";
            List<ReportDetailtForTotalMoneyType> listDataGradation = new ReportBL().ReportDetailtMTGradationCompareForOneConvert(toYear, int.Parse(ArrayData[0]), reportTypeID, marketID);

            foreach (ReportDetailtForTotalMoneyType item in listDataGradation)
            {
                item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;
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
            table.Columns.Add("TDS1", typeof(double));

            table.Columns.Add("MarketName", typeof(string));

            foreach (ReportDetailtForTotalMoneyType item in listDataGradation)
            {

                // Cùng kì
                ReportDetailtForTotalMoneyType dataItemLastYear = listDataGradation.Find(x => x.PartnerCode == item.PartnerCode && x.Year == (toYear - 1).ToString());
                ReportDetailtForTotalMoneyType dataItemYear = listDataGradation.Find(x => x.PartnerCode == item.PartnerCode && x.Year == toYear.ToString());

                // Trường hợp năm trước có đối tác và năm nay không có
                if (dataItemLastYear != null && dataItemYear == null)
                {
                    dataItemYear = new ReportDetailtForTotalMoneyType();
                    dataItemYear.PartnerCode = dataItemLastYear.PartnerCode;
                    dataItemYear.PartnerName = dataItemLastYear.PartnerName;
                    dataItemYear.Year = toYear.ToString();
                }

                // Trường hợp năm trước không có đối tác và năm nay có đối tác
                if (dataItemYear != null && dataItemLastYear == null)
                {
                    dataItemLastYear = new ReportDetailtForTotalMoneyType();
                    dataItemLastYear.PartnerCode = dataItemYear.PartnerCode;
                    dataItemLastYear.PartnerName = dataItemYear.PartnerName;
                    dataItemLastYear.Year = (toYear - 1).ToString();
                }

                double sumVND = dataItemYear.VND - dataItemLastYear.VND;
                double sumUSD = dataItemYear.USD - dataItemLastYear.USD;
                double sumEUR = dataItemYear.EUR - dataItemLastYear.EUR;
                double sumCAD = dataItemYear.CAD - dataItemLastYear.CAD;
                double sumAUD = dataItemYear.AUD - dataItemLastYear.AUD;
                double sumGBP = dataItemYear.GBP - dataItemLastYear.GBP;
                double sumTongDS = sumVND + sumUSD + sumEUR + sumCAD + sumAUD + sumGBP;

                // Check tồn tại của item
                string value = string.Format("PartnerName='{0}'", item.PartnerName);
                DataRow[] foundRows = table.Select(value);
                if (dataItemLastYear != null && dataItemYear != null && foundRows.Count() == 0)
                {
                    // add item vào table
                    table.Rows.Add(item.PartnerName
                        , Math.Round(sumVND, 2, MidpointRounding.ToEven)
                        , Math.Round(sumUSD, 2, MidpointRounding.ToEven)
                        , Math.Round(sumEUR, 2, MidpointRounding.ToEven)
                        , Math.Round(sumCAD, 2, MidpointRounding.ToEven)
                        , Math.Round(sumAUD, 2, MidpointRounding.ToEven)
                        , Math.Round(sumGBP, 2, MidpointRounding.ToEven)
                        , Math.Round(sumTongDS, 2, MidpointRounding.ToEven)
                        , item.MarketName
                        );
                }
            }


            //DataRow row = table.NewRow();
            //row["PartnerName"] = "Tổng";
            //row["VND1"] = table.Compute("Sum(VND1)", "");
            //row["USD1"] = table.Compute("Sum(USD1)", "");
            //row["EUR1"] = table.Compute("Sum(EUR1)", "");
            //row["CAD1"] = table.Compute("Sum(CAD1)", "");
            //row["AUD1"] = table.Compute("Sum(AUD1)", "");
            //row["GBP1"] = table.Compute("Sum(GBP1)", "");
            //row["TDS1"] = table.Compute("Sum(TDS1)", "");

            //row["MarketName"] = "";
            //table.Rows.Add(row);

            //ViewBag.ListDataGradation = listDataGradation;
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
        public ActionResult SearchGridReportForGradationForOneCompare([DataSourceRequest]DataSourceRequest request, int gradation, int year, string reportTypeID, string marketID)
        {
            List<ReportDetailtForTotalMoneyType> listDataGradation = new ReportBL().ReportDetailtMTGradationCompareForOneConvert(year, gradation, reportTypeID, marketID);

            foreach (ReportDetailtForTotalMoneyType item in listDataGradation)
            {
                item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;
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
            table.Columns.Add("TDS1", typeof(double));

            table.Columns.Add("MarketName", typeof(string));

            foreach (ReportDetailtForTotalMoneyType item in listDataGradation)
            {

                // Cùng kì
                ReportDetailtForTotalMoneyType dataItemLastYear = listDataGradation.Find(x => x.PartnerCode == item.PartnerCode && x.Year == (year - 1).ToString());
                ReportDetailtForTotalMoneyType dataItemYear = listDataGradation.Find(x => x.PartnerCode == item.PartnerCode && x.Year == year.ToString());

                // Trường hợp năm trước có đối tác và năm nay không có
                if (dataItemLastYear != null && dataItemYear == null)
                {
                    dataItemYear = new ReportDetailtForTotalMoneyType();
                    dataItemYear.PartnerCode = dataItemLastYear.PartnerCode;
                    dataItemYear.PartnerName = dataItemLastYear.PartnerName;
                    dataItemYear.Year = year.ToString();
                }

                // Trường hợp năm trước không có đối tác và năm nay có đối tác
                if (dataItemYear != null && dataItemLastYear == null)
                {
                    dataItemLastYear = new ReportDetailtForTotalMoneyType();
                    dataItemLastYear.PartnerCode = dataItemYear.PartnerCode;
                    dataItemLastYear.PartnerName = dataItemYear.PartnerName;
                    dataItemLastYear.Year = (year - 1).ToString();
                }

                double sumVND = dataItemYear.VND - dataItemLastYear.VND;
                double sumUSD = dataItemYear.USD - dataItemLastYear.USD;
                double sumEUR = dataItemYear.EUR - dataItemLastYear.EUR;
                double sumCAD = dataItemYear.CAD - dataItemLastYear.CAD;
                double sumAUD = dataItemYear.AUD - dataItemLastYear.AUD;
                double sumGBP = dataItemYear.GBP - dataItemLastYear.GBP;
                double sumTongDS = sumVND + sumUSD + sumEUR + sumCAD + sumAUD + sumGBP;

                // Check tồn tại của item
                string value = string.Format("PartnerName='{0}'", item.PartnerName);
                DataRow[] foundRows = table.Select(value);
                if (dataItemLastYear != null && dataItemYear != null && foundRows.Count() == 0)
                {
                    // add item vào table
                    table.Rows.Add(item.PartnerName
                        , Math.Round(sumVND, 2, MidpointRounding.ToEven)
                        , Math.Round(sumUSD, 2, MidpointRounding.ToEven)
                        , Math.Round(sumEUR, 2, MidpointRounding.ToEven)
                        , Math.Round(sumCAD, 2, MidpointRounding.ToEven)
                        , Math.Round(sumAUD, 2, MidpointRounding.ToEven)
                        , Math.Round(sumGBP, 2, MidpointRounding.ToEven)
                        , Math.Round(sumTongDS, 2, MidpointRounding.ToEven)
                        , item.MarketName
                        );
                }
            }


            //DataRow row = table.NewRow();
            //row["PartnerName"] = "Tổng";
            //row["VND1"] = table.Compute("Sum(VND1)", "");
            //row["USD1"] = table.Compute("Sum(USD1)", "");
            //row["EUR1"] = table.Compute("Sum(EUR1)", "");
            //row["CAD1"] = table.Compute("Sum(CAD1)", "");
            //row["AUD1"] = table.Compute("Sum(AUD1)", "");
            //row["GBP1"] = table.Compute("Sum(GBP1)", "");
            //row["TDS1"] = table.Compute("Sum(TDS1)", "");

            //table.Rows.Add(row);

            //ViewBag.ListDataGradation = listDataGradation;
            return Json(table.ToDataSourceResult(request), JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// Get data cho việc vẽ biểu đồ cột chồng cho so sánh giai đoạn
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult ColumnChartGradationCompareStackForOne(string reportTypeID)
        {
            int year = DateTime.Now.Year;
            // Với mặt định chọn giai đoạn 3 tháng đầu năm
            int gradationID = 1;
            // Giá trị ban đầu
            string marketID = "001";
            List<ReportDetailtForTotalMoneyType> listDataGradation = new ReportBL().ReportDetailtMTGradationCompareForOnePercent(year, gradationID, reportTypeID, marketID);
            List<ReportDetailtForTotalMoneyType> listDataGradationClone = new List<ReportDetailtForTotalMoneyType>(listDataGradation);

            foreach (ReportDetailtForTotalMoneyType item in listDataGradationClone)
            {
                ReportDetailtForTotalMoneyType dataItemYear = listDataGradation.Find(x => x.PartnerCode == item.PartnerCode && x.Year == year.ToString());
                ReportDetailtForTotalMoneyType dataItemLastYear = listDataGradation.Find(x => x.PartnerCode == item.PartnerCode && x.Year == (year - 1).ToString());

                // Trường hợp năm trước có đối tác và năm nay không có
                if (dataItemLastYear != null && dataItemYear == null)
                {
                    dataItemYear = new ReportDetailtForTotalMoneyType();
                    dataItemYear.PartnerCode = dataItemLastYear.PartnerCode;
                    dataItemYear.PartnerName = dataItemLastYear.PartnerName;
                    dataItemYear.Year = year.ToString();

                    // Add item thiếu vào
                    listDataGradation.Add(dataItemYear);
                }

                // Trường hợp năm trước không có đối tác và năm nay có đối tác
                if (dataItemYear != null && dataItemLastYear == null)
                {
                    dataItemLastYear = new ReportDetailtForTotalMoneyType();
                    dataItemLastYear.PartnerCode = dataItemYear.PartnerCode;
                    dataItemLastYear.PartnerName = dataItemYear.PartnerName;
                    dataItemLastYear.Year = (year - 1).ToString();

                    // Add item thiếu vào
                    listDataGradation.Add(dataItemLastYear);
                }
            }


            // Số mảng cần tạo
            int arrayCount = 6;
            GradationCharColumn[] arrayGradation = new GradationCharColumn[listDataGradation.Count * arrayCount];
            int count = 0;
            foreach (ReportDetailtForTotalMoneyType item in listDataGradation)
            {
                // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                arrayGradation[count] = new GradationCharColumn()
                {
                    Serie = "VND",
                    Segmento = string.Format("{0} {1}", item.PartnerName, item.Year),
                    Valor1 = item.VND,
                    //Tooltip = tooltip
                };
                count++;

                arrayGradation[count] = new GradationCharColumn()
                {
                    Serie = "USD",
                    Segmento = string.Format("{0} {1}", item.PartnerName, item.Year),
                    Valor1 = item.USD,
                    //Tooltip = tooltip
                };
                count++;

                arrayGradation[count] = new GradationCharColumn()
                {
                    Serie = "EUR",
                    Segmento = string.Format("{0} {1}", item.PartnerName, item.Year),
                    Valor1 = item.EUR,
                    //Tooltip = tooltip
                };
                count++;

                arrayGradation[count] = new GradationCharColumn()
                {
                    Serie = "CAD",
                    Segmento = string.Format("{0} {1}", item.PartnerName, item.Year),
                    Valor1 = item.CAD,
                    //Tooltip = tooltip
                };
                count++;

                arrayGradation[count] = new GradationCharColumn()
                {
                    Serie = "AUD",
                    Segmento = string.Format("{0} {1}", item.PartnerName, item.Year),
                    Valor1 = item.AUD,
                    //Tooltip = tooltip
                };
                count++;

                arrayGradation[count] = new GradationCharColumn()
                {
                    Serie = "GBP",
                    Segmento = string.Format("{0} {1}", item.PartnerName, item.Year),
                    Valor1 = item.GBP,
                    //Tooltip = tooltip
                };
                count++;
                //tooltip++;
            }

            if (listDataGradation.Count == 0)
            {
                arrayGradation = new GradationCharColumn[1];
                arrayGradation[0] = new GradationCharColumn()
                {
                    Serie = "1",
                    Segmento = "1",
                    Valor1 = 0,
                    Tooltip = 1

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
        public ActionResult SearchColumnChartGradationCompareStackForOne([DataSourceRequest]DataSourceRequest request, int gradation, int year, string reportTypeID, string marketID)
        {
            List<ReportDetailtForTotalMoneyType> listDataGradation = new ReportBL().ReportDetailtMTGradationCompareForOnePercent(year, gradation, reportTypeID, marketID);
            List<ReportDetailtForTotalMoneyType> listDataGradationClone = new List<ReportDetailtForTotalMoneyType>(listDataGradation);

            foreach (ReportDetailtForTotalMoneyType item in listDataGradationClone)
            {
                ReportDetailtForTotalMoneyType dataItemYear = listDataGradation.Find(x => x.PartnerCode == item.PartnerCode && x.Year == year.ToString());
                ReportDetailtForTotalMoneyType dataItemLastYear = listDataGradation.Find(x => x.PartnerCode == item.PartnerCode && x.Year == (year - 1).ToString());

                // Trường hợp năm trước có đối tác và năm nay không có
                if (dataItemLastYear != null && dataItemYear == null)
                {
                    dataItemYear = new ReportDetailtForTotalMoneyType();
                    dataItemYear.PartnerCode = dataItemLastYear.PartnerCode;
                    dataItemYear.PartnerName = dataItemLastYear.PartnerName;
                    dataItemYear.Year = year.ToString();

                    // Add item thiếu vào
                    listDataGradation.Add(dataItemYear);
                }

                // Trường hợp năm trước không có đối tác và năm nay có đối tác
                if (dataItemYear != null && dataItemLastYear == null)
                {
                    dataItemLastYear = new ReportDetailtForTotalMoneyType();
                    dataItemLastYear.PartnerCode = dataItemYear.PartnerCode;
                    dataItemLastYear.PartnerName = dataItemYear.PartnerName;
                    dataItemLastYear.Year = (year - 1).ToString();

                    // Add item thiếu vào
                    listDataGradation.Add(dataItemLastYear);
                }
            }

            // Số mảng cần tạo
            int arrayCount = 6;
            GradationCharColumn[] arrayGradation = new GradationCharColumn[listDataGradation.Count * arrayCount];
            int count = 0;
            foreach (ReportDetailtForTotalMoneyType item in listDataGradation)
            {
                // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                arrayGradation[count] = new GradationCharColumn()
                {
                    Serie = "VND",
                    Segmento = string.Format("{0} {1}", item.PartnerName, item.Year),
                    Valor1 = item.VND,
                    //Tooltip = tooltip
                };
                count++;

                arrayGradation[count] = new GradationCharColumn()
                {
                    Serie = "USD",
                    Segmento = string.Format("{0} {1}", item.PartnerName, item.Year),
                    Valor1 = item.USD,
                    //Tooltip = tooltip
                };
                count++;

                arrayGradation[count] = new GradationCharColumn()
                {
                    Serie = "EUR",
                    Segmento = string.Format("{0} {1}", item.PartnerName, item.Year),
                    Valor1 = item.EUR,
                    //Tooltip = tooltip
                };
                count++;

                arrayGradation[count] = new GradationCharColumn()
                {
                    Serie = "CAD",
                    Segmento = string.Format("{0} {1}", item.PartnerName, item.Year),
                    Valor1 = item.CAD,
                    //Tooltip = tooltip
                };
                count++;

                arrayGradation[count] = new GradationCharColumn()
                {
                    Serie = "AUD",
                    Segmento = string.Format("{0} {1}", item.PartnerName, item.Year),
                    Valor1 = item.AUD,
                    //Tooltip = tooltip
                };
                count++;

                arrayGradation[count] = new GradationCharColumn()
                {
                    Serie = "GBP",
                    Segmento = string.Format("{0} {1}", item.PartnerName, item.Year),
                    Valor1 = item.GBP,
                    //Tooltip = tooltip
                };
                count++;
                //tooltip++;
            }

            if (listDataGradation.Count == 0)
            {
                arrayGradation = new GradationCharColumn[1];
                arrayGradation[0] = new GradationCharColumn()
                {
                    Serie = "1",
                    Segmento = "1",
                    Valor1 = 0,
                    Tooltip = 1

                };
            }

            return Json(arrayGradation);
        }

        /// <summary>
        /// Get data cho việc vẽ biểu đồ cột cho so sánh giai đoạn
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult ColumnsChartGradationCompareForOne(string reportTypeID)
        {
            int gradationID = 1;
            int year = DateTime.Today.Year;
            string marketID = "001";
            List<ReportDetailtForTotalMoneyType> listDataGradation = new ReportBL().ReportDetailtMTGradationCompareForOneConvert(year, gradationID, reportTypeID, marketID);

            // Số record của mảng
            int countArray = 6;
            GradationCompare[] arrayGradation = new GradationCompare[countArray * listDataGradation.Count];
            int count = 0;
            foreach (ReportDetailtForTotalMoneyType item in listDataGradation)
            {
                // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                arrayGradation[count] = new GradationCompare()
                {
                    NameGradationCompare = "VND",
                    amount = item.VND,
                    NameType = string.Format("{0} {1}", item.PartnerName, item.Year),
                };

                count++;
                arrayGradation[count] = new GradationCompare()
                {
                    NameGradationCompare = "USD",
                    amount = item.USD,
                    NameType = string.Format("{0} {1}", item.PartnerName, item.Year),
                };

                count++;
                arrayGradation[count] = new GradationCompare()
                {
                    NameGradationCompare = "EUR",
                    amount = item.EUR,
                    NameType = string.Format("{0} {1}", item.PartnerName, item.Year),
                };

                count++;
                arrayGradation[count] = new GradationCompare()
                {
                    NameGradationCompare = "CAD",
                    amount = item.CAD,
                    NameType = string.Format("{0} {1}", item.PartnerName, item.Year),
                };

                count++;
                arrayGradation[count] = new GradationCompare()
                {
                    NameGradationCompare = "AUD",
                    amount = item.AUD,
                    NameType = string.Format("{0} {1}", item.PartnerName, item.Year),
                };

                count++;
                arrayGradation[count] = new GradationCompare()
                {
                    NameGradationCompare = "GBP",
                    amount = item.GBP,
                    NameType = string.Format("{0} {1}", item.PartnerName, item.Year),
                };

                // Tăng count lên 1 đơn vị
                count++;
                year = year - 1;
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
        /// Get data cho việc vẽ biểu đồ cột cho so sánh theo giai đoạn
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult SearchColumnsChartGradationCompareForOne([DataSourceRequest]DataSourceRequest request, int gradation, int year, string reportTypeID, string marketID)
        {
            List<ReportDetailtForTotalMoneyType> listDataGradation = new ReportBL().ReportDetailtMTGradationCompareForOneConvert(year, gradation, reportTypeID, marketID);

            // Số record của mảng
            int countArray = 6;
            GradationCompare[] arrayGradation = new GradationCompare[countArray * listDataGradation.Count];
            int count = 0;
            foreach (ReportDetailtForTotalMoneyType item in listDataGradation)
            {
                // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                arrayGradation[count] = new GradationCompare()
                {
                    NameGradationCompare = "VND",
                    amount = item.VND,
                    NameType = string.Format("{0} {1}", item.PartnerName, item.Year),
                };

                count++;
                arrayGradation[count] = new GradationCompare()
                {
                    NameGradationCompare = "USD",
                    amount = item.USD,
                    NameType = string.Format("{0} {1}", item.PartnerName, item.Year),
                };

                count++;
                arrayGradation[count] = new GradationCompare()
                {
                    NameGradationCompare = "EUR",
                    amount = item.EUR,
                    NameType = string.Format("{0} {1}", item.PartnerName, item.Year),
                };

                count++;
                arrayGradation[count] = new GradationCompare()
                {
                    NameGradationCompare = "CAD",
                    amount = item.CAD,
                    NameType = string.Format("{0} {1}", item.PartnerName, item.Year),
                };

                count++;
                arrayGradation[count] = new GradationCompare()
                {
                    NameGradationCompare = "AUD",
                    amount = item.AUD,
                    NameType = string.Format("{0} {1}", item.PartnerName, item.Year),
                };

                count++;
                arrayGradation[count] = new GradationCompare()
                {
                    NameGradationCompare = "GBP",
                    amount = item.GBP,
                    NameType = string.Format("{0} {1}", item.PartnerName, item.Year),
                };

                // Tăng count lên 1 đơn vị
                count++;
                year = year - 1;
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
        /// Get data cho báo cáo so sánh giai đoạn theo loại tiền chi trả
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult ReportDetailtCompareMonthForAll([DataSourceRequest]DataSourceRequest request, string reportTypeID)
        {
            // Danh sach của data gradation gồm key và value
            int toYear = DateTime.Now.Year;
            int toMonth = DateTime.Now.Month;

            List<ReportDetailtForTotalMoneyType> listDataCompareMonth = new ReportBL().ReportDetailtMTCompareMonthForAllConvert(toYear, toMonth, reportTypeID);

            foreach (ReportDetailtForTotalMoneyType item in listDataCompareMonth)
            {
                item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;
            }

            // Danh sách mã thị trường của Tất cả
            List<string> listMarket = new List<string>() { "003", "005", "001", "002", "014", "004" };

            // Khởi tạo datatable
            DataTable table = new DataTable();
            // Tạo các cột cho datatable
            // Cùng kì năm trước
            table.Columns.Add("MarketName", typeof(String));
            table.Columns.Add("VND1", typeof(double));
            table.Columns.Add("USD1", typeof(double));
            table.Columns.Add("EUR1", typeof(double));
            table.Columns.Add("CAD1", typeof(double));
            table.Columns.Add("AUD1", typeof(double));
            table.Columns.Add("GBP1", typeof(double));
            table.Columns.Add("TDS1", typeof(double));

            // tháng trước
            table.Columns.Add("VND2", typeof(double));
            table.Columns.Add("USD2", typeof(double));
            table.Columns.Add("EUR2", typeof(double));
            table.Columns.Add("CAD2", typeof(double));
            table.Columns.Add("AUD2", typeof(double));
            table.Columns.Add("GBP2", typeof(double));
            table.Columns.Add("TDS2", typeof(double));

            // Năm hiện tại
            table.Columns.Add("VND3", typeof(double));
            table.Columns.Add("USD3", typeof(double));
            table.Columns.Add("EUR3", typeof(double));
            table.Columns.Add("CAD3", typeof(double));
            table.Columns.Add("AUD3", typeof(double));
            table.Columns.Add("GBP3", typeof(double));
            table.Columns.Add("TDS3", typeof(double));

            foreach (string item in listMarket)
            {
                // Cùng kì
                ReportDetailtForTotalMoneyType dataItemLastYear = listDataCompareMonth.Find(x => x.MarketCode == item && x.Month == toMonth.ToString() && x.Year == (toYear - 1).ToString());
                ReportDetailtForTotalMoneyType dataItemYear = listDataCompareMonth.Find(x => x.MarketCode == item && x.Month == toMonth.ToString() && x.Year == toYear.ToString());
                ReportDetailtForTotalMoneyType dataItemLastMonth = listDataCompareMonth.Find(x => x.MarketCode == item && x.Month == (toMonth - 1).ToString() && x.Year == toYear.ToString());

                if (dataItemLastYear != null && dataItemYear != null && dataItemLastMonth != null)
                {
                    // add item vào table
                    table.Rows.Add(dataItemYear.MarketName
                        , dataItemLastYear.VND, dataItemLastYear.USD, dataItemLastYear.EUR, dataItemLastYear.CAD, dataItemLastYear.AUD, dataItemLastYear.GBP, dataItemLastYear.TongDS
                        , dataItemLastMonth.VND, dataItemLastMonth.USD, dataItemLastMonth.EUR, dataItemLastMonth.CAD, dataItemLastMonth.AUD, dataItemLastMonth.GBP, dataItemLastMonth.TongDS
                        , dataItemYear.VND, dataItemYear.USD, dataItemYear.EUR, dataItemYear.CAD, dataItemYear.AUD, dataItemYear.GBP, dataItemYear.TongDS
                        );
                }
            }

            DataRow row = table.NewRow();
            row["MarketName"] = "Tổng";
            row["VND1"] = table.Compute("Sum(VND1)", "");
            row["USD1"] = table.Compute("Sum(USD1)", "");
            row["EUR1"] = table.Compute("Sum(EUR1)", "");
            row["CAD1"] = table.Compute("Sum(CAD1)", "");
            row["AUD1"] = table.Compute("Sum(AUD1)", "");
            row["GBP1"] = table.Compute("Sum(GBP1)", "");
            row["TDS1"] = table.Compute("Sum(TDS1)", "");

            row["VND2"] = table.Compute("Sum(VND2)", "");
            row["USD2"] = table.Compute("Sum(USD2)", "");
            row["EUR2"] = table.Compute("Sum(EUR2)", "");
            row["CAD2"] = table.Compute("Sum(CAD2)", "");
            row["AUD2"] = table.Compute("Sum(AUD2)", "");
            row["GBP2"] = table.Compute("Sum(GBP2)", "");
            row["TDS2"] = table.Compute("Sum(TDS2)", "");

            row["VND3"] = table.Compute("Sum(VND3)", "");
            row["USD3"] = table.Compute("Sum(USD3)", "");
            row["EUR3"] = table.Compute("Sum(EUR3)", "");
            row["CAD3"] = table.Compute("Sum(CAD3)", "");
            row["AUD3"] = table.Compute("Sum(AUD3)", "");
            row["GBP3"] = table.Compute("Sum(GBP3)", "");
            row["TDS3"] = table.Compute("Sum(TDS3)", "");

            table.Rows.Add(row);

            //ViewBag.ListDataGradation = listDataGradation;
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
            List<ReportDetailtForTotalMoneyType> listDataCompareMonth = new ReportBL().ReportDetailtMTCompareMonthForAllConvert(year, month, reportTypeID);

            foreach (ReportDetailtForTotalMoneyType item in listDataCompareMonth)
            {
                item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;
            }

            // Danh sách mã thị trường của Tất cả
            List<string> listMarket = new List<string>() { "003", "005", "001", "002", "014", "004" };

            // Khởi tạo datatable
            DataTable table = new DataTable();
            // Tạo các cột cho datatable
            // Cùng kì năm trước
            table.Columns.Add("MarketName", typeof(String));
            table.Columns.Add("VND1", typeof(double));
            table.Columns.Add("USD1", typeof(double));
            table.Columns.Add("EUR1", typeof(double));
            table.Columns.Add("CAD1", typeof(double));
            table.Columns.Add("AUD1", typeof(double));
            table.Columns.Add("GBP1", typeof(double));
            table.Columns.Add("TDS1", typeof(double));

            // tháng trước
            table.Columns.Add("VND2", typeof(double));
            table.Columns.Add("USD2", typeof(double));
            table.Columns.Add("EUR2", typeof(double));
            table.Columns.Add("CAD2", typeof(double));
            table.Columns.Add("AUD2", typeof(double));
            table.Columns.Add("GBP2", typeof(double));
            table.Columns.Add("TDS2", typeof(double));

            // Năm hiện tại
            table.Columns.Add("VND3", typeof(double));
            table.Columns.Add("USD3", typeof(double));
            table.Columns.Add("EUR3", typeof(double));
            table.Columns.Add("CAD3", typeof(double));
            table.Columns.Add("AUD3", typeof(double));
            table.Columns.Add("GBP3", typeof(double));
            table.Columns.Add("TDS3", typeof(double));

            foreach (string item in listMarket)
            {
                // Cùng kì
                ReportDetailtForTotalMoneyType dataItemLastYear = listDataCompareMonth.Find(x => x.MarketCode == item && x.Month == month.ToString() && x.Year == (year - 1).ToString());
                ReportDetailtForTotalMoneyType dataItemYear = listDataCompareMonth.Find(x => x.MarketCode == item && x.Month == month.ToString() && x.Year == year.ToString());
                ReportDetailtForTotalMoneyType dataItemLastMonth = listDataCompareMonth.Find(x => x.MarketCode == item && x.Month == (month - 1).ToString() && x.Year == year.ToString());

                if (dataItemLastYear != null && dataItemYear != null && dataItemLastMonth != null)
                {
                    // add item vào table
                    table.Rows.Add(dataItemYear.MarketName
                        , dataItemLastYear.VND, dataItemLastYear.USD, dataItemLastYear.EUR, dataItemLastYear.CAD, dataItemLastYear.AUD, dataItemLastYear.GBP, dataItemLastYear.TongDS
                        , dataItemLastMonth.VND, dataItemLastMonth.USD, dataItemLastMonth.EUR, dataItemLastMonth.CAD, dataItemLastMonth.AUD, dataItemLastMonth.GBP, dataItemLastMonth.TongDS
                        , dataItemYear.VND, dataItemYear.USD, dataItemYear.EUR, dataItemYear.CAD, dataItemYear.AUD, dataItemYear.GBP, dataItemYear.TongDS
                        );
                }
            }

            DataRow row = table.NewRow();
            row["MarketName"] = "Tổng";
            row["VND1"] = table.Compute("Sum(VND1)", "");
            row["USD1"] = table.Compute("Sum(USD1)", "");
            row["EUR1"] = table.Compute("Sum(EUR1)", "");
            row["CAD1"] = table.Compute("Sum(CAD1)", "");
            row["AUD1"] = table.Compute("Sum(AUD1)", "");
            row["GBP1"] = table.Compute("Sum(GBP1)", "");
            row["TDS1"] = table.Compute("Sum(TDS1)", "");

            row["VND2"] = table.Compute("Sum(VND2)", "");
            row["USD2"] = table.Compute("Sum(USD2)", "");
            row["EUR2"] = table.Compute("Sum(EUR2)", "");
            row["CAD2"] = table.Compute("Sum(CAD2)", "");
            row["AUD2"] = table.Compute("Sum(AUD2)", "");
            row["GBP2"] = table.Compute("Sum(GBP2)", "");
            row["TDS2"] = table.Compute("Sum(TDS2)", "");

            row["VND3"] = table.Compute("Sum(VND3)", "");
            row["USD3"] = table.Compute("Sum(USD3)", "");
            row["EUR3"] = table.Compute("Sum(EUR3)", "");
            row["CAD3"] = table.Compute("Sum(CAD3)", "");
            row["AUD3"] = table.Compute("Sum(AUD3)", "");
            row["GBP3"] = table.Compute("Sum(GBP3)", "");
            row["TDS3"] = table.Compute("Sum(TDS3)", "");

            table.Rows.Add(row);

            //ViewBag.ListDataGradation = listDataGradation;
            return Json(table.ToDataSourceResult(request), JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// Get data cho báo cáo chi tiết theo tháng
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult ReportDetailtCompareMonthForAllCompare([DataSourceRequest]DataSourceRequest request, string reportTypeID)
        {
            int toYear = DateTime.Now.Year;
            int toMonth = DateTime.Now.Month;

            List<ReportDetailtForTotalMoneyType> listDataCompareMonth = new ReportBL().ReportDetailtMTCompareMonthForAllConvert(toYear, toMonth, reportTypeID);

            foreach (ReportDetailtForTotalMoneyType item in listDataCompareMonth)
            {
                item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;
            }

            // Khởi tạo datatable
            DataTable table = new DataTable();
            // Tạo các cột cho datatable
            table.Columns.Add("MarketName", typeof(String));

            // So sánh với tháng trước
            table.Columns.Add("VND1", typeof(double));
            table.Columns.Add("USD1", typeof(double));
            table.Columns.Add("EUR1", typeof(double));
            table.Columns.Add("CAD1", typeof(double));
            table.Columns.Add("AUD1", typeof(double));
            table.Columns.Add("GBP1", typeof(double));
            table.Columns.Add("TDS1", typeof(double));

            // So sánh với cùng kì năm trước
            table.Columns.Add("VND2", typeof(double));
            table.Columns.Add("USD2", typeof(double));
            table.Columns.Add("EUR2", typeof(double));
            table.Columns.Add("CAD2", typeof(double));
            table.Columns.Add("AUD2", typeof(double));
            table.Columns.Add("GBP2", typeof(double));
            table.Columns.Add("TDS2", typeof(double));

            foreach (ReportDetailtForTotalMoneyType item in listDataCompareMonth)
            {
                // Cùng kì
                ReportDetailtForTotalMoneyType dataItemLastYear = listDataCompareMonth.Find(x => x.MarketCode == item.MarketCode && x.Month == toMonth.ToString() && x.Year == (toYear - 1).ToString());
                ReportDetailtForTotalMoneyType dataItemYear = listDataCompareMonth.Find(x => x.MarketCode == item.MarketCode && x.Month == toMonth.ToString() && x.Year == toYear.ToString());
                ReportDetailtForTotalMoneyType dataItemLastMonth = listDataCompareMonth.Find(x => x.MarketCode == item.MarketCode && x.Month == (toMonth - 1).ToString() && x.Year == toYear.ToString());

                // Cùng kì năm trước
                if (dataItemLastYear == null)
                {
                    dataItemLastYear.MarketCode = item.MarketCode;
                    dataItemLastYear.MarketName = item.MarketName;
                }

                // Tháng hiện tại
                if (dataItemYear == null)
                {
                    dataItemYear.MarketCode = item.MarketCode;
                    dataItemYear.MarketName = item.MarketName;
                }

                // Tháng trước
                if (dataItemLastMonth == null)
                {
                    dataItemLastMonth.MarketCode = item.MarketCode;
                    dataItemLastMonth.MarketName = item.MarketName;
                }

                // So với cùng kì năm trước
                double sumVNDLastYear = Math.Round(dataItemYear.VND - dataItemLastYear.VND, 2, MidpointRounding.ToEven);
                double sumUSDLastYear = Math.Round(dataItemYear.USD - dataItemLastYear.USD, 2, MidpointRounding.ToEven);
                double sumEURLastYear = Math.Round(dataItemYear.EUR - dataItemLastYear.EUR, 2, MidpointRounding.ToEven);
                double sumCADLastYear = Math.Round(dataItemYear.CAD - dataItemLastYear.CAD, 2, MidpointRounding.ToEven);
                double sumAUDLastYear = Math.Round(dataItemYear.AUD - dataItemLastYear.AUD, 2, MidpointRounding.ToEven);
                double sumGBPLastYear = Math.Round(dataItemYear.GBP - dataItemLastYear.GBP, 2, MidpointRounding.ToEven);

                double sumTDSLastYear = sumVNDLastYear + sumUSDLastYear + sumEURLastYear + sumCADLastYear + sumAUDLastYear + sumGBPLastYear;

                // so với tháng trước
                double sumVNDLastMonth = Math.Round(dataItemYear.VND - dataItemLastMonth.VND, 2, MidpointRounding.ToEven);
                double sumUSDLastMonth = Math.Round(dataItemYear.USD - dataItemLastMonth.USD, 2, MidpointRounding.ToEven);
                double sumEURLastMonth = Math.Round(dataItemYear.EUR - dataItemLastMonth.EUR, 2, MidpointRounding.ToEven);
                double sumCADLastMonth = Math.Round(dataItemYear.CAD - dataItemLastMonth.CAD, 2, MidpointRounding.ToEven);
                double sumAUDLastMonth = Math.Round(dataItemYear.AUD - dataItemLastMonth.AUD, 2, MidpointRounding.ToEven);
                double sumGBPLastMonth = Math.Round(dataItemYear.GBP - dataItemLastMonth.GBP, 2, MidpointRounding.ToEven);

                double sumTDSLastMonth = sumVNDLastMonth + sumUSDLastMonth + sumEURLastMonth + sumCADLastMonth + sumAUDLastMonth + sumGBPLastMonth;

                // Check tồn tại của item
                string value = string.Format("MarketName='{0}'", item.MarketName);
                DataRow[] foundRows = table.Select(value);

                if (dataItemLastYear != null && dataItemYear != null && dataItemLastMonth != null && foundRows.Count() == 0)
                {
                    // add item vào table
                    table.Rows.Add(dataItemYear.MarketName
                        , sumVNDLastMonth, sumUSDLastMonth, sumEURLastMonth, sumCADLastMonth, sumAUDLastMonth, sumGBPLastMonth
                        , Math.Round(sumTDSLastMonth, 2, MidpointRounding.ToEven)
                        , sumVNDLastYear, sumUSDLastYear, sumEURLastYear, sumCADLastYear, sumAUDLastYear, sumGBPLastYear
                        , Math.Round(sumTDSLastYear, 2, MidpointRounding.ToEven)
                        );
                }
            }

            DataRow row = table.NewRow();
            row["MarketName"] = "Tổng";
            row["VND1"] = table.Compute("Sum(VND1)", "");
            row["USD1"] = table.Compute("Sum(USD1)", "");
            row["EUR1"] = table.Compute("Sum(EUR1)", "");
            row["CAD1"] = table.Compute("Sum(CAD1)", "");
            row["AUD1"] = table.Compute("Sum(AUD1)", "");
            row["GBP1"] = table.Compute("Sum(GBP1)", "");
            row["TDS1"] = table.Compute("Sum(TDS1)", "");

            row["VND2"] = table.Compute("Sum(VND2)", "");
            row["USD2"] = table.Compute("Sum(USD2)", "");
            row["EUR2"] = table.Compute("Sum(EUR2)", "");
            row["CAD2"] = table.Compute("Sum(CAD2)", "");
            row["AUD2"] = table.Compute("Sum(AUD2)", "");
            row["GBP2"] = table.Compute("Sum(GBP2)", "");
            row["TDS2"] = table.Compute("Sum(TDS2)", "");

            table.Rows.Add(row);

            //ViewBag.ListDataGradation = listDataGradation;
            return Json(table.ToDataSourceResult(request), JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// Get data cho báo cáo chi tiết theo tháng
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult SearchReportDetailtCompareMonthForAllCompare([DataSourceRequest]DataSourceRequest request, int month, int year, string reportTypeID)
        {
            List<ReportDetailtForTotalMoneyType> listDataCompareMonth = new ReportBL().ReportDetailtMTCompareMonthForAllConvert(year, month, reportTypeID);

            foreach (ReportDetailtForTotalMoneyType item in listDataCompareMonth)
            {
                item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;
            }

            // Khởi tạo datatable
            DataTable table = new DataTable();
            // Tạo các cột cho datatable
            table.Columns.Add("MarketName", typeof(String));

            // So sánh với tháng trước
            table.Columns.Add("VND1", typeof(double));
            table.Columns.Add("USD1", typeof(double));
            table.Columns.Add("EUR1", typeof(double));
            table.Columns.Add("CAD1", typeof(double));
            table.Columns.Add("AUD1", typeof(double));
            table.Columns.Add("GBP1", typeof(double));
            table.Columns.Add("TDS1", typeof(double));

            // So sánh với cùng kì năm trước
            table.Columns.Add("VND2", typeof(double));
            table.Columns.Add("USD2", typeof(double));
            table.Columns.Add("EUR2", typeof(double));
            table.Columns.Add("CAD2", typeof(double));
            table.Columns.Add("AUD2", typeof(double));
            table.Columns.Add("GBP2", typeof(double));
            table.Columns.Add("TDS2", typeof(double));

            foreach (ReportDetailtForTotalMoneyType item in listDataCompareMonth)
            {
                // Cùng kì
                ReportDetailtForTotalMoneyType dataItemLastYear = listDataCompareMonth.Find(x => x.MarketCode == item.MarketCode && x.Month == month.ToString() && x.Year == (year - 1).ToString());
                ReportDetailtForTotalMoneyType dataItemYear = listDataCompareMonth.Find(x => x.MarketCode == item.MarketCode && x.Month == month.ToString() && x.Year == year.ToString());
                ReportDetailtForTotalMoneyType dataItemLastMonth = listDataCompareMonth.Find(x => x.MarketCode == item.MarketCode && x.Month == (month - 1).ToString() && x.Year == year.ToString());

                // Cùng kì năm trước
                if (dataItemLastYear == null)
                {
                    dataItemLastYear.MarketCode = item.MarketCode;
                    dataItemLastYear.MarketName = item.MarketName;
                }

                // Tháng hiện tại
                if (dataItemYear == null)
                {
                    dataItemYear.MarketCode = item.MarketCode;
                    dataItemYear.MarketName = item.MarketName;
                }

                // Tháng trước
                if (dataItemLastMonth == null)
                {
                    dataItemLastMonth.MarketCode = item.MarketCode;
                    dataItemLastMonth.MarketName = item.MarketName;
                }

                // So với cùng kì năm trước
                double sumVNDLastYear = Math.Round(dataItemYear.VND - dataItemLastYear.VND, 2, MidpointRounding.ToEven);
                double sumUSDLastYear = Math.Round(dataItemYear.USD - dataItemLastYear.USD, 2, MidpointRounding.ToEven);
                double sumEURLastYear = Math.Round(dataItemYear.EUR - dataItemLastYear.EUR, 2, MidpointRounding.ToEven);
                double sumCADLastYear = Math.Round(dataItemYear.CAD - dataItemLastYear.CAD, 2, MidpointRounding.ToEven);
                double sumAUDLastYear = Math.Round(dataItemYear.AUD - dataItemLastYear.AUD, 2, MidpointRounding.ToEven);
                double sumGBPLastYear = Math.Round(dataItemYear.GBP - dataItemLastYear.GBP, 2, MidpointRounding.ToEven);

                double sumTDSLastYear = sumVNDLastYear + sumUSDLastYear + sumEURLastYear + sumCADLastYear + sumAUDLastYear + sumGBPLastYear;

                // so với tháng trước
                double sumVNDLastMonth = Math.Round(dataItemYear.VND - dataItemLastMonth.VND, 2, MidpointRounding.ToEven);
                double sumUSDLastMonth = Math.Round(dataItemYear.USD - dataItemLastMonth.USD, 2, MidpointRounding.ToEven);
                double sumEURLastMonth = Math.Round(dataItemYear.EUR - dataItemLastMonth.EUR, 2, MidpointRounding.ToEven);
                double sumCADLastMonth = Math.Round(dataItemYear.CAD - dataItemLastMonth.CAD, 2, MidpointRounding.ToEven);
                double sumAUDLastMonth = Math.Round(dataItemYear.AUD - dataItemLastMonth.AUD, 2, MidpointRounding.ToEven);
                double sumGBPLastMonth = Math.Round(dataItemYear.GBP - dataItemLastMonth.GBP, 2, MidpointRounding.ToEven);

                double sumTDSLastMonth = sumVNDLastMonth + sumUSDLastMonth + sumEURLastMonth + sumCADLastMonth + sumAUDLastMonth + sumGBPLastMonth;

                // Check tồn tại của item
                string value = string.Format("MarketName='{0}'", item.MarketName);
                DataRow[] foundRows = table.Select(value);

                if (dataItemLastYear != null && dataItemYear != null && dataItemLastMonth != null && foundRows.Count() == 0)
                {
                    // add item vào table
                    table.Rows.Add(dataItemYear.MarketName
                        , sumVNDLastMonth, sumUSDLastMonth, sumEURLastMonth, sumCADLastMonth, sumAUDLastMonth, sumGBPLastMonth
                        , Math.Round(sumTDSLastMonth, 2, MidpointRounding.ToEven)
                        , sumVNDLastYear, sumUSDLastYear, sumEURLastYear, sumCADLastYear, sumAUDLastYear, sumGBPLastYear
                        , Math.Round(sumTDSLastYear, 2, MidpointRounding.ToEven)
                        );
                }
            }

            DataRow row = table.NewRow();
            row["MarketName"] = "Tổng";
            row["VND1"] = table.Compute("Sum(VND1)", "");
            row["USD1"] = table.Compute("Sum(USD1)", "");
            row["EUR1"] = table.Compute("Sum(EUR1)", "");
            row["CAD1"] = table.Compute("Sum(CAD1)", "");
            row["AUD1"] = table.Compute("Sum(AUD1)", "");
            row["GBP1"] = table.Compute("Sum(GBP1)", "");
            row["TDS1"] = table.Compute("Sum(TDS1)", "");

            row["VND2"] = table.Compute("Sum(VND2)", "");
            row["USD2"] = table.Compute("Sum(USD2)", "");
            row["EUR2"] = table.Compute("Sum(EUR2)", "");
            row["CAD2"] = table.Compute("Sum(CAD2)", "");
            row["AUD2"] = table.Compute("Sum(AUD2)", "");
            row["GBP2"] = table.Compute("Sum(GBP2)", "");
            row["TDS2"] = table.Compute("Sum(TDS2)", "");

            table.Rows.Add(row);

            //ViewBag.ListDataGradation = listDataGradation;
            return Json(table.ToDataSourceResult(request), JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// Get data cho việc vẽ biểu đồ cột chồng cho so sánh giai đoạn
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult ColumnChartStackCompareMonthForAll(string reportTypeID)
        {
            int year = DateTime.Now.Year;
            int toMonth = DateTime.Now.Month;

            List<ReportDetailtForTotalMoneyType> listDataCompareMonth = new ReportBL().ColumnChartStackDetailtMTCompareMonthForAllPercent(year, toMonth, reportTypeID);

            GradationCharColumn[] arrayGradation = null;

            if (listDataCompareMonth.Count > 0)
            {
                List<ReportDetailtForTotalMoneyType> listDataYear = listDataCompareMonth.Where(x => x.Month == toMonth.ToString() && x.Year == year.ToString()).ToList();
                List<ReportDetailtForTotalMoneyType> listDataLastMonth = listDataCompareMonth.Where(x => x.Month == (toMonth - 1).ToString() && x.Year == year.ToString()).ToList();
                List<ReportDetailtForTotalMoneyType> listDataLastYear = listDataCompareMonth.Where(x => x.Month == toMonth.ToString() && x.Year == (year - 1).ToString()).ToList();

                // Trường hợp đủ dữ liệu tháng hiện tại, tháng trước và cùng kì năm trước
                if (listDataYear.Count > 0 && listDataLastMonth.Count > 0 && listDataLastYear.Count > 0)
                {
                    // Số record của mảng
                    int countArray = 6;
                    arrayGradation = new GradationCharColumn[countArray * listDataCompareMonth.Count];
                    int count = 0;
                    int tooltip = 1;

                    foreach (ReportDetailtForTotalMoneyType item in listDataCompareMonth)
                    {
                        // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                        arrayGradation[count] = new GradationCharColumn()
                        {
                            Serie = "VND",
                            Segmento = string.Format("{0} {1}/{2}", item.MarketName, item.Month, item.Year),
                            Valor1 = item.VND,
                            Tooltip = tooltip
                        };

                        count++;

                        arrayGradation[count] = new GradationCharColumn()
                        {
                            Serie = "USD",
                            Segmento = string.Format("{0} {1}/{2}", item.MarketName, item.Month, item.Year),
                            Valor1 = item.USD,
                            Tooltip = tooltip
                        };

                        count++;

                        arrayGradation[count] = new GradationCharColumn()
                        {
                            Serie = "EUR",
                            Segmento = string.Format("{0} {1}/{2}", item.MarketName, item.Month, item.Year),
                            Valor1 = item.EUR,
                            Tooltip = tooltip
                        };

                        count++;

                        arrayGradation[count] = new GradationCharColumn()
                        {
                            Serie = "CAD",
                            Segmento = string.Format("{0} {1}/{2}", item.MarketName, item.Month, item.Year),
                            Valor1 = item.CAD,
                            Tooltip = tooltip
                        };

                        count++;

                        arrayGradation[count] = new GradationCharColumn()
                        {
                            Serie = "AUD",
                            Segmento = string.Format("{0} {1}/{2}", item.MarketName, item.Month, item.Year),
                            Valor1 = item.AUD,
                            Tooltip = tooltip
                        };

                        count++;

                        arrayGradation[count] = new GradationCharColumn()
                        {
                            Serie = "GBP",
                            Segmento = string.Format("{0} {1}/{2}", item.MarketName, item.Month, item.Year),
                            Valor1 = item.GBP,
                            Tooltip = tooltip
                        };

                        count++;
                        tooltip++;
                    }
                }
            }

            if (arrayGradation == null)
            {
                arrayGradation = new GradationCharColumn[1];
                arrayGradation[0] = new GradationCharColumn()
                {
                    Serie = "1",
                    Segmento = "1",
                    Valor1 = 0,
                    Tooltip = 1

                };
            }

            return Json(arrayGradation);
        }

        /// <summary>
        /// Get data cho việc vẽ biểu đồ cột chồng cho so sánh giai đoạn
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult SearchColumnChartStackCompareMonthForAll([DataSourceRequest]DataSourceRequest request, int month, int year, string reportTypeID)
        {
            List<ReportDetailtForTotalMoneyType> listDataCompareMonth = new ReportBL().ColumnChartStackDetailtMTCompareMonthForAllPercent(year, month, reportTypeID);

            GradationCharColumn[] arrayGradation = null;

            if (listDataCompareMonth.Count > 0)
            {
                List<ReportDetailtForTotalMoneyType> listDataYear = listDataCompareMonth.Where(x => x.Month == month.ToString() && x.Year == year.ToString()).ToList();
                List<ReportDetailtForTotalMoneyType> listDataLastMonth = listDataCompareMonth.Where(x => x.Month == (month - 1).ToString() && x.Year == year.ToString()).ToList();
                List<ReportDetailtForTotalMoneyType> listDataLastYear = listDataCompareMonth.Where(x => x.Month == month.ToString() && x.Year == (year - 1).ToString()).ToList();

                // Trường hợp đủ dữ liệu tháng hiện tại, tháng trước và cùng kì năm trước
                if (listDataYear.Count > 0 && listDataLastMonth.Count > 0 && listDataLastYear.Count > 0)
                {
                    // Số record của mảng
                    int countArray = 6;
                    arrayGradation = new GradationCharColumn[countArray * listDataCompareMonth.Count];
                    int count = 0;
                    int tooltip = 1;

                    foreach (ReportDetailtForTotalMoneyType item in listDataCompareMonth)
                    {
                        // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                        arrayGradation[count] = new GradationCharColumn()
                        {
                            Serie = "VND",
                            Segmento = string.Format("{0} {1}/{2}", item.MarketName, item.Month, item.Year),
                            Valor1 = item.VND,
                            Tooltip = tooltip
                        };

                        count++;

                        arrayGradation[count] = new GradationCharColumn()
                        {
                            Serie = "USD",
                            Segmento = string.Format("{0} {1}/{2}", item.MarketName, item.Month, item.Year),
                            Valor1 = item.USD,
                            Tooltip = tooltip
                        };

                        count++;

                        arrayGradation[count] = new GradationCharColumn()
                        {
                            Serie = "EUR",
                            Segmento = string.Format("{0} {1}/{2}", item.MarketName, item.Month, item.Year),
                            Valor1 = item.EUR,
                            Tooltip = tooltip
                        };

                        count++;

                        arrayGradation[count] = new GradationCharColumn()
                        {
                            Serie = "CAD",
                            Segmento = string.Format("{0} {1}/{2}", item.MarketName, item.Month, item.Year),
                            Valor1 = item.CAD,
                            Tooltip = tooltip
                        };

                        count++;

                        arrayGradation[count] = new GradationCharColumn()
                        {
                            Serie = "AUD",
                            Segmento = string.Format("{0} {1}/{2}", item.MarketName, item.Month, item.Year),
                            Valor1 = item.AUD,
                            Tooltip = tooltip
                        };

                        count++;

                        arrayGradation[count] = new GradationCharColumn()
                        {
                            Serie = "GBP",
                            Segmento = string.Format("{0} {1}/{2}", item.MarketName, item.Month, item.Year),
                            Valor1 = item.GBP,
                            Tooltip = tooltip
                        };

                        count++;
                        tooltip++;
                    }
                }
            }

            if (arrayGradation == null)
            {
                arrayGradation = new GradationCharColumn[1];
                arrayGradation[0] = new GradationCharColumn()
                {
                    Serie = "1",
                    Segmento = "1",
                    Valor1 = 0,
                    Tooltip = 1

                };
            }

            return Json(arrayGradation);
        }

        /// <summary>
        /// Get data cho việc vẽ biểu đồ cột cho so sánh giai đoạn
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult ColumnsChartCompareMonthForAll(string reportTypeID)
        {
            int year = DateTime.Now.Year;
            int toMonth = DateTime.Now.Month;

            List<ReportDetailtForTotalMoneyType> listDataCompareMonth = new ReportBL().ReportDetailtMTCompareMonthForAllConvert(year, toMonth, reportTypeID);

            GradationCompare[] arrayGradation = null;
            if (listDataCompareMonth.Count > 0)
            {
                List<ReportDetailtForTotalMoneyType> listDataYear = listDataCompareMonth.Where(x => x.Month == toMonth.ToString() && x.Year == year.ToString()).ToList();
                List<ReportDetailtForTotalMoneyType> listDataLastMonth = listDataCompareMonth.Where(x => x.Month == (toMonth - 1).ToString() && x.Year == year.ToString()).ToList();
                List<ReportDetailtForTotalMoneyType> listDataLastYear = listDataCompareMonth.Where(x => x.Month == toMonth.ToString() && x.Year == (year - 1).ToString()).ToList();

                // Trường hợp đủ dữ liệu tháng hiện tại, tháng trước và cùng kì năm trước
                if (listDataYear.Count > 0 && listDataLastMonth.Count > 0 && listDataLastYear.Count > 0)
                {
                    // Số record của mảng
                    int countArray = 6;
                    arrayGradation = new GradationCompare[countArray * listDataCompareMonth.Count];
                    int count = 0;

                    foreach (ReportDetailtForTotalMoneyType item in listDataCompareMonth)
                    {
                        // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                        arrayGradation[count] = new GradationCompare()
                        {
                            NameGradationCompare = "VND",
                            amount = item.VND,
                            NameType = string.Format("{0} {1}/{2}", item.MarketName, item.Month, item.Year)
                        };

                        count++;

                        arrayGradation[count] = new GradationCompare()
                        {
                            NameGradationCompare = "USD",
                            amount = item.USD,
                            NameType = string.Format("{0} {1}/{2}", item.MarketName, item.Month, item.Year)
                        };

                        count++;

                        arrayGradation[count] = new GradationCompare()
                        {
                            NameGradationCompare = "EUR",
                            amount = item.EUR,
                            NameType = string.Format("{0} {1}/{2}", item.MarketName, item.Month, item.Year)
                        };

                        count++;

                        arrayGradation[count] = new GradationCompare()
                        {
                            NameGradationCompare = "CAD",
                            amount = item.CAD,
                            NameType = string.Format("{0} {1}/{2}", item.MarketName, item.Month, item.Year)
                        };

                        count++;

                        arrayGradation[count] = new GradationCompare()
                        {
                            NameGradationCompare = "AUD",
                            amount = item.AUD,
                            NameType = string.Format("{0} {1}/{2}", item.MarketName, item.Month, item.Year)
                        };

                        count++;

                        arrayGradation[count] = new GradationCompare()
                        {
                            NameGradationCompare = "GBP",
                            amount = item.GBP,
                            NameType = string.Format("{0} {1}/{2}", item.MarketName, item.Month, item.Year)
                        };

                        count++;

                        year = year - 1;
                    }
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
        /// Get data cho việc vẽ biểu đồ cột cho so sánh giai đoạn
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult SearchColumnsChartCompareMonthForAll([DataSourceRequest]DataSourceRequest request, int month, int year, string reportTypeID)
        {
            List<ReportDetailtForTotalMoneyType> listDataCompareMonth = new ReportBL().ReportDetailtMTCompareMonthForAllConvert(year, month, reportTypeID);

            GradationCompare[] arrayGradation = null;
            if (listDataCompareMonth.Count > 0)
            {
                List<ReportDetailtForTotalMoneyType> listDataYear = listDataCompareMonth.Where(x => x.Month == month.ToString() && x.Year == year.ToString()).ToList();
                List<ReportDetailtForTotalMoneyType> listDataLastMonth = listDataCompareMonth.Where(x => x.Month == (month - 1).ToString() && x.Year == year.ToString()).ToList();
                List<ReportDetailtForTotalMoneyType> listDataLastYear = listDataCompareMonth.Where(x => x.Month == month.ToString() && x.Year == (year - 1).ToString()).ToList();

                // Trường hợp đủ dữ liệu tháng hiện tại, tháng trước và cùng kì năm trước
                if (listDataYear.Count > 0 && listDataLastMonth.Count > 0 && listDataLastYear.Count > 0)
                {
                    // Số record của mảng
                    int countArray = 6;
                    arrayGradation = new GradationCompare[countArray * listDataCompareMonth.Count];
                    int count = 0;

                    foreach (ReportDetailtForTotalMoneyType item in listDataCompareMonth)
                    {
                        // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                        arrayGradation[count] = new GradationCompare()
                        {
                            NameGradationCompare = "VND",
                            amount = item.VND,
                            NameType = string.Format("{0} {1}/{2}", item.MarketName, item.Month, item.Year)
                        };

                        count++;

                        arrayGradation[count] = new GradationCompare()
                        {
                            NameGradationCompare = "USD",
                            amount = item.USD,
                            NameType = string.Format("{0} {1}/{2}", item.MarketName, item.Month, item.Year)
                        };

                        count++;

                        arrayGradation[count] = new GradationCompare()
                        {
                            NameGradationCompare = "EUR",
                            amount = item.EUR,
                            NameType = string.Format("{0} {1}/{2}", item.MarketName, item.Month, item.Year)
                        };

                        count++;

                        arrayGradation[count] = new GradationCompare()
                        {
                            NameGradationCompare = "CAD",
                            amount = item.CAD,
                            NameType = string.Format("{0} {1}/{2}", item.MarketName, item.Month, item.Year)
                        };

                        count++;

                        arrayGradation[count] = new GradationCompare()
                        {
                            NameGradationCompare = "AUD",
                            amount = item.AUD,
                            NameType = string.Format("{0} {1}/{2}", item.MarketName, item.Month, item.Year)
                        };

                        count++;

                        arrayGradation[count] = new GradationCompare()
                        {
                            NameGradationCompare = "GBP",
                            amount = item.GBP,
                            NameType = string.Format("{0} {1}/{2}", item.MarketName, item.Month, item.Year)
                        };

                        count++;

                        year = year - 1;
                    }
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
        /// Get data cho báo cáo so sánh giai đoạn
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult ReportDetailtCompareMonthForOne([DataSourceRequest]DataSourceRequest request, string reportTypeID)
        {
            // Danh sach của data gradation gồm key và value
            int toYear = DateTime.Now.Year;
            int toMonth = DateTime.Now.Month;
            string marketID = "001";

            List<ReportDetailtForTotalMoneyType> listDataCompareMonth = new ReportBL().ReportDetailtMTCompareMonthForOneConvert(toYear, toMonth, reportTypeID, marketID);

            foreach (ReportDetailtForTotalMoneyType item in listDataCompareMonth)
            {
                item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;
            }

            // Danh sách mã thị trường của Tất cả
            List<string> listMarket = new List<string>() { "003", "005", "001", "002", "014", "004" };

            // Khởi tạo datatable
            DataTable table = new DataTable();
            // Tạo các cột cho datatable
            // Cùng kì năm trước
            table.Columns.Add("PartnerName", typeof(String));
            table.Columns.Add("VND1", typeof(double));
            table.Columns.Add("USD1", typeof(double));
            table.Columns.Add("EUR1", typeof(double));
            table.Columns.Add("CAD1", typeof(double));
            table.Columns.Add("AUD1", typeof(double));
            table.Columns.Add("GBP1", typeof(double));
            table.Columns.Add("TDS1", typeof(double));

            // tháng trước
            table.Columns.Add("VND2", typeof(double));
            table.Columns.Add("USD2", typeof(double));
            table.Columns.Add("EUR2", typeof(double));
            table.Columns.Add("CAD2", typeof(double));
            table.Columns.Add("AUD2", typeof(double));
            table.Columns.Add("GBP2", typeof(double));
            table.Columns.Add("TDS2", typeof(double));

            // Năm hiện tại
            table.Columns.Add("VND3", typeof(double));
            table.Columns.Add("USD3", typeof(double));
            table.Columns.Add("EUR3", typeof(double));
            table.Columns.Add("CAD3", typeof(double));
            table.Columns.Add("AUD3", typeof(double));
            table.Columns.Add("GBP3", typeof(double));
            table.Columns.Add("TDS3", typeof(double));

            table.Columns.Add("MarketName", typeof(String));

            List<string> listTemp = new List<string>();

            foreach (ReportDetailtForTotalMoneyType item in listDataCompareMonth)
            {
                // Cùng kì
                ReportDetailtForTotalMoneyType dataItemLastYear = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == toMonth.ToString() && x.Year == (toYear - 1).ToString());
                ReportDetailtForTotalMoneyType dataItemYear = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == toMonth.ToString() && x.Year == toYear.ToString());
                ReportDetailtForTotalMoneyType dataItemLastMonth = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == (toMonth - 1).ToString() && x.Year == toYear.ToString());

                if (!listTemp.Contains(item.PartnerCode))
                {
                    // Trường hợp năm trước có đối tác không có
                    if (dataItemLastYear == null)
                    {
                        dataItemLastYear = new ReportDetailtForTotalMoneyType();
                        dataItemLastYear.MarketCode = item.MarketCode;
                        dataItemLastYear.MarketName = item.MarketName;
                        dataItemLastYear.PartnerName = item.PartnerName;
                        dataItemLastYear.Year = item.Year;
                        dataItemLastYear.Month = item.Month;
                    }

                    // Trường hợp năm có đối tác không có
                    if (dataItemYear == null)
                    {
                        dataItemYear = new ReportDetailtForTotalMoneyType();
                        dataItemYear.MarketCode = item.MarketCode;
                        dataItemYear.MarketName = item.MarketName;
                        dataItemYear.PartnerName = item.PartnerName;
                        dataItemYear.Year = item.Year;
                        dataItemYear.Month = item.Month;
                    }

                    // Trường hợp năm có tháng trước có đối tác không có
                    if (dataItemLastMonth == null)
                    {
                        dataItemLastMonth = new ReportDetailtForTotalMoneyType();
                        dataItemLastMonth.MarketCode = item.MarketCode;
                        dataItemLastMonth.MarketName = item.MarketName;
                        dataItemLastMonth.PartnerName = item.PartnerName;
                        dataItemLastMonth.Year = item.Year;
                        dataItemLastMonth.Month = item.Month;
                    }

                    // Check tồn tại của item
                    string value = string.Format("PartnerName='{0}'", item.PartnerName);
                    DataRow[] foundRows = table.Select(value);

                    if (dataItemLastYear != null && dataItemYear != null && dataItemLastMonth != null && foundRows.Count() == 0)
                    {
                        // add item vào table
                        table.Rows.Add(item.PartnerName
                            , dataItemYear.VND, dataItemYear.USD, dataItemYear.EUR, dataItemYear.CAD, dataItemYear.AUD, dataItemYear.GBP, dataItemYear.TongDS
                            , dataItemLastMonth.VND, dataItemLastMonth.USD, dataItemLastMonth.EUR, dataItemLastMonth.CAD, dataItemLastMonth.AUD, dataItemLastMonth.GBP, dataItemLastMonth.TongDS
                            , dataItemLastYear.VND, dataItemLastYear.USD, dataItemLastYear.EUR, dataItemLastYear.CAD, dataItemLastYear.AUD, dataItemLastYear.GBP, dataItemLastYear.TongDS
                            , item.MarketName
                            );
                    }
                }
            }

            return Json(table.ToDataSourceResult(request), JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// Get data cho báo cáo so sánh giai đoạn
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult SearchReportDetailtCompareMonthForOne([DataSourceRequest]DataSourceRequest request, int month, int year, string reportTypeID, string marketID)
        {
            List<ReportDetailtForTotalMoneyType> listDataCompareMonth = new ReportBL().ReportDetailtMTCompareMonthForOneConvert(year, month, reportTypeID, marketID);

            foreach (ReportDetailtForTotalMoneyType item in listDataCompareMonth)
            {
                item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;
            }

            // Danh sách mã thị trường của Tất cả
            List<string> listMarket = new List<string>() { "003", "005", "001", "002", "014", "004" };

            // Khởi tạo datatable
            DataTable table = new DataTable();
            // Tạo các cột cho datatable
            // Cùng kì năm trước
            table.Columns.Add("PartnerName", typeof(String));
            table.Columns.Add("VND1", typeof(double));
            table.Columns.Add("USD1", typeof(double));
            table.Columns.Add("EUR1", typeof(double));
            table.Columns.Add("CAD1", typeof(double));
            table.Columns.Add("AUD1", typeof(double));
            table.Columns.Add("GBP1", typeof(double));
            table.Columns.Add("TDS1", typeof(double));

            // tháng trước
            table.Columns.Add("VND2", typeof(double));
            table.Columns.Add("USD2", typeof(double));
            table.Columns.Add("EUR2", typeof(double));
            table.Columns.Add("CAD2", typeof(double));
            table.Columns.Add("AUD2", typeof(double));
            table.Columns.Add("GBP2", typeof(double));
            table.Columns.Add("TDS2", typeof(double));

            // Năm hiện tại
            table.Columns.Add("VND3", typeof(double));
            table.Columns.Add("USD3", typeof(double));
            table.Columns.Add("EUR3", typeof(double));
            table.Columns.Add("CAD3", typeof(double));
            table.Columns.Add("AUD3", typeof(double));
            table.Columns.Add("GBP3", typeof(double));
            table.Columns.Add("TDS3", typeof(double));

            table.Columns.Add("MarketName", typeof(String));

            List<string> listTemp = new List<string>();

            foreach (ReportDetailtForTotalMoneyType item in listDataCompareMonth)
            {
                // Cùng kì
                ReportDetailtForTotalMoneyType dataItemLastYear = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == month.ToString() && x.Year == (year - 1).ToString());
                ReportDetailtForTotalMoneyType dataItemYear = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == month.ToString() && x.Year == year.ToString());
                ReportDetailtForTotalMoneyType dataItemLastMonth = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == (month - 1).ToString() && x.Year == year.ToString());

                if (!listTemp.Contains(item.PartnerCode))
                {
                    // Trường hợp năm trước có đối tác không có
                    if (dataItemLastYear == null)
                    {
                        dataItemLastYear = new ReportDetailtForTotalMoneyType();
                        dataItemLastYear.MarketCode = item.MarketCode;
                        dataItemLastYear.MarketName = item.MarketName;
                        dataItemLastYear.PartnerName = item.PartnerName;
                        dataItemLastYear.Year = item.Year;
                        dataItemLastYear.Month = item.Month;
                    }

                    // Trường hợp năm có đối tác không có
                    if (dataItemYear == null)
                    {
                        dataItemYear = new ReportDetailtForTotalMoneyType();
                        dataItemYear.MarketCode = item.MarketCode;
                        dataItemYear.MarketName = item.MarketName;
                        dataItemYear.PartnerName = item.PartnerName;
                        dataItemYear.Year = item.Year;
                        dataItemYear.Month = item.Month;
                    }

                    // Trường hợp năm có tháng trước có đối tác không có
                    if (dataItemLastMonth == null)
                    {
                        dataItemLastMonth = new ReportDetailtForTotalMoneyType();
                        dataItemLastMonth.MarketCode = item.MarketCode;
                        dataItemLastMonth.MarketName = item.MarketName;
                        dataItemLastMonth.PartnerName = item.PartnerName;
                        dataItemLastMonth.Year = item.Year;
                        dataItemLastMonth.Month = item.Month;
                    }

                    // Check tồn tại của item
                    string value = string.Format("PartnerName='{0}'", item.PartnerName);
                    DataRow[] foundRows = table.Select(value);

                    if (dataItemLastYear != null && dataItemYear != null && dataItemLastMonth != null && foundRows.Count() == 0)
                    {
                        // add item vào table
                        table.Rows.Add(item.PartnerName
                            , dataItemYear.VND, dataItemYear.USD, dataItemYear.EUR, dataItemYear.CAD, dataItemYear.AUD, dataItemYear.GBP, dataItemYear.TongDS
                            , dataItemLastMonth.VND, dataItemLastMonth.USD, dataItemLastMonth.EUR, dataItemLastMonth.CAD, dataItemLastMonth.AUD, dataItemLastMonth.GBP, dataItemLastMonth.TongDS
                            , dataItemLastYear.VND, dataItemLastYear.USD, dataItemLastYear.EUR, dataItemLastYear.CAD, dataItemLastYear.AUD, dataItemLastYear.GBP, dataItemLastYear.TongDS
                            , item.MarketName
                            );
                    }
                }
            }

            return Json(table.ToDataSourceResult(request), JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// Get data cho báo cáo so sánh giai đoạn
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult ReportDetailtCompareMonthForOneCompare([DataSourceRequest]DataSourceRequest request, string reportTypeID)
        {
            // Danh sach của data gradation gồm key và value
            int toYear = DateTime.Now.Year;
            int toMonth = DateTime.Now.Month;
            string marketID = "001";

            List<ReportDetailtForTotalMoneyType> listDataCompareMonth = new ReportBL().ReportDetailtMTCompareMonthForOneConvert(toYear, toMonth, reportTypeID, marketID);

            foreach (ReportDetailtForTotalMoneyType item in listDataCompareMonth)
            {
                item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;
            }

            // Khởi tạo datatable
            DataTable table = new DataTable();
            // Tạo các cột cho datatable
            // So sánh với tháng trước
            table.Columns.Add("PartnerName", typeof(String));
            table.Columns.Add("VND1", typeof(double));
            table.Columns.Add("USD1", typeof(double));
            table.Columns.Add("EUR1", typeof(double));
            table.Columns.Add("CAD1", typeof(double));
            table.Columns.Add("AUD1", typeof(double));
            table.Columns.Add("GBP1", typeof(double));
            table.Columns.Add("TDS1", typeof(double));

            // Cùng kì năm trước
            table.Columns.Add("VND2", typeof(double));
            table.Columns.Add("USD2", typeof(double));
            table.Columns.Add("EUR2", typeof(double));
            table.Columns.Add("CAD2", typeof(double));
            table.Columns.Add("AUD2", typeof(double));
            table.Columns.Add("GBP2", typeof(double));
            table.Columns.Add("TDS2", typeof(double));

            table.Columns.Add("MarketName", typeof(String));

            List<string> listTemp = new List<string>();

            foreach (ReportDetailtForTotalMoneyType item in listDataCompareMonth)
            {
                // Cùng kì
                ReportDetailtForTotalMoneyType dataItemLastYear = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == toMonth.ToString() && x.Year == (toYear - 1).ToString());
                ReportDetailtForTotalMoneyType dataItemYear = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == toMonth.ToString() && x.Year == toYear.ToString());
                ReportDetailtForTotalMoneyType dataItemLastMonth = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == (toMonth - 1).ToString() && x.Year == toYear.ToString());

                if (!listTemp.Contains(item.PartnerCode))
                {
                    // Trường hợp năm trước không có đối tác
                    if (dataItemLastYear == null)
                    {
                        dataItemLastYear = new ReportDetailtForTotalMoneyType();
                        dataItemLastYear.MarketCode = item.MarketCode;
                        dataItemLastYear.MarketName = item.MarketName;
                        dataItemLastYear.PartnerName = item.PartnerName;
                        dataItemLastYear.Year = item.Year;
                        dataItemLastYear.Month = item.Month;
                    }

                    // Trường hợp năm hiện tại không có đối tác
                    if (dataItemYear == null)
                    {
                        dataItemYear = new ReportDetailtForTotalMoneyType();
                        dataItemYear.MarketCode = item.MarketCode;
                        dataItemYear.MarketName = item.MarketName;
                        dataItemYear.PartnerName = item.PartnerName;
                        dataItemYear.Year = item.Year;
                        dataItemYear.Month = item.Month;
                    }

                    // Trường hợp tháng trước không có
                    if (dataItemLastMonth == null)
                    {
                        dataItemLastMonth = new ReportDetailtForTotalMoneyType();
                        dataItemLastMonth.MarketCode = item.MarketCode;
                        dataItemLastMonth.MarketName = item.MarketName;
                        dataItemLastMonth.PartnerName = item.PartnerName;
                        dataItemLastMonth.Year = item.Year;
                        dataItemLastMonth.Month = item.Month;
                    }

                    // so với tháng trước
                    double sumVND = dataItemYear.VND - dataItemLastMonth.VND;
                    double sumUSD = dataItemYear.USD - dataItemLastMonth.USD;
                    double sumEUR = dataItemYear.EUR - dataItemLastMonth.EUR;
                    double sumCAD = dataItemYear.CAD - dataItemLastMonth.CAD;
                    double sumAUD = dataItemYear.AUD - dataItemLastMonth.AUD;
                    double sumGBP = dataItemYear.GBP - dataItemLastMonth.GBP;

                    double sumTongDS = dataItemYear.TongDS - dataItemLastMonth.TongDS;

                    // so với cùng kì năm trước
                    double sumVNDLastYear = dataItemYear.VND - dataItemLastYear.VND;
                    double sumUSDLastYear = dataItemYear.USD - dataItemLastYear.USD;
                    double sumEURLastYear = dataItemYear.EUR - dataItemLastYear.EUR;
                    double sumCADLastYear = dataItemYear.CAD - dataItemLastYear.CAD;
                    double sumAUDLastYear = dataItemYear.AUD - dataItemLastYear.AUD;
                    double sumGBPLastYear = dataItemYear.GBP - dataItemLastYear.GBP;

                    double sumTongDSLastYear = dataItemYear.TongDS - dataItemLastYear.TongDS;

                    // Check tồn tại của item
                    string value = string.Format("PartnerName='{0}'", item.PartnerName);
                    DataRow[] foundRows = table.Select(value);

                    if (dataItemLastYear != null && dataItemYear != null && dataItemLastMonth != null && foundRows.Count() == 0)
                    {
                        // add item vào table
                        table.Rows.Add(item.PartnerName
                            // So với tháng hiện tại
                            , Math.Round(sumVND, 2, MidpointRounding.ToEven)
                            , Math.Round(sumUSD, 2, MidpointRounding.ToEven)
                            , Math.Round(sumEUR, 2, MidpointRounding.ToEven)
                            , Math.Round(sumCAD, 2, MidpointRounding.ToEven)
                            , Math.Round(sumAUD, 2, MidpointRounding.ToEven)
                            , Math.Round(sumGBP, 2, MidpointRounding.ToEven)
                            , Math.Round(sumTongDS, 2, MidpointRounding.ToEven)

                            // So với cùng kì năm trước
                            , Math.Round(sumVNDLastYear, 2, MidpointRounding.ToEven)
                            , Math.Round(sumUSDLastYear, 2, MidpointRounding.ToEven)
                            , Math.Round(sumEURLastYear, 2, MidpointRounding.ToEven)
                            , Math.Round(sumCADLastYear, 2, MidpointRounding.ToEven)
                            , Math.Round(sumAUDLastYear, 2, MidpointRounding.ToEven)
                            , Math.Round(sumGBPLastYear, 2, MidpointRounding.ToEven)
                            , Math.Round(sumTongDSLastYear, 2, MidpointRounding.ToEven)
                            , item.MarketName
                            );
                    }

                    // Add partnerCode để kiểm tra
                    listTemp.Add(item.PartnerCode);
                }
            }

            //DataRow row = table.NewRow();
            //row["PartnerName"] = "Tổng";
            //row["CQ1"] = table.Compute("Sum(CQ1)", "");
            //row["CN1"] = table.Compute("Sum(CN1)", "");
            //row["CK1"] = table.Compute("Sum(CK1)", "");
            //row["TDS1"] = table.Compute("Sum(TDS1)", "");

            //row["CQ2"] = table.Compute("Sum(CQ2)", "");
            //row["CN2"] = table.Compute("Sum(CN2)", "");
            //row["CK2"] = table.Compute("Sum(CK2)", "");
            //row["TDS2"] = table.Compute("Sum(TDS2)", "");
            //table.Rows.Add(row);

            //ViewBag.ListDataGradation = listDataGradation;
            return Json(table.ToDataSourceResult(request), JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// Get data cho báo cáo so sánh giai đoạn
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult SearchReportDetailtCompareMonthForOneCompare([DataSourceRequest]DataSourceRequest request, int month, int year, string reportTypeID, string marketID)
        {
            List<ReportDetailtForTotalMoneyType> listDataCompareMonth = new ReportBL().ReportDetailtMTCompareMonthForOneConvert(year, month, reportTypeID, marketID);

            foreach (ReportDetailtForTotalMoneyType item in listDataCompareMonth)
            {
                item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;
            }

            // Khởi tạo datatable
            DataTable table = new DataTable();
            // Tạo các cột cho datatable
            // So sánh với tháng trước
            table.Columns.Add("PartnerName", typeof(String));
            table.Columns.Add("VND1", typeof(double));
            table.Columns.Add("USD1", typeof(double));
            table.Columns.Add("EUR1", typeof(double));
            table.Columns.Add("CAD1", typeof(double));
            table.Columns.Add("AUD1", typeof(double));
            table.Columns.Add("GBP1", typeof(double));
            table.Columns.Add("TDS1", typeof(double));

            // Cùng kì năm trước
            table.Columns.Add("VND2", typeof(double));
            table.Columns.Add("USD2", typeof(double));
            table.Columns.Add("EUR2", typeof(double));
            table.Columns.Add("CAD2", typeof(double));
            table.Columns.Add("AUD2", typeof(double));
            table.Columns.Add("GBP2", typeof(double));
            table.Columns.Add("TDS2", typeof(double));

            table.Columns.Add("MarketName", typeof(String));

            List<string> listTemp = new List<string>();

            foreach (ReportDetailtForTotalMoneyType item in listDataCompareMonth)
            {
                // Cùng kì
                ReportDetailtForTotalMoneyType dataItemLastYear = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == month.ToString() && x.Year == (year - 1).ToString());
                ReportDetailtForTotalMoneyType dataItemYear = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == month.ToString() && x.Year == year.ToString());
                ReportDetailtForTotalMoneyType dataItemLastMonth = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == (month - 1).ToString() && x.Year == year.ToString());

                if (!listTemp.Contains(item.PartnerCode))
                {
                    // Trường hợp năm trước không có đối tác
                    if (dataItemLastYear == null)
                    {
                        dataItemLastYear = new ReportDetailtForTotalMoneyType();
                        dataItemLastYear.MarketCode = item.MarketCode;
                        dataItemLastYear.MarketName = item.MarketName;
                        dataItemLastYear.PartnerName = item.PartnerName;
                        dataItemLastYear.Year = item.Year;
                        dataItemLastYear.Month = item.Month;
                    }

                    // Trường hợp năm hiện tại không có đối tác
                    if (dataItemYear == null)
                    {
                        dataItemYear = new ReportDetailtForTotalMoneyType();
                        dataItemYear.MarketCode = item.MarketCode;
                        dataItemYear.MarketName = item.MarketName;
                        dataItemYear.PartnerName = item.PartnerName;
                        dataItemYear.Year = item.Year;
                        dataItemYear.Month = item.Month;
                    }

                    // Trường hợp tháng trước không có
                    if (dataItemLastMonth == null)
                    {
                        dataItemLastMonth = new ReportDetailtForTotalMoneyType();
                        dataItemLastMonth.MarketCode = item.MarketCode;
                        dataItemLastMonth.MarketName = item.MarketName;
                        dataItemLastMonth.PartnerName = item.PartnerName;
                        dataItemLastMonth.Year = item.Year;
                        dataItemLastMonth.Month = item.Month;
                    }

                    // so với tháng trước
                    double sumVND = dataItemYear.VND - dataItemLastMonth.VND;
                    double sumUSD = dataItemYear.USD - dataItemLastMonth.USD;
                    double sumEUR = dataItemYear.EUR - dataItemLastMonth.EUR;
                    double sumCAD = dataItemYear.CAD - dataItemLastMonth.CAD;
                    double sumAUD = dataItemYear.AUD - dataItemLastMonth.AUD;
                    double sumGBP = dataItemYear.GBP - dataItemLastMonth.GBP;

                    double sumTongDS = dataItemYear.TongDS - dataItemLastMonth.TongDS;

                    // so với cùng kì năm trước
                    double sumVNDLastYear = dataItemYear.VND - dataItemLastYear.VND;
                    double sumUSDLastYear = dataItemYear.USD - dataItemLastYear.USD;
                    double sumEURLastYear = dataItemYear.EUR - dataItemLastYear.EUR;
                    double sumCADLastYear = dataItemYear.CAD - dataItemLastYear.CAD;
                    double sumAUDLastYear = dataItemYear.AUD - dataItemLastYear.AUD;
                    double sumGBPLastYear = dataItemYear.GBP - dataItemLastYear.GBP;

                    double sumTongDSLastYear = dataItemYear.TongDS - dataItemLastYear.TongDS;

                    // Check tồn tại của item
                    string value = string.Format("PartnerName='{0}'", item.PartnerName);
                    DataRow[] foundRows = table.Select(value);

                    if (dataItemLastYear != null && dataItemYear != null && dataItemLastMonth != null && foundRows.Count() == 0)
                    {
                        // add item vào table
                        table.Rows.Add(item.PartnerName
                            // So với tháng hiện tại
                            , Math.Round(sumVND, 2, MidpointRounding.ToEven)
                            , Math.Round(sumUSD, 2, MidpointRounding.ToEven)
                            , Math.Round(sumEUR, 2, MidpointRounding.ToEven)
                            , Math.Round(sumCAD, 2, MidpointRounding.ToEven)
                            , Math.Round(sumAUD, 2, MidpointRounding.ToEven)
                            , Math.Round(sumGBP, 2, MidpointRounding.ToEven)
                            , Math.Round(sumTongDS, 2, MidpointRounding.ToEven)

                            // So với cùng kì năm trước
                            , Math.Round(sumVNDLastYear, 2, MidpointRounding.ToEven)
                            , Math.Round(sumUSDLastYear, 2, MidpointRounding.ToEven)
                            , Math.Round(sumEURLastYear, 2, MidpointRounding.ToEven)
                            , Math.Round(sumCADLastYear, 2, MidpointRounding.ToEven)
                            , Math.Round(sumAUDLastYear, 2, MidpointRounding.ToEven)
                            , Math.Round(sumGBPLastYear, 2, MidpointRounding.ToEven)
                            , Math.Round(sumTongDSLastYear, 2, MidpointRounding.ToEven)
                            , item.MarketName
                            );
                    }

                    // Add partnerCode để kiểm tra
                    listTemp.Add(item.PartnerCode);
                }
            }

            //DataRow row = table.NewRow();
            //row["PartnerName"] = "Tổng";
            //row["CQ1"] = table.Compute("Sum(CQ1)", "");
            //row["CN1"] = table.Compute("Sum(CN1)", "");
            //row["CK1"] = table.Compute("Sum(CK1)", "");
            //row["TDS1"] = table.Compute("Sum(TDS1)", "");

            //row["CQ2"] = table.Compute("Sum(CQ2)", "");
            //row["CN2"] = table.Compute("Sum(CN2)", "");
            //row["CK2"] = table.Compute("Sum(CK2)", "");
            //row["TDS2"] = table.Compute("Sum(TDS2)", "");
            //table.Rows.Add(row);

            //ViewBag.ListDataGradation = listDataGradation;
            return Json(table.ToDataSourceResult(request), JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// Get data cho việc vẽ biểu đồ cột chồng cho so sánh giai đoạn
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult ColumnChartCompareMonthStackForOne(string reportTypeID)
        {
            int toYear = DateTime.Now.Year;
            int toMonth = DateTime.Now.Month;
            string marketID = "001";

            List<ReportDetailtForTotalMoneyType> listDataCompareMonth = new ReportBL().ReportDetailtMTCompareMonthForOneConvertPercent(toYear, toMonth, reportTypeID, marketID);

            List<ReportDetailtForTotalMoneyType> listDataCommpareMonthClone = new List<ReportDetailtForTotalMoneyType>(listDataCompareMonth);

            List<string> listTemp = new List<string>();

            foreach (ReportDetailtForTotalMoneyType item in listDataCommpareMonthClone)
            {
                // Cùng kì
                ReportDetailtForTotalMoneyType dataItemLastYear = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == toMonth.ToString() && x.Year == (toYear - 1).ToString());
                ReportDetailtForTotalMoneyType dataItemYear = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == toMonth.ToString() && x.Year == toYear.ToString());
                ReportDetailtForTotalMoneyType dataItemLastMonth = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == (toMonth - 1).ToString() && x.Year == toYear.ToString());

                if (!listTemp.Contains(item.PartnerCode))
                {
                    // Trường hợp năm không có đối tác
                    if (dataItemLastYear == null)
                    {
                        dataItemLastYear = new ReportDetailtForTotalMoneyType();
                        dataItemLastYear.MarketCode = item.MarketCode;
                        dataItemLastYear.MarketName = item.MarketName;
                        dataItemLastYear.PartnerName = item.PartnerName;
                        dataItemLastYear.Year = (toYear - 1).ToString();
                        dataItemLastYear.Month = toMonth.ToString();
                        listDataCompareMonth.Add(dataItemLastYear);
                    }

                    // Trường hợp năm hiện tại không có đối tác
                    if (dataItemYear == null)
                    {
                        dataItemYear = new ReportDetailtForTotalMoneyType();
                        dataItemYear.MarketCode = item.MarketCode;
                        dataItemYear.MarketName = item.MarketName;
                        dataItemYear.PartnerName = item.PartnerName;
                        dataItemYear.Year = toYear.ToString();
                        dataItemYear.Month = toMonth.ToString();
                        listDataCompareMonth.Add(dataItemYear);
                    }

                    // Trường hợp tháng trước không có
                    if (dataItemLastMonth == null)
                    {
                        dataItemLastMonth = new ReportDetailtForTotalMoneyType();
                        dataItemLastMonth.MarketCode = item.MarketCode;
                        dataItemLastMonth.MarketName = item.MarketName;
                        dataItemLastMonth.PartnerName = item.PartnerName;
                        dataItemLastMonth.Year = toYear.ToString();
                        dataItemLastMonth.Month = (toMonth - 1).ToString();
                        listDataCompareMonth.Add(dataItemLastMonth);
                    }

                    // Add partnerCode để kiểm tra
                    listTemp.Add(item.PartnerCode);
                }
            }


            // Số mảng cần tạo
            int arrayCount = 6;
            GradationCharColumn[] arrayGradation = new GradationCharColumn[listDataCompareMonth.Count * arrayCount];
            int count = 0;

            foreach (ReportDetailtForTotalMoneyType item in listDataCompareMonth)
            {
                // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                arrayGradation[count] = new GradationCharColumn()
                {
                    Serie = "VND",
                    Segmento = string.Format("{0} {1}/{2}", item.PartnerName, item.Month, item.Year),
                    Valor1 = item.VND,
                    Month = item.Month,
                    Year = item.Year,
                    //Tooltip = tooltip
                };
                count++;

                arrayGradation[count] = new GradationCharColumn()
                {
                    Serie = "USD",
                    Segmento = string.Format("{0} {1}/{2}", item.PartnerName, item.Month, item.Year),
                    Valor1 = item.USD,
                    Month = item.Month,
                    Year = item.Year,
                    //Tooltip = tooltip
                };
                count++;

                arrayGradation[count] = new GradationCharColumn()
                {
                    Serie = "EUR",
                    Segmento = string.Format("{0} {1}/{2}", item.PartnerName, item.Month, item.Year),
                    Valor1 = item.EUR,
                    Month = item.Month,
                    Year = item.Year,
                    //Tooltip = tooltip
                };
                count++;

                arrayGradation[count] = new GradationCharColumn()
                {
                    Serie = "CAD",
                    Segmento = string.Format("{0} {1}/{2}", item.PartnerName, item.Month, item.Year),
                    Valor1 = item.CAD,
                    Month = item.Month,
                    Year = item.Year,
                    //Tooltip = tooltip
                };
                count++;

                arrayGradation[count] = new GradationCharColumn()
                {
                    Serie = "AUD",
                    Segmento = string.Format("{0} {1}/{2}", item.PartnerName, item.Month, item.Year),
                    Valor1 = item.AUD,
                    Month = item.Month,
                    Year = item.Year,
                    //Tooltip = tooltip
                };
                count++;

                arrayGradation[count] = new GradationCharColumn()
                {
                    Serie = "GBP",
                    Segmento = string.Format("{0} {1}/{2}", item.PartnerName, item.Month, item.Year),
                    Valor1 = item.GBP,
                    Month = item.Month,
                    Year = item.Year,
                    //Tooltip = tooltip
                };
                count++;
            }

            if (listDataCompareMonth.Count == 0)
            {
                arrayGradation = new GradationCharColumn[1];
                arrayGradation[0] = new GradationCharColumn()
                {
                    Serie = "1",
                    Segmento = "1",
                    Valor1 = 0,
                    Tooltip = 1

                };
            }

            return Json(arrayGradation);
        }

        /// <summary>
        /// Get data cho việc vẽ biểu đồ cột chồng cho so sánh giai đoạn
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult SearchColumnChartCompareMonthStackForOne([DataSourceRequest]DataSourceRequest request, int month, int year, string reportTypeID, string marketID)
        {
            List<ReportDetailtForTotalMoneyType> listDataCompareMonth = new ReportBL().ReportDetailtMTCompareMonthForOneConvertPercent(year, month, reportTypeID, marketID);

            List<ReportDetailtForTotalMoneyType> listDataCommpareMonthClone = new List<ReportDetailtForTotalMoneyType>(listDataCompareMonth);

            List<string> listTemp = new List<string>();

            foreach (ReportDetailtForTotalMoneyType item in listDataCommpareMonthClone)
            {
                // Cùng kì
                ReportDetailtForTotalMoneyType dataItemLastYear = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == month.ToString() && x.Year == (year - 1).ToString());
                ReportDetailtForTotalMoneyType dataItemYear = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == month.ToString() && x.Year == year.ToString());
                ReportDetailtForTotalMoneyType dataItemLastMonth = listDataCompareMonth.Find(x => x.PartnerCode == item.PartnerCode && x.Month == (month - 1).ToString() && x.Year == year.ToString());

                if (!listTemp.Contains(item.PartnerCode))
                {
                    // Trường hợp năm không có đối tác
                    if (dataItemLastYear == null)
                    {
                        dataItemLastYear = new ReportDetailtForTotalMoneyType();
                        dataItemLastYear.MarketCode = item.MarketCode;
                        dataItemLastYear.MarketName = item.MarketName;
                        dataItemLastYear.PartnerName = item.PartnerName;
                        dataItemLastYear.Year = (year - 1).ToString();
                        dataItemLastYear.Month = month.ToString();
                        listDataCompareMonth.Add(dataItemLastYear);
                    }

                    // Trường hợp năm hiện tại không có đối tác
                    if (dataItemYear == null)
                    {
                        dataItemYear = new ReportDetailtForTotalMoneyType();
                        dataItemYear.MarketCode = item.MarketCode;
                        dataItemYear.MarketName = item.MarketName;
                        dataItemYear.PartnerName = item.PartnerName;
                        dataItemYear.Year = year.ToString();
                        dataItemYear.Month = month.ToString();
                        listDataCompareMonth.Add(dataItemYear);
                    }

                    // Trường hợp tháng trước không có
                    if (dataItemLastMonth == null)
                    {
                        dataItemLastMonth = new ReportDetailtForTotalMoneyType();
                        dataItemLastMonth.MarketCode = item.MarketCode;
                        dataItemLastMonth.MarketName = item.MarketName;
                        dataItemLastMonth.PartnerName = item.PartnerName;
                        dataItemLastMonth.Year = year.ToString();
                        dataItemLastMonth.Month = (month - 1).ToString();
                        listDataCompareMonth.Add(dataItemLastMonth);
                    }

                    // Add partnerCode để kiểm tra
                    listTemp.Add(item.PartnerCode);
                }
            }


            // Số mảng cần tạo
            int arrayCount = 6;
            GradationCharColumn[] arrayGradation = new GradationCharColumn[listDataCompareMonth.Count * arrayCount];
            int count = 0;

            foreach (ReportDetailtForTotalMoneyType item in listDataCompareMonth)
            {
                // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                arrayGradation[count] = new GradationCharColumn()
                {
                    Serie = "VND",
                    Segmento = string.Format("{0} {1}/{2}", item.PartnerName, item.Month, item.Year),
                    Valor1 = item.VND,
                    //Tooltip = tooltip
                };
                count++;

                arrayGradation[count] = new GradationCharColumn()
                {
                    Serie = "USD",
                    Segmento = string.Format("{0} {1}/{2}", item.PartnerName, item.Month, item.Year),
                    Valor1 = item.USD,
                    //Tooltip = tooltip
                };
                count++;

                arrayGradation[count] = new GradationCharColumn()
                {
                    Serie = "EUR",
                    Segmento = string.Format("{0} {1}/{2}", item.PartnerName, item.Month, item.Year),
                    Valor1 = item.EUR,
                    //Tooltip = tooltip
                };
                count++;

                arrayGradation[count] = new GradationCharColumn()
                {
                    Serie = "CAD",
                    Segmento = string.Format("{0} {1}/{2}", item.PartnerName, item.Month, item.Year),
                    Valor1 = item.CAD,
                    //Tooltip = tooltip
                };
                count++;

                arrayGradation[count] = new GradationCharColumn()
                {
                    Serie = "AUD",
                    Segmento = string.Format("{0} {1}/{2}", item.PartnerName, item.Month, item.Year),
                    Valor1 = item.AUD,
                    //Tooltip = tooltip
                };
                count++;

                arrayGradation[count] = new GradationCharColumn()
                {
                    Serie = "GBP",
                    Segmento = string.Format("{0} {1}/{2}", item.PartnerName, item.Month, item.Year),
                    Valor1 = item.GBP,
                    //Tooltip = tooltip
                };
                count++;
            }

            if (listDataCompareMonth.Count == 0)
            {
                arrayGradation = new GradationCharColumn[1];
                arrayGradation[0] = new GradationCharColumn()
                {
                    Serie = "1",
                    Segmento = "1",
                    Valor1 = 0,
                    Tooltip = 1

                };
            }

            return Json(arrayGradation);
        }

        /// <summary>
        /// Get data cho việc vẽ biểu đồ cột cho so sánh giai đoạn
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult ColumnsChartCompareMonthForOne(string reportTypeID)
        {
            int year = DateTime.Now.Year;
            int toMonth = DateTime.Now.Month;
            string marketID = "001";

            List<ReportDetailtForTotalMoneyType> listDataCompareMonth = new ReportBL().ReportDetailtMTCompareMonthForOneConvert(year, toMonth, reportTypeID, marketID);

            // Số record của mảng
            int countArray = 6;
            GradationCompare[] arrayGradation = new GradationCompare[countArray * listDataCompareMonth.Count];
            int count = 0;
            foreach (ReportDetailtForTotalMoneyType item in listDataCompareMonth)
            {
                // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                arrayGradation[count] = new GradationCompare()
                {
                    NameGradationCompare = "VND",
                    amount = item.VND,
                    NameType = string.Format("{0} {1}/{2}", item.PartnerName, item.Month, item.Year)
                };

                count++;
                arrayGradation[count] = new GradationCompare()
                {
                    NameGradationCompare = "USD",
                    amount = item.USD,
                    NameType = string.Format("{0} {1}/{2}", item.PartnerName, item.Month, item.Year)
                };

                count++;

                arrayGradation[count] = new GradationCompare()
                {
                    NameGradationCompare = "EUR",
                    amount = item.EUR,
                    NameType = string.Format("{0} {1}/{2}", item.PartnerName, item.Month, item.Year)
                };

                count++;

                arrayGradation[count] = new GradationCompare()
                {
                    NameGradationCompare = "CAD",
                    amount = item.CAD,
                    NameType = string.Format("{0} {1}/{2}", item.PartnerName, item.Month, item.Year)
                };

                count++;

                arrayGradation[count] = new GradationCompare()
                {
                    NameGradationCompare = "AUD",
                    amount = item.AUD,
                    NameType = string.Format("{0} {1}/{2}", item.PartnerName, item.Month, item.Year)
                };

                count++;

                arrayGradation[count] = new GradationCompare()
                {
                    NameGradationCompare = "GBP",
                    amount = item.GBP,
                    NameType = string.Format("{0} {1}/{2}", item.PartnerName, item.Month, item.Year)
                };

                count++;

                year = year - 1;
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
        /// Get data cho việc vẽ biểu đồ cột cho so sánh giai đoạn
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpPost]
        public ActionResult SearchColumnsChartCompareMonthForOne([DataSourceRequest]DataSourceRequest request, int month, int year, string reportTypeID, string marketID)
        {
            List<ReportDetailtForTotalMoneyType> listDataCompareMonth = new ReportBL().ReportDetailtMTCompareMonthForOneConvert(year, month, reportTypeID, marketID);

            // Số record của mảng
            int countArray = 6;
            GradationCompare[] arrayGradation = new GradationCompare[countArray * listDataCompareMonth.Count];
            int count = 0;
            foreach (ReportDetailtForTotalMoneyType item in listDataCompareMonth)
            {
                // Tạo mảng insert dữ liệu để vẽ biểu đồ cột
                arrayGradation[count] = new GradationCompare()
                {
                    NameGradationCompare = "VND",
                    amount = item.VND,
                    NameType = string.Format("{0} {1}/{2}", item.PartnerName, item.Month, item.Year)
                };

                count++;
                arrayGradation[count] = new GradationCompare()
                {
                    NameGradationCompare = "USD",
                    amount = item.USD,
                    NameType = string.Format("{0} {1}/{2}", item.PartnerName, item.Month, item.Year)
                };

                count++;

                arrayGradation[count] = new GradationCompare()
                {
                    NameGradationCompare = "EUR",
                    amount = item.EUR,
                    NameType = string.Format("{0} {1}/{2}", item.PartnerName, item.Month, item.Year)
                };

                count++;

                arrayGradation[count] = new GradationCompare()
                {
                    NameGradationCompare = "CAD",
                    amount = item.CAD,
                    NameType = string.Format("{0} {1}/{2}", item.PartnerName, item.Month, item.Year)
                };

                count++;

                arrayGradation[count] = new GradationCompare()
                {
                    NameGradationCompare = "AUD",
                    amount = item.AUD,
                    NameType = string.Format("{0} {1}/{2}", item.PartnerName, item.Month, item.Year)
                };

                count++;

                arrayGradation[count] = new GradationCompare()
                {
                    NameGradationCompare = "GBP",
                    amount = item.GBP,
                    NameType = string.Format("{0} {1}/{2}", item.PartnerName, item.Month, item.Year)
                };

                count++;

                year = year - 1;
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

    }
}
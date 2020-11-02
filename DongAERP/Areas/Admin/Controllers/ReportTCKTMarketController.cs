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
    public class ReportTCKTMarketController : Controller
    {
        // GET: Admin/ReportTCKTMarket
        public ActionResult Index()
        {
            return View();
        }
        
        /// <summary>
        /// Màn hình báo cáo cho ngày
        /// </summary>
        /// <returns></returns>
        public ActionResult ReportDay(DateTime? fromDate, DateTime? toDate, string reportTypeID, string marketID)
        {
            string nameUrl = "Doanh số/Thị trường/Tổng hợp/";
            ViewBag.NameURL = nameUrl;

            if (fromDate != null && toDate != null && reportTypeID != null && marketID != null)
            {
                List<string> listData = new List<string>()
                {
                    fromDate.Value.ToString("MM/dd/yyyy"),
                    toDate.Value.ToString("MM/dd/yyyy"),
                    reportTypeID,
                    marketID
                };

                ViewData["listData"] = listData;
            }

            return View();
        }

        /// <summary>
        /// Màn hình báo cáo cho ngày
        /// </summary>
        /// <returns></returns>
        public ActionResult ReportMonth(DateTime? fromDate, DateTime? toDate, string reportTypeID, string marketID)
        {
            string nameUrl = "Doanh số/Thị trường/Tổng hợp/";
            ViewBag.NameURL = nameUrl;

            if (fromDate != null && toDate != null && reportTypeID != null && marketID != null)
            {
                List<string> listData = new List<string>()
                {
                    fromDate.Value.ToString("MM/dd/yyyy"),
                    toDate.Value.ToString("MM/dd/yyyy"),
                    reportTypeID,
                    marketID
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
        public ActionResult SearchReportDay([DataSourceRequest]DataSourceRequest request, DateTime fromDate, DateTime toDate, string reportTypeID, string marketID)
        {
            List<ReportDetailtSTMarket> listData = new ReportBL().SearchMarketTCKTForDay(fromDate, toDate, reportTypeID);

            // Khởi tạo datatable
            DataTable table = new DataTable();

            if (!string.IsNullOrEmpty(marketID))
            {
                List<string> ListMarket = new List<string>();
                if (marketID == "0")
                {
                    foreach (ReportDetailtSTMarket item in listData)
                    {
                        // CHâu Á
                        if (item.ParentCode == "005")
                        {
                            item.MarketName = "Châu Á";
                        }
                        if (!ListMarket.Contains(item.MarketName))
                        {
                            ListMarket.Add(item.MarketName);
                        }
                        item.TongDS = item.DSChiNha + item.DSChiQuay + item.DSCK;
                    }

                    table.Columns.Add("MarketName", typeof(String));
                    table.Columns.Add("PartnerName", typeof(String));
                    // Add column
                    for (DateTime x = fromDate; x <= toDate; x = x.AddDays(1))
                    {
                        string nameDate = string.Format("COL{0}", x.Day.ToString("00"));

                        table.Columns.Add(nameDate, typeof(double));

                    }

                    table.Columns.Add("TongDS", typeof(double));
                    
                    // List thị trường
                    foreach(string item in ListMarket)
                    {
                        // List Partner thuộc 1 thị trường
                        List<ReportDetailtSTMarket> listDataItem = listData.Where(x => x.MarketName == item).ToList();

                        List<string> listPartner = new List<string>();

                        foreach(ReportDetailtSTMarket item1 in listDataItem)
                        {
                            if (listPartner.Contains(item1.PartnerName))
                            {
                                continue;
                            }
                            DataRow rows = table.NewRow();
                            rows["MarketName"] = item1.MarketName;
                            rows["PartnerName"] = item1.PartnerName;

                            double tongDS = 0;
                            // Từ ngày đến ngày
                            for (DateTime x = fromDate; x <= toDate; x = x.AddDays(1))
                            {
                                string nameDate = string.Format("COL{0}", x.Day.ToString("00"));

                                ReportDetailtSTMarket dataItem = listDataItem.Find(o => o.CreatedDate == x && o.PartnerName == item1.PartnerName);
                                // Trường hợp không có dữ liệu
                                if(dataItem == null)
                                {
                                    dataItem = new ReportDetailtSTMarket()
                                    {
                                        PartnerName = item1.PartnerName,
                                        MarketName = item1.MarketName
                                    };
                                }
                                tongDS = tongDS + dataItem.TongDS;
                               rows[nameDate] = dataItem.TongDS;
                            }

                            // add dòng tổng
                            rows["TongDS"] = tongDS;
                            // Add row vào table
                            table.Rows.Add(rows);
                            // Add giá trị của partner
                            listPartner.Add(item1.PartnerName);
                        }
                    }
                }
                else
                {
                    List<ReportDetailtSTMarket> listDataConvert = listData.Where(x => x.ParentCode == "005").ToList();

                    foreach (ReportDetailtSTMarket item in listDataConvert)
                    {
                        if (!ListMarket.Contains(item.MarketName))
                        {
                            ListMarket.Add(item.MarketName);
                        }
                        item.TongDS = item.DSChiNha + item.DSChiQuay + item.DSCK;
                    }

                    table.Columns.Add("MarketName", typeof(String));
                    table.Columns.Add("PartnerName", typeof(String));
                    // Add column
                    for (DateTime x = fromDate; x <= toDate; x = x.AddDays(1))
                    {
                        string nameDate = string.Format("COL{0}", x.Day.ToString("00"));

                        table.Columns.Add(nameDate, typeof(double));

                    }

                    table.Columns.Add("TongDS", typeof(double));

                    // List thị trường
                    foreach (string item in ListMarket)
                    {
                        // List Partner thuộc 1 thị trường
                        List<ReportDetailtSTMarket> listDataItem = listDataConvert.Where(x => x.MarketName == item).ToList();

                        List<string> listPartner = new List<string>();

                        foreach (ReportDetailtSTMarket item1 in listDataItem)
                        {
                            if (listPartner.Contains(item1.PartnerName))
                            {
                                continue;
                            }
                            DataRow rows = table.NewRow();
                            rows["MarketName"] = item1.MarketName;
                            rows["PartnerName"] = item1.PartnerName;

                            double tongDS = 0;
                            // Từ ngày đến ngày
                            for (DateTime x = fromDate; x <= toDate; x = x.AddDays(1))
                            {
                                string nameDate = string.Format("COL{0}", x.Day.ToString("00"));

                                ReportDetailtSTMarket dataItem = listDataItem.Find(o => o.CreatedDate == x && o.PartnerName == item1.PartnerName);
                                // Trường hợp không có dữ liệu
                                if (dataItem == null)
                                {
                                    dataItem = new ReportDetailtSTMarket()
                                    {
                                        PartnerName = item1.PartnerName,
                                        MarketName = item1.MarketName
                                    };
                                }
                                tongDS = tongDS + dataItem.TongDS;
                                rows[nameDate] = dataItem.TongDS;
                            }

                            // add dòng tổng
                            rows["TongDS"] = tongDS;
                            // Add row vào table
                            table.Rows.Add(rows);
                            // Add giá trị của partner
                            listPartner.Add(item1.PartnerName);
                        }
                    }
                }
            }

            return Json(table.ToDataSourceResult(request), JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// Search report day theo ngày nhập vào
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public ActionResult SearchReportMonth([DataSourceRequest]DataSourceRequest request, DateTime fromDate, DateTime toDate, string reportTypeID, string marketID)
        {
            List<ReportDetailtSTMarket> listData = new ReportBL().SearchMarketTCKTForMonth(fromDate, toDate, reportTypeID);

            // Khởi tạo datatable
            DataTable table = new DataTable();

            if (!string.IsNullOrEmpty(marketID))
            {
                List<string> ListMarket = new List<string>();
                if (marketID == "0")
                {
                    foreach (ReportDetailtSTMarket item in listData)
                    {
                        // CHâu Á
                        if (item.ParentCode == "005")
                        {
                            item.MarketName = "Châu Á";
                        }
                        if (!ListMarket.Contains(item.MarketName))
                        {
                            ListMarket.Add(item.MarketName);
                        }
                        item.TongDS = item.DSChiNha + item.DSChiQuay + item.DSCK;
                    }

                    table.Columns.Add("MarketName", typeof(String));
                    table.Columns.Add("PartnerName", typeof(String));

                    for(int i = fromDate.Month; i < toDate.Month; i++)
                    {
                        string nameDate = string.Format("COL{0}", i.ToString("00"));

                        table.Columns.Add(nameDate, typeof(double));
                    }
                    

                    table.Columns.Add("TongDS", typeof(double));

                    // List thị trường
                    foreach (string item in ListMarket)
                    {
                        // List Partner thuộc 1 thị trường
                        List<ReportDetailtSTMarket> listDataItem = listData.Where(x => x.MarketName == item).ToList();

                        List<string> listPartner = new List<string>();

                        foreach (ReportDetailtSTMarket item1 in listDataItem)
                        {
                            if (listPartner.Contains(item1.PartnerName))
                            {
                                continue;
                            }
                            DataRow rows = table.NewRow();
                            rows["MarketName"] = item1.MarketName;
                            rows["PartnerName"] = item1.PartnerName;

                            double tongDS = 0;
                            // Từ ngày đến ngày
                            for (int i = fromDate.Month; i < toDate.Month; i++)
                            {
                                string nameDate = string.Format("COL{0}", i.ToString("00"));

                                ReportDetailtSTMarket dataItem = listDataItem.Find(o => o.Month == i.ToString() && o.PartnerName == item1.PartnerName);
                                // Trường hợp không có dữ liệu
                                if (dataItem == null)
                                {
                                    dataItem = new ReportDetailtSTMarket()
                                    {
                                        PartnerName = item1.PartnerName,
                                        MarketName = item1.MarketName
                                    };
                                }
                                tongDS = tongDS + dataItem.TongDS;
                                rows[nameDate] = dataItem.TongDS;
                            }

                            // add dòng tổng
                            rows["TongDS"] = tongDS;
                            // Add row vào table
                            table.Rows.Add(rows);
                            // Add giá trị của partner
                            listPartner.Add(item1.PartnerName);
                        }
                    }
                }
                else
                {
                    List<ReportDetailtSTMarket> listDataConvert = listData.Where(x => x.ParentCode == "005").ToList();

                    foreach (ReportDetailtSTMarket item in listDataConvert)
                    {
                        if (!ListMarket.Contains(item.MarketName))
                        {
                            ListMarket.Add(item.MarketName);
                        }
                        item.TongDS = item.DSChiNha + item.DSChiQuay + item.DSCK;
                    }

                    table.Columns.Add("MarketName", typeof(String));
                    table.Columns.Add("PartnerName", typeof(String));
                    // Add column
                    for (int i = fromDate.Month; i < toDate.Month; i++)
                    {
                        string nameDate = string.Format("COL{0}", i.ToString("00"));

                        table.Columns.Add(nameDate, typeof(double));
                    }

                    table.Columns.Add("TongDS", typeof(double));

                    // List thị trường
                    foreach (string item in ListMarket)
                    {
                        // List Partner thuộc 1 thị trường
                        List<ReportDetailtSTMarket> listDataItem = listDataConvert.Where(x => x.MarketName == item).ToList();

                        List<string> listPartner = new List<string>();

                        foreach (ReportDetailtSTMarket item1 in listDataItem)
                        {
                            if (listPartner.Contains(item1.PartnerName))
                            {
                                continue;
                            }
                            DataRow rows = table.NewRow();
                            rows["MarketName"] = item1.MarketName;
                            rows["PartnerName"] = item1.PartnerName;

                            double tongDS = 0;
                            // Từ ngày đến ngày
                            for (int i = fromDate.Month; i < toDate.Month; i++)
                            {
                                string nameDate = string.Format("COL{0}", i.ToString("00"));

                                ReportDetailtSTMarket dataItem = listDataItem.Find(o => o.Month == i.ToString() && o.PartnerName == item1.PartnerName);
                                // Trường hợp không có dữ liệu
                                if (dataItem == null)
                                {
                                    dataItem = new ReportDetailtSTMarket()
                                    {
                                        PartnerName = item1.PartnerName,
                                        MarketName = item1.MarketName
                                    };
                                }
                                tongDS = tongDS + dataItem.TongDS;
                                rows[nameDate] = dataItem.TongDS;
                            }

                            // add dòng tổng
                            rows["TongDS"] = tongDS;
                            // Add row vào table
                            table.Rows.Add(rows);
                            // Add giá trị của partner
                            listPartner.Add(item1.PartnerName);
                        }
                    }
                }
            }

            return Json(table.ToDataSourceResult(request), JsonRequestBehavior.AllowGet);
        }
    }
}
// #################################################################
// # Copyright (C) 2010-2011, ASoft JSC.  All Rights Reserved.
// #
// # History：
// #	Date Time		Updated			Content
// #	10/06/2020		Truong Lam		Create New
// ##################################################################

using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DongA.Core;
using Models.Framework;
using DongA.DataAccess;
using static DongA.Core.DongAException;
using Models;
using DongA.Entities;

namespace DongA.Bussiness
{
    public class ReportBL: DongABaseDAL
    {
        private object dal;

        /// <summary>
        /// List Report
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtServiceType> InsertTableMarket()
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<ReportDetailtServiceType> result = dal.InsertTableMarket();
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }

        }

        /// <summary>
        /// List Report
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<Entities.Account> GetListAccount()
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<Entities.Account> result = dal.GetListAccount();
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }

        }

        /// <summary>
        /// List Report
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<Report> GetListReport(string reportType)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<Report> result = dal.GetListReport(reportType);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }

        }

        /// <summary>
        /// List Report
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<Report> SearchDay(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<Report> result = dal.SearchDay(fromDate, toDate, reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }

        }

        /// <summary>
        /// List Report
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<Report> GetListReportMonth(DateTime now, string reportTypeID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<Report> result = dal.GetListReportMonth(now, reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// List Report
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<Report> SearchMonth(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<Report> result = dal.SearchMonth(fromDate, toDate, reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }

        }

        /// <summary>
        /// List Report by year
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<Report> GetListReportYear(string reportTypeID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<Report> result = dal.GetListReportYear(reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// List Report
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<Report> SearchYear(DateTime fromYear, DateTime toYear, string reportTypeID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<Report> result = dal.SearchYear(fromYear, toYear, reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }

        }

        /// <summary>
        /// List Report cho so sánh giai đoạn
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<Report> ListDataGradationCompare(int ToYear, int typeID, string reportTypeID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<Report> result = dal.ListDataGradationCompare(ToYear, typeID, reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// List data hiển thị phần trăm cho so sánh giai đoạn
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<Report> ListDataGradationComparePercent(int typeID, int ToYear, string reportTypeID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<Report> result = dal.ListDataGradationCompare(ToYear, typeID, reportTypeID);
                List<Report> DataConvertPercent = new List<Report>();
                // Xử lý việc chuyển đổi thành %
                foreach (Report item in result)
                {
                    item.TongDS = item.DSChiQuay + item.DSChiNha + item.DSCK;
                    Report itemConvert = new Report()
                    {
                        DSChiQuay = Math.Round(item.DSChiQuay / item.TongDS * 100, 2, MidpointRounding.ToEven),
                        DSChiNha = Math.Round(item.DSChiNha / item.TongDS * 100, 2, MidpointRounding.ToEven),
                        DSCK = Math.Round(item.DSCK / item.TongDS * 100, 2, MidpointRounding.ToEven),
                        TongDS = Math.Round(item.TongDS / item.TongDS * 100, 2, MidpointRounding.ToEven),
                        Year = item.Year
                    };
                    DataConvertPercent.Add(itemConvert);
                }

                return DataConvertPercent;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }

        }

        /// <summary>
        /// List data hiển thị phần trăm cho so sánh giai đoạn
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<Report> ListDataGradationComparePie(int typeID, int ToYear, string reportTypeID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<Report> DataConvertPercent = new List<Report>();
                List<Report> result = dal.ListDataGradationCompare(ToYear, typeID, reportTypeID);
                if (result.Count > 0)
                {
                    Report resultYear = result.Find(x => x.Year == ToYear.ToString());
                    if (resultYear != null)
                    {
                        resultYear.TongDS = resultYear.DSChiQuay + resultYear.DSChiNha + resultYear.DSCK;
                        Report itemConvert = new Report()
                        {
                            DSChiQuay = Math.Round(resultYear.DSChiQuay / resultYear.TongDS * 100, 2, MidpointRounding.ToEven),
                            DSChiNha = Math.Round(resultYear.DSChiNha / resultYear.TongDS * 100, 2, MidpointRounding.ToEven),
                            DSCK = Math.Round(resultYear.DSCK / resultYear.TongDS * 100, 2, MidpointRounding.ToEven),
                            TongDS = Math.Round(resultYear.TongDS / resultYear.TongDS * 100, 2, MidpointRounding.ToEven),
                        };
                        DataConvertPercent.Add(itemConvert);
                    }
                }
                return DataConvertPercent;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// List Report cho so sánh last month
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<Report> ListDataMonthCompareGrid(int ToYear, int ToMonth, string reportTypeID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<Report> result = dal.ListDataMonthCompareGrid(ToYear, ToMonth, reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }

        }

        /// <summary>
        /// List Report cho so sánh last month
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<Report> ListDataLastMonthCompareProportion(int toYear, int toMonth, string reportTypeID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<Report> result = dal.ListDataMonthCompareGrid(toYear, toMonth, reportTypeID);
                List<Report> listDataPercent = new List<Report>();

                if (result.Count.Equals(3))
                {
                    foreach (Report item in result)
                    {
                        item.TongDS = item.DSChiQuay + item.DSChiNha + item.DSCK;
                        Report dataItem = new Report()
                        {
                            Year = item.Year,
                            Month =  item.Month,
                            CreatedDate = item.CreatedDate,
                            DSChiQuay = item.TongDS == 0 ? 0 : Math.Round(item.DSChiQuay / item.TongDS * 100, 2, MidpointRounding.ToEven),
                            DSChiNha = item.TongDS == 0 ? 0 : Math.Round(item.DSChiNha / item.TongDS * 100, 2, MidpointRounding.ToEven),
                            DSCK = item.TongDS == 0 ? 0 : Math.Round(item.DSCK / item.TongDS * 100, 2, MidpointRounding.ToEven)
                        };
                        listDataPercent.Add(dataItem);
                    }
                }
                return listDataPercent;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// List Report cho so sánh last month
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<Report> ListDataLastMonthCompareProportionPercent(int toYear, int toMonth, string reportTypeID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<Report> DataConvertPercent = new List<Report>();
                List<Report> result = dal.ListDataMonthCompareGrid(toYear, toMonth, reportTypeID);
                // Tháng hiện tại, tháng trước
                if (result.Count.Equals(3))
                {
                    foreach (Report item in result)
                    {
                        item.TongDS = item.DSChiQuay + item.DSChiNha + item.DSCK;
                        Report reportPercent = new Report()
                        {
                            ReportID = item.ReportID,
                            Day = item.Day,
                            Month = item.Month,
                            Year = item.Year,
                            GradationID = item.GradationID,
                            DSChiQuay = Math.Round(item.DSChiQuay / item.TongDS * 100, 2, MidpointRounding.ToEven),
                            DSChiNha = Math.Round(item.DSChiNha / item.TongDS * 100, 2, MidpointRounding.ToEven),
                            DSCK = Math.Round(item.DSCK / item.TongDS * 100, 2, MidpointRounding.ToEven)
                        };
                        DataConvertPercent.Add(reportPercent);
                    }
                }
                return DataConvertPercent;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// List Report
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportForMaket> DataReportMaketForDay(string reportTypeID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<ReportForMaket> result = dal.DataReportMaketForDay(reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }

        }

        /// <summary>
        /// Danh sách report theo thị trường
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportForMaket> SearchReportMaketForDay(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<ReportForMaket> result = dal.SearchReportMaketForDay(fromDate, toDate, reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }

        }

        /// <summary>
        /// List Report
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportForMaket> DataReportMaketForMonth(string reportTypeID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<ReportForMaket> result = dal.DataReportMaketForMonth(reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// Danh sách report tháng theo thị trường
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportForMaket> SearchReportMaketForMonth(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();

                List<ReportForMaket> result = dal.SearchReportMaketForMonth(fromDate, toDate, reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// List Report
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportForMaket> DataReportMaketForYear(string reportTypeID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<ReportForMaket> result = dal.DataReportMaketForYear(reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// Danh sách report tháng theo thị trường
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportForMaket> SearchReportMaketForYear(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<ReportForMaket> result = dal.SearchReportMaketForYear(fromDate, toDate, reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// List Report
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportForMaket> DataReportMaketForGradationCompare(int year, int typeID, string reportTypeID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<ReportForMaket> result = dal.DataReportMaketForGradationCompare(year, typeID, reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// List Report
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportForMaket> DataReportMaketForGradationComparePercent(int year, int typeID, string reportTypeID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<ReportForMaket> listDataPercent = new List<ReportForMaket>();
                List<ReportForMaket> result = dal.DataReportMaketForGradationCompareCompare(year, typeID, reportTypeID);

                if (result.Count.Equals(2))
                {
                    foreach (ReportForMaket item in result)
                    {
                        item.TongDS = item.American + item.Asia + item.Global + item.Europe + item.Canada + item.Australia;
                        ReportForMaket dataItem = new ReportForMaket()
                        {
                            Year = item.Year,
                            American = item.TongDS == 0 ? 0 : Math.Round(item.American / item.TongDS * 100, 2, MidpointRounding.ToEven),
                            Asia = item.TongDS == 0 ? 0 : Math.Round(item.Asia / item.TongDS * 100, 2, MidpointRounding.ToEven),
                            Global = item.TongDS == 0 ? 0 : Math.Round(item.Global / item.TongDS * 100, 2, MidpointRounding.ToEven),
                            Europe = item.TongDS == 0 ? 0 : Math.Round(item.Europe / item.TongDS * 100, 2, MidpointRounding.ToEven),
                            Canada = item.TongDS == 0 ? 0 : Math.Round(item.Canada / item.TongDS * 100, 2, MidpointRounding.ToEven),
                            Australia = item.TongDS == 0 ? 0 : Math.Round(item.Australia / item.TongDS * 100, 2, MidpointRounding.ToEven)
                        };
                        listDataPercent.Add(dataItem);
                    }
                }
                return listDataPercent;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// List Report cho so sánh last month
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportForMaket> DataReportCompareForMonth(int ToYear, int ToMonth, string reportTypeID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<ReportForMaket> result = dal.DataReportCompareForMonth(ToYear, ToMonth, reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// List Report cho so sánh last month
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportForMaket> DataReportMarketCompareForMonthPercent(int ToYear, int ToMonth, string reportTypeID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<ReportForMaket> result = dal.DataReportCompareForMonth(ToYear, ToMonth, reportTypeID);
                List<ReportForMaket> resultPercent = new List<ReportForMaket>();

                if (result.Count.Equals(3))
                {
                    foreach (ReportForMaket item in result)
                    {
                        item.TongDS = item.American + item.Asia + item.Global + item.Europe + item.Canada + item.Australia;

                        ReportForMaket dataItem = new ReportForMaket()
                        {
                            Month = item.Month,
                            Year = item.Year,
                            American = item.TongDS == 0 ? 0 : Math.Round(item.American / item.TongDS * 100, 2, MidpointRounding.ToEven),
                            Asia = item.TongDS == 0 ? 0 : Math.Round(item.Asia / item.TongDS * 100, 2, MidpointRounding.ToEven),
                            Global = item.TongDS == 0 ? 0 : Math.Round(item.Global / item.TongDS * 100, 2, MidpointRounding.ToEven),
                            Europe = item.TongDS == 0 ? 0 : Math.Round(item.Europe / item.TongDS * 100, 2, MidpointRounding.ToEven),
                            Canada = item.TongDS == 0 ? 0 : Math.Round(item.Canada / item.TongDS * 100, 2, MidpointRounding.ToEven),
                            Australia = item.TongDS == 0 ? 0 : Math.Round(item.Australia / item.TongDS * 100, 2, MidpointRounding.ToEven)
                        };
                        resultPercent.Add(dataItem);
                    }
                }
                return resultPercent;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// Get data for report total payment
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportForTotalPayment> DataReportTPForDay(string reportTypeID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<ReportForTotalPayment> result = dal.DataReportTPForDay(reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }

        }

        /// <summary>
        /// Danh sách report theo Tổng doanh số chi trả
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportForTotalPayment> SearchReportTPForDay(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<ReportForTotalPayment> result = dal.SearchReportTPForDay(fromDate, toDate, reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }

        }

        /// <summary>
        /// List Report for month
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportForTotalPayment> DataReportTPForMonth(string reportTypeID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<ReportForTotalPayment> result = dal.DataReportTPForMonth(reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// Danh sách report tháng theo thị trường
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportForTotalPayment> SearchReportTPForMonth(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<ReportForTotalPayment> result = dal.SearchReportTPForMonth(fromDate, toDate, reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// List Report
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportForTotalPayment> DataReportTPForYear(string reportTypeID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<ReportForTotalPayment> result = dal.DataReportTPForYear(reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// Danh sách report tháng theo thị trường
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportForTotalPayment> SearchReportTPForYear(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<ReportForTotalPayment> result = dal.SearchReportTPForYear(fromDate, toDate, reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// List Report
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportForTotalPayment> DataReportTPForGradationCompare(int toYear, int typeID, string reportTypeID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<ReportForTotalPayment> result = dal.DataReportTPForGradationCompare(toYear, typeID, reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// List Report
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportForTotalPayment> DataReportTPCompareForMonth(int toYear, int toMonth, string reportTypeID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<ReportForTotalPayment> result = dal.DataReportTPCompareForMonth(toYear, toMonth, reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        ///// <summary>
        ///// List Report cho so sánh last month
        ///// </summary>
        ///// <returns></returns>
        ///// <history>
        /////     [Truong Lam]   Created [10/06/2020]
        ///// </history>
        //public List<ReportForTotalPayment> SearchReportTotalPaymentCompareForMonth(int typeID, int ToYear, int ToMonth)
        //{
        //    try
        //    {
        //        ReportDAL dal = new ReportDAL();
        //        List<ReportForTotalPayment> result = dal.SearchReportTotalPaymentCompareForMonth(typeID, ToYear, ToMonth);
        //        return result;
        //    }
        //    catch (Exception ex)
        //    {
        //        throw new DongAException(DongALayer.Business, ex.Message, ex);
        //    }
        //}

        /// <summary>
        /// List Report
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportForTotalMoneyType> DataReportTMTForDay(DateTime now, string reportTypeID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<ReportForTotalMoneyType> result = dal.DataReportTotalMoneyTypeForDay(now, reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// List Report
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportForTotalMoneyType> DataReportTMTForDayConvert(DateTime now, string reportTypeID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<ReportForTotalMoneyType> result = dal.DataReportTMTForDayConvert(now, reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// Danh sách report theo thị trường
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportForTotalMoneyType> SearchReportTMTForDay(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<ReportForTotalMoneyType> result = dal.SearchReportTMTForDay(fromDate, toDate, reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// Danh sách report theo thị trường
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportForTotalMoneyType> SearchReportTMTForDayConvert(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<ReportForTotalMoneyType> result = dal.SearchReportTMTForDayConvert(fromDate, toDate, reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// List Report
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportForTotalMoneyType> DataReportTMTForMonth(string reportTypeID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<ReportForTotalMoneyType> result = dal.DataReportTMTForMonth(reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// List Report
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportForTotalMoneyType> DataReportTMTForMonthConvert(string reportTypeID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<ReportForTotalMoneyType> result = dal.DataReportTMTForMonthConvert(reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// Danh sách report tháng - Nguyên tệ
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportForTotalMoneyType> SearchReportTMTForMonth(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<ReportForTotalMoneyType> result = dal.SearchReportTMTForMonth(fromDate, toDate, reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// Danh sách report tháng - Nguyên tệ
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportForTotalMoneyType> SearchReportTMTForMonthConvert(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<ReportForTotalMoneyType> result = dal.SearchReportTMTForMonthConvert(fromDate, toDate, reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// List Report theo năm - Nguyên tệ
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportForTotalMoneyType> DataReportTMTForYear(string reportTypeID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<ReportForTotalMoneyType> result = dal.DataReportTMTForYear(reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// List Report theo năm - USD
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportForTotalMoneyType> DataReportTMTForYearConvert(string reportTypeID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<ReportForTotalMoneyType> result = dal.DataReportTMTForYearConvert(reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// Danh sách report năm - Nguyên tệ
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportForTotalMoneyType> SearchReportTMTForYear(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<ReportForTotalMoneyType> result = dal.SearchReportTMTForYear(fromDate, toDate, reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// Danh sách report năm - USD
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportForTotalMoneyType> SearchReportTMTForYearConvert(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<ReportForTotalMoneyType> result = dal.SearchReportTMTForYearConvert(fromDate, toDate, reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// List Report theo giai đoạn
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportForTotalMoneyType> DataReportTMTForGradationCompare(int year, int typeID, string reportTypeID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<ReportForTotalMoneyType> result = dal.DataReportTMTForGradationCompare(year, typeID, reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// List Report theo giai đoạn
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportForTotalMoneyType> DataReportTMTForGradationCompareConvert(int year, int typeID, string reportTypeID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<ReportForTotalMoneyType> result = dal.DataReportTMTForGradationCompareConvert(year, typeID, reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// List Report theo giai đoạn - USD (%)
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportForTotalMoneyType> DataReportTMTForGradationComparePercent(int year, int typeID, string reportTypeID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<ReportForTotalMoneyType> result = dal.DataReportTMTForGradationCompareConvert(year, typeID, reportTypeID);
                List<ReportForTotalMoneyType> listDataPercent = new List<ReportForTotalMoneyType>();
                
                if (result.Count.Equals(2))
                {
                    foreach(ReportForTotalMoneyType item in result)
                    {
                        item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;
                        ReportForTotalMoneyType dataItem = new ReportForTotalMoneyType()
                        {
                            Year = item.Year,
                            VND = item.TongDS == 0 ? 0 : Math.Round(item.VND / item.TongDS * 100, 2, MidpointRounding.ToEven),
                            USD = item.TongDS == 0 ? 0 : Math.Round(item.USD / item.TongDS * 100, 2, MidpointRounding.ToEven),
                            EUR = item.TongDS == 0 ? 0 : Math.Round(item.EUR / item.TongDS * 100, 2, MidpointRounding.ToEven),
                            CAD = item.TongDS == 0 ? 0 : Math.Round(item.CAD / item.TongDS * 100, 2, MidpointRounding.ToEven),
                            AUD = item.TongDS == 0 ? 0 : Math.Round(item.AUD / item.TongDS * 100, 2, MidpointRounding.ToEven),
                            GBP = item.TongDS == 0 ? 0 : Math.Round(item.GBP / item.TongDS * 100, 2, MidpointRounding.ToEven)
                        };
                        listDataPercent.Add(dataItem);
                    }
                }
                return listDataPercent;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// List Report cho so sánh last month
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportForTotalMoneyType> DataReportTMTCompareForMonth(int toYear, int toMonth, string reportTypeID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<ReportForTotalMoneyType> result = dal.DataReportTMTCompareForMonth(toYear, toMonth, reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// List Report cho so sánh last month
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportForTotalMoneyType> DataReportTMTCompareForMonthConvert(int toYear, int toMonth, string reportTypeID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<ReportForTotalMoneyType> result = dal.DataReportTMTCompareForMonthConvert(toYear, toMonth, reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// List Report cho so sánh last month
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportForTotalMoneyType> DataReportTMTCompareForMonthPercent(int toYear, int toMonth, string reportTypeID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<ReportForTotalMoneyType> result = dal.DataReportTMTCompareForMonthConvert(toYear, toMonth, reportTypeID);
                List<ReportForTotalMoneyType> resultPercent = new List<ReportForTotalMoneyType>();
                if (result.Count.Equals(3))
                {
                    foreach(ReportForTotalMoneyType item in result)
                    {
                        item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;

                        ReportForTotalMoneyType dataItem = new ReportForTotalMoneyType()
                        {
                            Month = item.Month,
                            Year = item.Year,
                            VND = Math.Round(item.VND / item.TongDS * 100, 2, MidpointRounding.ToEven),
                            USD = Math.Round(item.USD / item.TongDS * 100, 2, MidpointRounding.ToEven),
                            EUR = Math.Round(item.EUR / item.TongDS * 100, 2, MidpointRounding.ToEven),
                            CAD = Math.Round(item.CAD / item.TongDS * 100, 2, MidpointRounding.ToEven),
                            AUD = Math.Round(item.AUD / item.TongDS * 100, 2, MidpointRounding.ToEven),
                            GBP = Math.Round(item.GBP / item.TongDS * 100, 2, MidpointRounding.ToEven)
                        };
                        resultPercent.Add(dataItem);
                    }
                }
                return resultPercent;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// List Report cho so sánh last month
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportForMaket> CreateDataMarket()
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<ReportForMaket> result = dal.CreateDataMarket();
                
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// List Report detailt theo ngày
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/08/2020]
        /// </history>
        public List<ReportDetailtServiceType> GetListReportDetailtForDay(string reportType)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<ReportDetailtServiceType> result = dal.GetListReportDetailtForDay(reportType);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// List Report
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtSTMarket> SearchReportDetailtForDay(DateTime fromDate, DateTime toDate, string reportTypeID, string marketID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<ReportDetailtSTMarket> result = dal.SearchReportDetailtForDay(fromDate, toDate, reportTypeID, marketID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// List Report
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtSTMarket> SearchMarketForTotalForMonth(DateTime fromDate, DateTime toDate, string reportTypeID, string marketID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<ReportDetailtSTMarket> result = dal.SearchMarketForTotalForMonth(fromDate, toDate, reportTypeID, marketID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// List Report detailt cho năm
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtSTMarket> SearchMarketForTotalForYear(DateTime fromDate, DateTime toDate, string reportTypeID, string marketID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<ReportDetailtSTMarket> result = dal.SearchMarketForTotalForYear(fromDate, toDate, reportTypeID, marketID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// List Report
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtServiceType> MarketForPartner(string reportTypeID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<ReportDetailtServiceType> result = dal.MarketForPartner(reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// List Report
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<Partner> ListPartner()
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<Partner> result = dal.ListPartner();
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// Danh sách các thị trường
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<Market> ListMarket()
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<Market> result = dal.ListMarket();
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// List Report detailt theo ngày
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/08/2020]
        /// </history>
        public List<ReportDetailtServiceType> GetListReportDetailtForOneMarket(string reportType)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<ReportDetailtServiceType> result = dal.GetListReportDetailtForOneMarket(reportType);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// List Report
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtServiceType> SearchMarketForOne(DateTime fromDate, DateTime toDate, string reportTypeID, string marketID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<ReportDetailtServiceType> result = dal.SearchMarketForOne(fromDate, toDate, reportTypeID, marketID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// List Report
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtServiceType> SearchMarketForOneForMonth(DateTime fromDate, DateTime toDate, string reportTypeID, string marketID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<ReportDetailtServiceType> result = dal.SearchMarketForOneForMonth(fromDate, toDate, reportTypeID, marketID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// List Report detailt cho báo cáo theo năm
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtServiceType> SearchMarketForOneForYear(DateTime fromDate, DateTime toDate, string reportTypeID, string marketID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<ReportDetailtServiceType> result = dal.SearchMarketForOneForYear(fromDate, toDate, reportTypeID, marketID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// List Report cho so sánh giai đoạn theo báo cáo chi tiết
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtServiceType> ReportDetailtGradationCompareForAll(int ToYear, int typeID, string reportTypeID, string marketID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<ReportDetailtServiceType> result = dal.ReportDetailtGradationCompareForAll(ToYear, typeID, reportTypeID, marketID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// List Report cho so sánh giai đoạn theo báo cáo chi tiết
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtServiceType> ReportDetailtGradationCompareForAllPercent(int ToYear, int typeID, string reportTypeID, string marketID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<ReportDetailtServiceType> result = dal.ReportDetailtGradationCompareForAll(ToYear, typeID, reportTypeID, marketID);
                List<ReportDetailtServiceType> resultConvert = new List<ReportDetailtServiceType>();
                double sumDSChiQuayYear = 0;
                double sumDSChiNhaYear = 0;
                double sumDSCKYear = 0;

                double sumDSChiQuayLastYear = 0;
                double sumDSChiNhaLastYear = 0;
                double sumDSCKLastYear = 0;

                if (result.Count > 0)
                {
                    // Last Year
                    sumDSChiQuayLastYear = result.Where(x=> x.Year == (ToYear - 1).ToString()).Sum(y=>y.DSChiQuay);
                    sumDSChiNhaLastYear = result.Where(x => x.Year == (ToYear - 1).ToString()).Sum(y => y.DSChiNha);
                    sumDSCKLastYear = result.Where(x => x.Year == (ToYear - 1).ToString()).Sum(y => y.DSCK);

                    // Year hiện tại
                    sumDSChiQuayYear = result.Where(x => x.Year == ToYear.ToString()).Sum(y => y.DSChiQuay);
                    sumDSChiNhaYear = result.Where(x => x.Year == ToYear.ToString()).Sum(y => y.DSChiNha);
                    sumDSCKYear = result.Where(x => x.Year == ToYear.ToString()).Sum(y => y.DSCK);
                }

                foreach (ReportDetailtServiceType item in result)
                {
                    item.TongDS = item.DSChiQuay + item.DSChiNha + item.DSCK;
                    ReportDetailtServiceType itemDetailtPercent = new ReportDetailtServiceType();

                    // Last year
                    if (item.Year == (ToYear - 1).ToString())
                    {
                        itemDetailtPercent = new ReportDetailtServiceType()
                        {
                            MarketCode = item.MarketCode,
                            MarketName = item.MarketName,
                            DSChiQuay = item.TongDS == 0 ? 0 : Math.Round(item.DSChiQuay / sumDSChiQuayLastYear * 100, 2, MidpointRounding.ToEven),
                            DSChiNha = item.TongDS == 0 ? 0 : Math.Round(item.DSChiNha / sumDSChiNhaLastYear * 100, 2, MidpointRounding.ToEven),
                            DSCK = item.TongDS == 0 ? 0 : Math.Round(item.DSCK / sumDSCKLastYear * 100, 2, MidpointRounding.ToEven),
                            Year = item.Year
                        };
                    }

                    // year hien tai
                    if (item.Year == ToYear.ToString())
                    {
                        itemDetailtPercent = new ReportDetailtServiceType()
                        {
                            MarketCode = item.MarketCode,
                            MarketName = item.MarketName,
                            DSChiQuay = item.TongDS == 0 ? 0 : Math.Round(item.DSChiQuay / sumDSChiQuayYear * 100, 2, MidpointRounding.ToEven),
                            DSChiNha = item.TongDS == 0 ? 0 : Math.Round(item.DSChiNha / sumDSChiNhaYear * 100, 2, MidpointRounding.ToEven),
                            DSCK = item.TongDS == 0 ? 0 : Math.Round(item.DSCK / sumDSCKYear * 100, 2, MidpointRounding.ToEven),
                            Year = item.Year
                        };
                    }
                    resultConvert.Add(itemDetailtPercent);
                }
                return resultConvert;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// List Report cho so sánh giai đoạn theo báo cáo chi tiết
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtServiceType> ReportDetailtGradationCompareForOne(int ToYear, int typeID, string reportTypeID, string marketID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<ReportDetailtServiceType> result = dal.ReportDetailtGradationCompareForOne(ToYear, typeID, reportTypeID, marketID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// List Report cho so sánh giai đoạn theo báo cáo chi tiết
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtServiceType> ReportDetailtGradationCompareForOnePercent(int ToYear, int typeID, string reportTypeID, string marketID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<ReportDetailtServiceType> result = dal.ReportDetailtGradationCompareForOne(ToYear, typeID, reportTypeID, marketID);
                List<ReportDetailtServiceType> resultConvert = new List<ReportDetailtServiceType>();
                double sumDSChiQuayYear = 0;
                double sumDSChiNhaYear = 0;
                double sumDSCKYear = 0;

                double sumDSChiQuayLastYear = 0;
                double sumDSChiNhaLastYear = 0;
                double sumDSCKLastYear = 0;

                if (result.Count > 0)
                {
                    // Last Year
                    sumDSChiQuayLastYear = result.Where(x => x.Year == (ToYear - 1).ToString()).Sum(y => y.DSChiQuay);
                    sumDSChiNhaLastYear = result.Where(x => x.Year == (ToYear - 1).ToString()).Sum(y => y.DSChiNha);
                    sumDSCKLastYear = result.Where(x => x.Year == (ToYear - 1).ToString()).Sum(y => y.DSCK);

                    // Year hiện tại
                    sumDSChiQuayYear = result.Where(x => x.Year == ToYear.ToString()).Sum(y => y.DSChiQuay);
                    sumDSChiNhaYear = result.Where(x => x.Year == ToYear.ToString()).Sum(y => y.DSChiNha);
                    sumDSCKYear = result.Where(x => x.Year == ToYear.ToString()).Sum(y => y.DSCK);
                }

                foreach (ReportDetailtServiceType item in result)
                {
                    item.TongDS = item.DSChiQuay + item.DSChiNha + item.DSCK;
                    ReportDetailtServiceType itemDetailtPercent = new ReportDetailtServiceType();

                    // Last year
                    if (item.Year == (ToYear - 1).ToString())
                    {
                        itemDetailtPercent = new ReportDetailtServiceType()
                        {
                            MarketCode = item.MarketCode,
                            MarketName = item.MarketName,
                            PartnerCode = item.PartnerCode,
                            PartnerName = item.PartnerName,
                            DSChiQuay = item.TongDS == 0 ? 0 : Math.Round(item.DSChiQuay / sumDSChiQuayLastYear * 100, 2, MidpointRounding.ToEven),
                            DSChiNha = item.TongDS == 0 ? 0 : Math.Round(item.DSChiNha / sumDSChiNhaLastYear * 100, 2, MidpointRounding.ToEven),
                            DSCK = item.TongDS == 0 ? 0 : Math.Round(item.DSCK / sumDSCKLastYear * 100, 2, MidpointRounding.ToEven),
                            Year = item.Year
                        };
                    }

                    // year hien tai
                    if (item.Year == ToYear.ToString())
                    {
                        itemDetailtPercent = new ReportDetailtServiceType()
                        {
                            MarketCode = item.MarketCode,
                            MarketName = item.MarketName,
                            DSChiQuay = item.TongDS == 0 ? 0 : Math.Round(item.DSChiQuay / sumDSChiQuayYear * 100, 2, MidpointRounding.ToEven),
                            DSChiNha = item.TongDS == 0 ? 0 : Math.Round(item.DSChiNha / sumDSChiNhaYear * 100, 2, MidpointRounding.ToEven),
                            DSCK = item.TongDS == 0 ? 0 : Math.Round(item.DSCK / sumDSCKYear * 100, 2, MidpointRounding.ToEven),
                            Year = item.Year
                        };
                    }
                    resultConvert.Add(itemDetailtPercent);
                }
                return resultConvert;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// List Report chi tiết cho báo cáo theo tháng
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtSTMarket> ReportDetailtCompareMonthForAll(int toYear, int toMonth, string reportTypeID, string marketID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<ReportDetailtSTMarket> result = dal.ReportDetailtCompareMonthForAll(toYear, toMonth, reportTypeID, marketID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// List Report cho so sánh giai đoạn theo báo cáo chi tiết
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtSTMarket> ColumnChartStackCompareMonthForAllPercent(int toYear, int toMonth, string reportTypeID, string marketID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<ReportDetailtSTMarket> result = dal.ReportDetailtCompareMonthForAll(toYear, toMonth, reportTypeID, marketID);
                List<ReportDetailtSTMarket> resultConvert = new List<ReportDetailtSTMarket>();
                // Tháng hiện tại
                double sumDSChiQuayYear = 0;
                double sumDSChiNhaYear = 0;
                double sumDSCKYear = 0;

                // Tháng trước
                double sumDSChiQuayLastMonth = 0;
                double sumDSChiNhaLastMonth = 0;
                double sumDSCKLastMonth = 0;

                // Cùng kì năm trước
                double sumDSChiQuayLastYear = 0;
                double sumDSChiNhaLastYear = 0;
                double sumDSCKLastYear = 0;

                if (result.Count > 0)
                {
                    // Cùng kì năm trước
                    sumDSChiQuayLastYear = result.Where(x => x.Year == (toYear - 1).ToString() && x.Month == toMonth.ToString()).Sum(y => y.DSChiQuay);
                    sumDSChiNhaLastYear = result.Where(x => x.Year == (toYear - 1).ToString() && x.Month == toMonth.ToString()).Sum(y => y.DSChiNha);
                    sumDSCKLastYear = result.Where(x => x.Year == (toYear - 1).ToString() && x.Month == toMonth.ToString()).Sum(y => y.DSCK);

                    // Tháng hiện tại
                    sumDSChiQuayYear = result.Where(x => x.Year == toYear.ToString() && x.Month == toMonth.ToString()).Sum(y => y.DSChiQuay);
                    sumDSChiNhaYear = result.Where(x => x.Year == toYear.ToString() && x.Month == toMonth.ToString()).Sum(y => y.DSChiNha);
                    sumDSCKYear = result.Where(x => x.Year == toYear.ToString() && x.Month == toMonth.ToString()).Sum(y => y.DSCK);

                    // Tháng trước
                    sumDSChiQuayLastMonth = result.Where(x => x.Year == toYear.ToString() && x.Month == (toMonth - 1).ToString()).Sum(y => y.DSChiQuay);
                    sumDSChiNhaLastMonth = result.Where(x => x.Year == toYear.ToString() && x.Month == (toMonth - 1).ToString()).Sum(y => y.DSChiNha);
                    sumDSCKLastMonth = result.Where(x => x.Year == toYear.ToString() && x.Month == (toMonth - 1).ToString()).Sum(y => y.DSCK);
                }

                foreach (ReportDetailtSTMarket item in result)
                {
                    item.TongDS = item.DSChiQuay + item.DSChiNha + item.DSCK;
                    ReportDetailtSTMarket itemDetailtPercent = new ReportDetailtSTMarket();

                    // Cùng kì năm trước
                    if (item.Year == (toYear - 1).ToString() && item.Month == toMonth.ToString())
                    {
                        itemDetailtPercent = new ReportDetailtSTMarket()
                        {
                            MarketCode = item.MarketCode,
                            MarketName = item.MarketName,
                            DSChiQuay = item.TongDS == 0 ? 0 : Math.Round(item.DSChiQuay / sumDSChiQuayLastYear * 100, 2, MidpointRounding.ToEven),
                            DSChiNha = item.TongDS == 0 ? 0 : Math.Round(item.DSChiNha / sumDSChiNhaLastYear * 100, 2, MidpointRounding.ToEven),
                            DSCK = item.TongDS == 0 ? 0 : Math.Round(item.DSCK / sumDSCKLastYear * 100, 2, MidpointRounding.ToEven),
                            Year = item.Year,
                            Month = item.Month
                        };
                    }

                    // Tháng hiện tại
                    if (item.Year == toYear.ToString() && item.Month == toMonth.ToString())
                    {
                        itemDetailtPercent = new ReportDetailtSTMarket()
                        {
                            MarketCode = item.MarketCode,
                            MarketName = item.MarketName,
                            DSChiQuay = item.TongDS == 0 ? 0 : Math.Round(item.DSChiQuay / sumDSChiQuayYear * 100, 2, MidpointRounding.ToEven),
                            DSChiNha = item.TongDS == 0 ? 0 : Math.Round(item.DSChiNha / sumDSChiNhaYear * 100, 2, MidpointRounding.ToEven),
                            DSCK = item.TongDS == 0 ? 0 : Math.Round(item.DSCK / sumDSCKYear * 100, 2, MidpointRounding.ToEven),
                            Year = item.Year,
                            Month = item.Month
                        };
                    }

                    // Tháng trước
                    if (item.Year == toYear.ToString() && item.Month == (toMonth - 1).ToString())
                    {
                        itemDetailtPercent = new ReportDetailtSTMarket()
                        {
                            MarketCode = item.MarketCode,
                            MarketName = item.MarketName,
                            DSChiQuay = item.TongDS == 0 ? 0 : Math.Round(item.DSChiQuay / sumDSChiQuayLastMonth * 100, 2, MidpointRounding.ToEven),
                            DSChiNha = item.TongDS == 0 ? 0 : Math.Round(item.DSChiNha / sumDSChiNhaLastMonth * 100, 2, MidpointRounding.ToEven),
                            DSCK = item.TongDS == 0 ? 0 : Math.Round(item.DSCK / sumDSCKLastMonth * 100, 2, MidpointRounding.ToEven),
                            Year = item.Year,
                            Month = item.Month
                        };
                    }
                    resultConvert.Add(itemDetailtPercent);
                }
                return resultConvert;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// List Report chi tiết cho báo cáo theo tháng
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtServiceType> ReportDetailtCompareMonthForOne(int toYear, int toMonth, string reportTypeID, string marketID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<ReportDetailtServiceType> result = dal.ReportDetailtCompareMonthForOne(toYear, toMonth, reportTypeID, marketID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }


        #region Báo cáo chi tiết thị trường theo loại tiền chi trả

        /// <summary>
        /// List Report detailt theo ngày
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/08/2020]
        /// </history>
        public List<ReportDetailtForTotalMoneyType> ListReportDetailtMTForAll(string reportType)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<ReportDetailtForTotalMoneyType> result = dal.ListReportDetailtMTForAll(reportType);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// List Report
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtForTotalMoneyType> SearchReportDetailtMTForAll(DateTime fromDate, DateTime toDate, string reportTypeID, string marketID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<ReportDetailtForTotalMoneyType> result = dal.SearchReportDetailtMTForAll(fromDate, toDate, reportTypeID, marketID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// List Report chi tiết cho loại tiền cho tháng
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtForTotalMoneyType> SearchReportDetailtMTForAllForMonth(DateTime fromDate, DateTime toDate, string reportTypeID, string marketID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                // get first day in fromMonth
                DateTime fromDateRecent = new DateTime(fromDate.Year, fromDate.Month, 1);

                // get last day in toMonth
                int lastDayInToDate = DateTime.DaysInMonth(toDate.Year, toDate.Month);
                DateTime toDateRecent = new DateTime(toDate.Year, toDate.Month, lastDayInToDate);

                List<ReportDetailtForTotalMoneyType> result = dal.SearchReportDetailtMTForAll(fromDateRecent, toDateRecent, reportTypeID, marketID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// List Report chi tiết cho loại tiền cho tháng
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtForTotalMoneyType> SearchReportDetailtMTForAllForYear(DateTime fromDate, DateTime toDate, string reportTypeID, string marketID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                // get first year
                DateTime fromDateRecent = new DateTime(fromDate.Year, 1, 1);

                // Ngày cuối năm
                DateTime toDateRecent = new DateTime(toDate.Year, 12, 31);

                List<ReportDetailtForTotalMoneyType> result = dal.SearchReportDetailtMTForAll(fromDateRecent, toDateRecent, reportTypeID, marketID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// List Report detailt theo ngày
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/08/2020]
        /// </history>
        public List<ReportDetailtForTotalMoneyType> ListReportDetailtMTForAllConvert(string reportType)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<ReportDetailtForTotalMoneyType> result = dal.ListReportDetailtMTForAllConvert(reportType);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// List Report
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtForTotalMoneyType> SearchReportDetailtMTForAllConvert(DateTime fromDate, DateTime toDate, string reportTypeID, string marketID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<ReportDetailtForTotalMoneyType> result = dal.SearchReportDetailtMTForAllConvert(fromDate, toDate, reportTypeID, marketID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// List Report chi tiết cho loại tiền cho tháng
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtForTotalMoneyType> SearchReportDetailtMTForAllForMonthConvert(DateTime fromDate, DateTime toDate, string reportTypeID, string marketID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                // get first day in fromMonth
                DateTime fromDateRecent = new DateTime(fromDate.Year, fromDate.Month, 1);

                // get last day in toMonth
                int lastDayInToDate = DateTime.DaysInMonth(toDate.Year, toDate.Month);
                DateTime toDateRecent = new DateTime(toDate.Year, toDate.Month, lastDayInToDate);

                List<ReportDetailtForTotalMoneyType> result = dal.SearchReportDetailtMTForAllConvert(fromDateRecent, toDateRecent, reportTypeID, marketID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// List Report chi tiết cho loại tiền cho tháng
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtForTotalMoneyType> SearchReportDetailtMTForAllForYearConvert(DateTime fromDate, DateTime toDate, string reportTypeID, string marketID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();

                // get first year
                DateTime fromDateRecent = new DateTime(fromDate.Year, 1, 1);

                // Ngày cuối năm
                DateTime toDateRecent = new DateTime(toDate.Year, 12, 31);

                List<ReportDetailtForTotalMoneyType> result = dal.SearchReportDetailtMTForAllConvert(fromDateRecent, toDateRecent, reportTypeID, marketID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// List Report detailt theo ngày
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/08/2020]
        /// </history>
        public List<ReportDetailtForTotalMoneyType> ListReportDetailtMTForOne(string reportType)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<ReportDetailtForTotalMoneyType> result = dal.ListReportDetailtMTForOne(reportType);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// List Report
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtForTotalMoneyType> SearchReportDetailtMTForOne(DateTime fromDate, DateTime toDate, string reportTypeID, string marketID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<ReportDetailtForTotalMoneyType> result = dal.SearchReportDetailtMTForOne(fromDate, toDate, reportTypeID, marketID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// List Report detailt theo ngày
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/08/2020]
        /// </history>
        public List<ReportDetailtForTotalMoneyType> ListReportDetailtMTForOneConvert(string reportType)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<ReportDetailtForTotalMoneyType> result = dal.ListReportDetailtMTForOneConvert(reportType);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// List Report
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtForTotalMoneyType> SearchReportDetailtMTForOneConvert(DateTime fromDate, DateTime toDate, string reportTypeID, string marketID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<ReportDetailtForTotalMoneyType> result = dal.SearchReportDetailtMTForOneConvert(fromDate, toDate, reportTypeID, marketID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// List Report cho so sánh giai đoạn theo báo cáo chi tiết
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtForTotalMoneyType> ReportDetailtMTGradationCompareForAllConvert(int ToYear, int typeID, string reportTypeID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<ReportDetailtForTotalMoneyType> result = dal.ReportDetailtMTGradationCompareForAllConvert(ToYear, typeID, reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// List Report cho so sánh giai đoạn theo báo cáo chi tiết
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtForTotalMoneyType> ReportDetailtMTGradationCompareForAllPercent(int ToYear, int typeID, string reportTypeID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<ReportDetailtForTotalMoneyType> result = dal.ReportDetailtMTGradationCompareForAllConvert(ToYear, typeID, reportTypeID);
                List<ReportDetailtForTotalMoneyType> resultConvert = new List<ReportDetailtForTotalMoneyType>();

                foreach (ReportDetailtForTotalMoneyType item in result)
                {
                    item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;
                    ReportDetailtForTotalMoneyType itemDetailtPercent = new ReportDetailtForTotalMoneyType();

                    itemDetailtPercent = new ReportDetailtForTotalMoneyType()
                    {
                        MarketCode = item.MarketCode,
                        MarketName = item.MarketName,
                        VND = item.TongDS == 0 ? 0 : Math.Round((item.VND / item.TongDS) * 100, 2, MidpointRounding.ToEven),
                        USD = item.TongDS == 0 ? 0 : Math.Round((item.USD / item.TongDS) * 100, 2, MidpointRounding.ToEven),
                        EUR = item.TongDS == 0 ? 0 : Math.Round((item.EUR / item.TongDS) * 100, 2, MidpointRounding.ToEven),
                        CAD = item.TongDS == 0 ? 0 : Math.Round((item.CAD / item.TongDS) * 100, 2, MidpointRounding.ToEven),
                        AUD = item.TongDS == 0 ? 0 : Math.Round((item.AUD / item.TongDS) * 100, 2, MidpointRounding.ToEven),
                        GBP = item.TongDS == 0 ? 0 : Math.Round((item.GBP / item.TongDS) * 100, 2, MidpointRounding.ToEven),
                        Year = item.Year
                    };

                    resultConvert.Add(itemDetailtPercent);
                }
                return resultConvert;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// List Report cho so sánh giai đoạn theo báo cáo chi tiết
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtForTotalMoneyType> ReportDetailtMTGradationCompareForOneConvert(int toYear, int typeID, string reportTypeID, string marketID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<ReportDetailtForTotalMoneyType> result = dal.ReportDetailtMTGradationCompareForOneConvert(toYear, typeID, reportTypeID, marketID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }


        /// <summary>
        /// List Report cho so sánh giai đoạn theo báo cáo chi tiết
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtForTotalMoneyType> ReportDetailtMTGradationCompareForOnePercent(int ToYear, int typeID, string reportTypeID, string marketID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<ReportDetailtForTotalMoneyType> result = dal.ReportDetailtMTGradationCompareForOneConvert(ToYear, typeID, reportTypeID, marketID);
                List<ReportDetailtForTotalMoneyType> resultConvert = new List<ReportDetailtForTotalMoneyType>();

                foreach (ReportDetailtForTotalMoneyType item in result)
                {
                    item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;
                    ReportDetailtForTotalMoneyType itemDetailtPercent = new ReportDetailtForTotalMoneyType();

                    itemDetailtPercent = new ReportDetailtForTotalMoneyType()
                    {
                        MarketCode = item.MarketCode,
                        MarketName = item.MarketName,
                        PartnerCode = item.PartnerCode,
                        PartnerName = item.PartnerName,
                        VND = item.TongDS == 0 ? 0 : Math.Round((item.VND / item.TongDS) * 100, 2, MidpointRounding.ToEven),
                        USD = item.TongDS == 0 ? 0 : Math.Round((item.USD / item.TongDS) * 100, 2, MidpointRounding.ToEven),
                        EUR = item.TongDS == 0 ? 0 : Math.Round((item.EUR / item.TongDS) * 100, 2, MidpointRounding.ToEven),
                        CAD = item.TongDS == 0 ? 0 : Math.Round((item.CAD / item.TongDS) * 100, 2, MidpointRounding.ToEven),
                        AUD = item.TongDS == 0 ? 0 : Math.Round((item.AUD / item.TongDS) * 100, 2, MidpointRounding.ToEven),
                        GBP = item.TongDS == 0 ? 0 : Math.Round((item.GBP / item.TongDS) * 100, 2, MidpointRounding.ToEven),
                        Year = item.Year
                    };


                    resultConvert.Add(itemDetailtPercent);
                }
                return resultConvert;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// List Report chi tiết cho báo cáo theo tháng
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtForTotalMoneyType> ReportDetailtMTCompareMonthForAllConvert(int toYear, int toMonth, string reportTypeID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<ReportDetailtForTotalMoneyType> result = dal.ReportDetailtMTCompareMonthForAllConvert(toYear, toMonth, reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// List Report cho so sánh giai đoạn theo báo cáo chi tiết
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtForTotalMoneyType> ColumnChartStackDetailtMTCompareMonthForAllPercent(int toYear, int toMonth, string reportTypeID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<ReportDetailtForTotalMoneyType> result = dal.ReportDetailtMTCompareMonthForAllConvert(toYear, toMonth, reportTypeID);
                List<ReportDetailtForTotalMoneyType> resultConvert = new List<ReportDetailtForTotalMoneyType>();

                foreach(ReportDetailtForTotalMoneyType item in result)
                {
                    item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;

                    ReportDetailtForTotalMoneyType dataItem = new ReportDetailtForTotalMoneyType()
                    {
                        Month = item.Month,
                        Year = item.Year,
                        MarketCode = item.MarketCode,
                        MarketName = item.MarketName,
                        VND = item.TongDS == 0 ? 0 : Math.Round((item.VND / item.TongDS) * 100, 2, MidpointRounding.ToEven),
                        USD = item.TongDS == 0 ? 0 : Math.Round((item.USD / item.TongDS) * 100, 2, MidpointRounding.ToEven),
                        EUR = item.TongDS == 0 ? 0 : Math.Round((item.EUR / item.TongDS) * 100, 2, MidpointRounding.ToEven),
                        CAD = item.TongDS == 0 ? 0 : Math.Round((item.CAD / item.TongDS) * 100, 2, MidpointRounding.ToEven),
                        AUD = item.TongDS == 0 ? 0 : Math.Round((item.AUD / item.TongDS) * 100, 2, MidpointRounding.ToEven),
                        GBP = item.TongDS == 0 ? 0 : Math.Round((item.GBP / item.TongDS) * 100, 2, MidpointRounding.ToEven)
                    };

                    resultConvert.Add(dataItem);
                }
                
                return resultConvert;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// List Report chi tiết cho báo cáo theo tháng
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtForTotalMoneyType> ReportDetailtMTCompareMonthForOneConvert(int year, int month, string reportTypeID, string marketID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<ReportDetailtForTotalMoneyType> result = dal.ReportDetailtMTCompareMonthForOneConvert(year, month, reportTypeID, marketID);
                List<ReportDetailtForTotalMoneyType> resultConvert = new List<ReportDetailtForTotalMoneyType>(result);

                List<string> listTemp = new List<string>();

                foreach (ReportDetailtForTotalMoneyType item in resultConvert)
                {
                    // Cùng kì
                    ReportDetailtForTotalMoneyType dataItemLastYear = result.Find(x => x.PartnerCode == item.PartnerCode && x.Month == month.ToString() && x.Year == (year - 1).ToString());
                    ReportDetailtForTotalMoneyType dataItemYear = result.Find(x => x.PartnerCode == item.PartnerCode && x.Month == month.ToString() && x.Year == year.ToString());
                    ReportDetailtForTotalMoneyType dataItemLastMonth = result.Find(x => x.PartnerCode == item.PartnerCode && x.Month == (month - 1).ToString() && x.Year == year.ToString());

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
                            result.Add(dataItemLastYear);
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
                            result.Add(dataItemYear);
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
                            result.Add(dataItemLastMonth);
                        }

                        // Add partnerCode để kiểm tra
                        listTemp.Add(item.PartnerCode);
                    }
                }
                
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// List Report chi tiết cho báo cáo theo tháng
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtForTotalMoneyType> ReportDetailtMTCompareMonthForOneConvertPercent(int year, int month, string reportTypeID, string marketID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<ReportDetailtForTotalMoneyType> result = dal.ReportDetailtMTCompareMonthForOneConvert(year, month, reportTypeID, marketID);
                List<ReportDetailtForTotalMoneyType> resultConvert = new List<ReportDetailtForTotalMoneyType>(result);
                List<ReportDetailtForTotalMoneyType> resultConvertPercent = new List<ReportDetailtForTotalMoneyType>();

                List<string> listTemp = new List<string>();

                foreach (ReportDetailtForTotalMoneyType item in resultConvert)
                {
                    // Cùng kì
                    ReportDetailtForTotalMoneyType dataItemLastYear = result.Find(x => x.PartnerCode == item.PartnerCode && x.Month == month.ToString() && x.Year == (year - 1).ToString());
                    ReportDetailtForTotalMoneyType dataItemYear = result.Find(x => x.PartnerCode == item.PartnerCode && x.Month == month.ToString() && x.Year == year.ToString());
                    ReportDetailtForTotalMoneyType dataItemLastMonth = result.Find(x => x.PartnerCode == item.PartnerCode && x.Month == (month - 1).ToString() && x.Year == year.ToString());

                    if (!listTemp.Contains(item.PartnerCode))
                    {
                        // Trường hợp năm không có đối tác
                        if (dataItemLastYear == null)
                        {
                            dataItemLastYear = new ReportDetailtForTotalMoneyType();
                            dataItemLastYear.MarketCode = item.MarketCode;
                            dataItemLastYear.MarketName = item.MarketName;
                            dataItemLastYear.PartnerCode = item.PartnerCode;
                            dataItemLastYear.PartnerName = item.PartnerName;
                            dataItemLastYear.Year = (year - 1).ToString();
                            dataItemLastYear.Month = month.ToString();
                            result.Add(dataItemLastYear);
                        }

                        // Trường hợp năm hiện tại không có đối tác
                        if (dataItemYear == null)
                        {
                            dataItemYear = new ReportDetailtForTotalMoneyType();
                            dataItemYear.MarketCode = item.MarketCode;
                            dataItemYear.MarketName = item.MarketName;
                            dataItemYear.PartnerCode = item.PartnerCode;
                            dataItemYear.PartnerName = item.PartnerName;
                            dataItemYear.Year = year.ToString();
                            dataItemYear.Month = month.ToString();
                            result.Add(dataItemYear);
                        }

                        // Trường hợp tháng trước không có
                        if (dataItemLastMonth == null)
                        {
                            dataItemLastMonth = new ReportDetailtForTotalMoneyType();
                            dataItemLastMonth.MarketCode = item.MarketCode;
                            dataItemLastMonth.MarketName = item.MarketName;
                            dataItemLastMonth.PartnerCode = item.PartnerCode;
                            dataItemLastMonth.PartnerName = item.PartnerName;
                            dataItemLastMonth.Year = year.ToString();
                            dataItemLastMonth.Month = (month - 1).ToString();
                            result.Add(dataItemLastMonth);
                        }

                        // Add partnerCode để kiểm tra
                        listTemp.Add(item.PartnerCode);
                    }
                }

                foreach (ReportDetailtForTotalMoneyType item in result)
                {
                    item.TongDS = item.VND + item.USD + item.EUR + item.CAD + item.AUD + item.GBP;

                    double sumVND = item.TongDS == 0 ? 0 : Math.Round((item.VND / item.TongDS) * 100, 2, MidpointRounding.ToEven);
                    double sumUSD = item.TongDS == 0 ? 0 : Math.Round((item.USD / item.TongDS) * 100, 2, MidpointRounding.ToEven);
                    double sumEUR = item.TongDS == 0 ? 0 : Math.Round((item.EUR / item.TongDS) * 100, 2, MidpointRounding.ToEven);
                    double sumCAD = item.TongDS == 0 ? 0 : Math.Round((item.CAD / item.TongDS) * 100, 2, MidpointRounding.ToEven);
                    double sumAUD = item.TongDS == 0 ? 0 : Math.Round((item.AUD / item.TongDS) * 100, 2, MidpointRounding.ToEven);
                    double sumGBP = item.TongDS == 0 ? 0 : Math.Round((item.GBP / item.TongDS) * 100, 2, MidpointRounding.ToEven);

                    ReportDetailtForTotalMoneyType dataItem = new ReportDetailtForTotalMoneyType()
                    {
                        PartnerCode = item.PartnerCode,
                        PartnerName = item.PartnerName,
                        MarketCode = item.MarketCode,
                        MarketName = item.MarketName,
                        VND = sumVND,
                        USD = sumUSD,
                        EUR = sumEUR,
                        CAD = sumCAD,
                        AUD = sumAUD,
                        GBP = sumGBP,
                        Month = item.Month,
                        Year = item.Year
                    };

                    resultConvertPercent.Add(dataItem);
                }
                
                return resultConvertPercent.OrderBy(x=>x.Year).ThenBy(x=>x.Month).ToList();
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        #endregion
    }
}

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
    public class HSReportBL : DongABaseDAL
    {
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
                HSReportDAL dal = new HSReportDAL();
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
        public List<Report> SearchMonth(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            try
            {
                HSReportDAL dal = new HSReportDAL();
                List<Report> result = dal.SearchMonth(fromDate, toDate, reportTypeID);
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
        public List<Report> SearchYear(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            try
            {
                HSReportDAL dal = new HSReportDAL();
                List<Report> result = dal.SearchYear(fromDate, toDate, reportTypeID);
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
        public List<Report> SearchGradationCompare(int ToYear, int typeID, string reportTypeID)
        {
            try
            {
                HSReportDAL dal = new HSReportDAL();
                List<Report> result = dal.SearchGradationCompare(ToYear, typeID, reportTypeID);
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
                HSReportDAL dal = new HSReportDAL();
                List<Report> result = dal.SearchGradationCompare(ToYear, typeID, reportTypeID);
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
        /// List Report cho so sánh last month
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<Report> SearchMonthCompareGrid(int ToYear, int ToMonth, string reportTypeID)
        {
            try
            {
                HSReportDAL dal = new HSReportDAL();
                List<Report> result = dal.SearchMonthCompareGrid(ToYear, ToMonth, reportTypeID);
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
        public List<Report> ListDataLastMonthCompareProportionPercent(int toYear, int toMonth, string reportTypeID)
        {
            try
            {
                HSReportDAL dal = new HSReportDAL();
                List<Report> DataConvertPercent = new List<Report>();
                List<Report> result = dal.SearchMonthCompareGrid(toYear, toMonth, reportTypeID);
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

        #region Get dữ liệu cho Loại tiền

        /// <summary>
        /// Danh sách report theo thị trường
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportForTotalMoneyType> SearchReportHSTotalMoneyTypeForDay(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            try
            {
                HSReportDAL dal = new HSReportDAL();
                List<ReportForTotalMoneyType> result = dal.SearchReportHSTotalMoneyTypeForDay(fromDate, toDate, reportTypeID);
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
        public List<ReportForTotalMoneyType> SearchReportHSTotalMoneyTypeForMonth(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            try
            {
                HSReportDAL dal = new HSReportDAL();
                List<ReportForTotalMoneyType> result = dal.SearchReportHSTotalMoneyTypeForMonth(fromDate, toDate, reportTypeID);
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
        public List<ReportForTotalMoneyType> SearchReportHSTotalMoneyTypeForYear(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            try
            {
                HSReportDAL dal = new HSReportDAL();
                List<ReportForTotalMoneyType> result = dal.SearchReportHSTotalMoneyTypeForYear(fromDate, toDate, reportTypeID);
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
        public List<ReportForTotalMoneyType> SearchReportHSTotalMoneyTypeForGradationCompare(int year, int typeID, string reportTypeID)
        {
            try
            {
                HSReportDAL dal = new HSReportDAL();
                List<ReportForTotalMoneyType> result = dal.SearchReportHSTotalMoneyTypeForGradationCompare(year, typeID, reportTypeID);
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
        public List<ReportForTotalMoneyType> SearchReportHSTotalMoneyTypeForGradationComparePercent(int year, int typeID, string reportTypeID)
        {
            try
            {
                HSReportDAL dal = new HSReportDAL();
                List<ReportForTotalMoneyType> result = dal.SearchReportHSTotalMoneyTypeForGradationCompare(year, typeID, reportTypeID);
                List<ReportForTotalMoneyType> listDataPercent = new List<ReportForTotalMoneyType>();

                if (result.Count.Equals(2))
                {
                    foreach (ReportForTotalMoneyType item in result)
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
        public List<ReportForTotalMoneyType> SearchReportHSTotalMoneyTypeForCompareMonth(int toYear, int toMonth, string reportTypeID)
        {
            try
            {
                HSReportDAL dal = new HSReportDAL();
                List<ReportForTotalMoneyType> result = dal.SearchReportHSTotalMoneyTypeForCompareMonth(toYear, toMonth, reportTypeID);
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
        public List<ReportForTotalMoneyType> SearchReportHSTotalMoneyTypeForCompareMonthPercent(int toYear, int toMonth, string reportTypeID)
        {
            try
            {
                HSReportDAL dal = new HSReportDAL();
                List<ReportForTotalMoneyType> result = dal.SearchReportHSTotalMoneyTypeForCompareMonth(toYear, toMonth, reportTypeID);
                List<ReportForTotalMoneyType> resultPercent = new List<ReportForTotalMoneyType>();
                if (result.Count.Equals(3))
                {
                    foreach (ReportForTotalMoneyType item in result)
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

        #endregion

        #region Get dữ liệu cho thị trường

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
                HSReportDAL dal = new HSReportDAL();
                List<ReportForMaket> result = dal.SearchReportMaketForDay(fromDate, toDate, reportTypeID);
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
        public List<ReportForMaket> SearchReportMaketForMonth(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            try
            {
                HSReportDAL dal = new HSReportDAL();
                List<ReportForMaket> result = dal.SearchReportMaketForMonth(fromDate, toDate, reportTypeID);
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
        public List<ReportForMaket> SearchReportMaketForYear(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            try
            {
                HSReportDAL dal = new HSReportDAL();
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
        public List<ReportForMaket> SearchReportMaketForGradationCompare(int year, int typeID, string reportTypeID)
        {
            try
            {
                HSReportDAL dal = new HSReportDAL();
                List<ReportForMaket> result = dal.SearchReportMaketForGradationCompare(year, typeID, reportTypeID);
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
        public List<ReportForMaket> SearchReportMaketForGradationComparePercent(int year, int typeID, string reportTypeID)
        {
            try
            {
                HSReportDAL dal = new HSReportDAL();
                List<ReportForMaket> listDataPercent = new List<ReportForMaket>();
                List<ReportForMaket> result = dal.SearchReportMaketForGradationCompare(year, typeID, reportTypeID);

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
        public List<ReportForMaket> SearchReportMaketForCompareForMonth(int ToYear, int ToMonth, string reportTypeID)
        {
            try
            {
                HSReportDAL dal = new HSReportDAL();
                List<ReportForMaket> result = dal.SearchReportMaketForCompareMonth(ToYear, ToMonth, reportTypeID);
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
        public List<ReportForMaket> SearchReportMaketForCompareMonthPercent(int ToYear, int ToMonth, string reportTypeID)
        {
            try
            {
                HSReportDAL dal = new HSReportDAL();
                List<ReportForMaket> result = dal.SearchReportMaketForCompareMonth(ToYear, ToMonth, reportTypeID);
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
        #endregion

        #region Get dữ liệu cho tổng hồ sơ

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
                HSReportDAL dal = new HSReportDAL();
                List<ReportForTotalPayment> result = dal.SearchReportTotalHSForDay(fromDate, toDate, reportTypeID);
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
        public List<ReportForTotalPayment> SearchReportTPForMonth(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            try
            {
                HSReportDAL dal = new HSReportDAL();
                List<ReportForTotalPayment> result = dal.SearchReportTotalHSForMonth(fromDate, toDate, reportTypeID);
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
        public List<ReportForTotalPayment> SearchReportTPForYear(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            try
            {
                HSReportDAL dal = new HSReportDAL();
                List<ReportForTotalPayment> result = dal.SearchReportTotalHSForYear(fromDate, toDate, reportTypeID);
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
        public List<ReportForTotalPayment> SearchReportTotalHSForGradationCompare(int toYear, int typeID, string reportTypeID)
        {
            try
            {
                HSReportDAL dal = new HSReportDAL();
                List<ReportForTotalPayment> result = dal.SearchReportTotalHSForGradationCompare(toYear, typeID, reportTypeID);
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
        public List<ReportForTotalPayment> SearchReportTotalHSForCompareMonth(int toYear, int toMonth, string reportTypeID)
        {
            try
            {
                HSReportDAL dal = new HSReportDAL();
                List<ReportForTotalPayment> result = dal.SearchReportTotalHSForCompareMonth(toYear, toMonth, reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }
        #endregion

        #region Get dữ liệu hồ sơ loại hình

        /// <summary>
        /// List Report
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtSTMarket> SearchMarketForTotalForDay(DateTime fromDate, DateTime toDate, string reportTypeID, string marketID)
        {
            try
            {
                HSReportDAL dal = new HSReportDAL();
                List<ReportDetailtSTMarket> result = dal.SearchMarketForTotalForDay(fromDate, toDate, reportTypeID, marketID);
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
                HSReportDAL dal = new HSReportDAL();
                List<ReportDetailtSTMarket> result = dal.SearchMarketForTotalForMonth(fromDate, toDate, reportTypeID, marketID);
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
        public List<ReportDetailtSTMarket> SearchMarketForTotalForYear(DateTime fromDate, DateTime toDate, string reportTypeID, string marketID)
        {
            try
            {
                HSReportDAL dal = new HSReportDAL();
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
        public List<ReportDetailtServiceType> SearchMarketForOne(DateTime fromDate, DateTime toDate, string reportTypeID, string marketID)
        {
            try
            {
                HSReportDAL dal = new HSReportDAL();
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
        public List<ReportDetailtServiceType> SearchMarketForOneforMonth(DateTime fromDate, DateTime toDate, string reportTypeID, string marketID)
        {
            try
            {
                HSReportDAL dal = new HSReportDAL();
                List<ReportDetailtServiceType> result = dal.SearchMarketForOneForMonth(fromDate, toDate, reportTypeID, marketID);
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
        public List<ReportDetailtServiceType> SearchMarketForOneForYear(DateTime fromDate, DateTime toDate, string reportTypeID, string marketID)
        {
            try
            {
                HSReportDAL dal = new HSReportDAL();
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
                HSReportDAL dal = new HSReportDAL();
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
                HSReportDAL dal = new HSReportDAL();
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
                HSReportDAL dal = new HSReportDAL();
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
        public List<ReportDetailtServiceType> ReportDetailtGradationCompareForOnePercent(int toYear, int typeID, string reportTypeID, string marketID)
        {
            try
            {
                HSReportDAL dal = new HSReportDAL();
                List<ReportDetailtServiceType> result = dal.ReportDetailtGradationCompareForOne(toYear, typeID, reportTypeID, marketID);
                List<ReportDetailtServiceType> resultConvert = new List<ReportDetailtServiceType>();

                foreach (ReportDetailtServiceType item in result)
                {
                    item.TongDS = item.DSChiQuay + item.DSChiNha + item.DSCK;
                }

                // Year
                List<ReportDetailtServiceType> listDataItemYear = result.Where(x => x.Year == toYear.ToString()).ToList();
                double sumDSChiQuayTotalYear = listDataItemYear.Sum(x => x.DSChiQuay);
                double sumDSChiNhaTotalYear = listDataItemYear.Sum(x => x.DSChiNha);
                double sumDSCKTotalYear = listDataItemYear.Sum(x => x.DSCK);
                double sumTongDSTotalYear = listDataItemYear.Sum(x => x.TongDS);

                // Last Year
                List<ReportDetailtServiceType> listDataItemLastYear = result.Where(x => x.Year == (toYear - 1).ToString()).ToList();
                double sumDSChiQuayTotalLastYear = listDataItemLastYear.Sum(x => x.DSChiQuay);
                double sumDSChiNhaTotalLastYear = listDataItemLastYear.Sum(x => x.DSChiNha);
                double sumDSCKTotalLastYear = listDataItemLastYear.Sum(x => x.DSCK);
                double sumTongDSTotalLastYear = listDataItemLastYear.Sum(x => x.TongDS);

                foreach (ReportDetailtServiceType item in result)
                {
                    ReportDetailtServiceType itemDetailtPercent = new ReportDetailtServiceType();

                    // Last year
                    if (item.Year == (toYear - 1).ToString())
                    {
                        itemDetailtPercent = new ReportDetailtServiceType()
                        {
                            MarketCode = item.MarketCode,
                            MarketName = item.MarketName,
                            PartnerCode = item.PartnerCode,
                            PartnerName = item.PartnerName,
                            DSChiQuay = sumDSChiQuayTotalLastYear == 0 ? 0 : Math.Round((item.DSChiQuay / sumDSChiQuayTotalLastYear) * 100, 2, MidpointRounding.ToEven),
                            DSChiNha = sumDSChiNhaTotalLastYear == 0 ? 0 : Math.Round((item.DSChiNha / sumDSChiNhaTotalLastYear) * 100, 2, MidpointRounding.ToEven),
                            DSCK = sumDSCKTotalLastYear == 0 ? 0 : Math.Round((item.DSCK / sumDSCKTotalLastYear) * 100, 2, MidpointRounding.ToEven),
                            TongDS = sumTongDSTotalLastYear == 0 ? 0 : Math.Round((item.TongDS / sumTongDSTotalLastYear) * 100, 2, MidpointRounding.ToEven),
                            Year = (toYear - 1).ToString()
                        };
                    }

                    // year hien tai
                    if (item.Year == toYear.ToString())
                    {
                        itemDetailtPercent = new ReportDetailtServiceType()
                        {
                            MarketCode = item.MarketCode,
                            MarketName = item.MarketName,
                            PartnerCode = item.PartnerCode,
                            PartnerName = item.PartnerName,
                            DSChiQuay = sumDSChiQuayTotalYear == 0 ? 0 : Math.Round((item.DSChiQuay / sumDSChiQuayTotalYear) * 100, 2, MidpointRounding.ToEven),
                            DSChiNha = sumDSChiNhaTotalYear == 0 ? 0 : Math.Round((item.DSChiNha / sumDSChiNhaTotalYear) * 100, 2, MidpointRounding.ToEven),
                            DSCK = sumDSCKTotalYear == 0 ? 0 : Math.Round((item.DSCK / sumDSCKTotalYear) * 100, 2, MidpointRounding.ToEven),
                            TongDS = sumTongDSTotalYear == 0 ? 0 : Math.Round((item.TongDS / sumTongDSTotalYear) * 100, 2, MidpointRounding.ToEven),
                            Year = toYear.ToString()
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
                HSReportDAL dal = new HSReportDAL();
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
                HSReportDAL dal = new HSReportDAL();
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

                if (marketID.Equals("0"))
                {
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
                        // Trường hợp tháng = 1
                        if (toMonth == 1)
                        {
                            sumDSChiQuayLastMonth = result.Where(x => x.Year == (toYear - 1).ToString() && x.Month == "12").Sum(y => y.DSChiQuay);
                            sumDSChiNhaLastMonth = result.Where(x => x.Year == (toYear - 1).ToString() && x.Month == "12").Sum(y => y.DSChiNha);
                            sumDSCKLastMonth = result.Where(x => x.Year == (toYear - 1).ToString() && x.Month == "12").Sum(y => y.DSCK);
                        }
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
                        if (toMonth == 1)
                        {
                            if (item.Year == (toYear - 1).ToString() && item.Month == "12")
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
                        }
                        else
                        {
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
                        }
                        resultConvert.Add(itemDetailtPercent);
                    }
                }
                else
                {
                    // List các thị trường của châu Á
                    List<string> listAsianItem = new List<string>();
                    foreach (ReportDetailtSTMarket item in result)
                    {
                        if (!listAsianItem.Contains(item.MarketName))
                        {
                            listAsianItem.Add(item.MarketName);
                        }
                    }

                    List<ReportDetailtSTMarket> listDataYear = result.Where(x => x.Month == toMonth.ToString() && x.Year == toYear.ToString()).ToList();
                    List<ReportDetailtSTMarket> listDataLastMonth = result.Where(x => x.Month == (toMonth - 1).ToString() && x.Year == toYear.ToString()).ToList();
                    // Trường hợp tháng 1
                    if (toMonth == 1)
                    {
                        listDataLastMonth = result.Where(x => x.Month == "12" && x.Year == (toYear - 1).ToString()).ToList();
                    }
                    List<ReportDetailtSTMarket> listDataLastYear = result.Where(x => x.Month == toMonth.ToString() && x.Year == (toYear - 1).ToString()).ToList();

                    // Tháng hiện tại
                    double sumDSChiQuayTotalYear = listDataYear.Sum(x => x.DSChiQuay);
                    double sumDSChiNhaTotalYear = listDataYear.Sum(x => x.DSChiNha);
                    double sumDSCKTotalYear = listDataYear.Sum(x => x.DSCK);

                    // Tháng trước
                    double sumDSChiQuayTotalLastMonth = listDataLastMonth.Sum(x => x.DSChiQuay);
                    double sumDSChiNhaTotalLastMonth = listDataLastMonth.Sum(x => x.DSChiNha);
                    double sumDSCKTotalLastMonth = listDataLastMonth.Sum(x => x.DSCK);

                    // Cùng kì
                    double sumDSChiQuayTotalLastYear = listDataLastYear.Sum(x => x.DSChiQuay);
                    double sumDSChiNhaTotalLastYear = listDataLastYear.Sum(x => x.DSChiNha);
                    double sumDSCKTotalLastYear = listDataLastYear.Sum(x => x.DSCK);

                    // Duyệt các thị trường con của Châu Á
                    foreach (string item in listAsianItem)
                    {
                        listDataYear = result.Where(x => x.MarketName == item && x.Month == toMonth.ToString() && x.Year == toYear.ToString()).ToList();
                        listDataLastMonth = result.Where(x => x.MarketName == item && x.Month == (toMonth - 1).ToString() && x.Year == toYear.ToString()).ToList();
                        // Trường hợp tháng 1
                        if (toMonth == 1)
                        {
                            listDataLastMonth = result.Where(x => x.MarketName == item && x.Month == "12" && x.Year == (toYear - 1).ToString()).ToList();
                        }

                        listDataLastYear = result.Where(x => x.MarketName == item && x.Month == toMonth.ToString() && x.Year == (toYear - 1).ToString()).ToList();

                        // Tháng hiện tại
                        sumDSChiQuayYear = listDataYear.Sum(x => x.DSChiQuay);
                        sumDSChiNhaYear = listDataYear.Sum(x => x.DSChiNha);
                        sumDSCKYear = listDataYear.Sum(x => x.DSCK);

                        // Tháng trước
                        sumDSChiQuayLastMonth = listDataLastMonth.Sum(x => x.DSChiQuay);
                        sumDSChiNhaLastMonth = listDataLastMonth.Sum(x => x.DSChiNha);
                        sumDSCKLastMonth = listDataLastMonth.Sum(x => x.DSCK);

                        // Cùng kì
                        sumDSChiQuayLastYear = listDataLastYear.Sum(x => x.DSChiQuay);
                        sumDSChiNhaLastYear = listDataLastYear.Sum(x => x.DSChiNha);
                        sumDSCKLastYear = listDataLastYear.Sum(x => x.DSCK);

                        ReportDetailtSTMarket itemDetailtPercent = new ReportDetailtSTMarket();

                        // Tháng hiện tại
                        itemDetailtPercent = new ReportDetailtSTMarket()
                        {
                            MarketName = item,
                            DSChiQuay = sumDSChiQuayTotalYear == 0 ? 0 : Math.Round(sumDSChiQuayYear / sumDSChiQuayTotalYear * 100, 2, MidpointRounding.ToEven),
                            DSChiNha = sumDSChiNhaTotalYear == 0 ? 0 : Math.Round(sumDSChiNhaYear / sumDSChiNhaTotalYear * 100, 2, MidpointRounding.ToEven),
                            DSCK = sumDSCKTotalYear == 0 ? 0 : Math.Round(sumDSCKYear / sumDSCKTotalYear * 100, 2, MidpointRounding.ToEven),
                            Year = toYear.ToString(),
                            Month = toMonth.ToString()
                        };

                        resultConvert.Add(itemDetailtPercent);

                        // Cùng kì
                        itemDetailtPercent = new ReportDetailtSTMarket()
                        {
                            MarketName = item,
                            DSChiQuay = sumDSChiQuayTotalLastYear == 0 ? 0 : Math.Round(sumDSChiQuayLastYear / sumDSChiQuayTotalLastYear * 100, 2, MidpointRounding.ToEven),
                            DSChiNha = sumDSChiNhaTotalLastYear == 0 ? 0 : Math.Round(sumDSChiNhaLastYear / sumDSChiNhaTotalLastYear * 100, 2, MidpointRounding.ToEven),
                            DSCK = sumDSCKTotalLastYear == 0 ? 0 : Math.Round(sumDSCKLastYear / sumDSCKTotalLastYear * 100, 2, MidpointRounding.ToEven),
                            Year = (toYear - 1).ToString(),
                            Month = toMonth.ToString()
                        };

                        resultConvert.Add(itemDetailtPercent);

                        if (toMonth == 1)
                        {
                            // Tháng trước
                            itemDetailtPercent = new ReportDetailtSTMarket()
                            {
                                MarketName = item,
                                DSChiQuay = sumDSChiQuayTotalLastMonth == 0 ? 0 : Math.Round(sumDSChiQuayLastMonth / sumDSChiQuayTotalLastMonth * 100, 2, MidpointRounding.ToEven),
                                DSChiNha = sumDSChiNhaTotalLastMonth == 0 ? 0 : Math.Round(sumDSChiNhaLastMonth / sumDSChiNhaTotalLastMonth * 100, 2, MidpointRounding.ToEven),
                                DSCK = sumDSCKTotalLastMonth == 0 ? 0 : Math.Round(sumDSCKLastMonth / sumDSCKTotalLastMonth * 100, 2, MidpointRounding.ToEven),
                                Year = (toYear - 1).ToString(),
                                Month = "12"
                            };
                        }
                        else
                        {
                            // Tháng trước
                            itemDetailtPercent = new ReportDetailtSTMarket()
                            {
                                MarketName = item,
                                DSChiQuay = sumDSChiQuayTotalLastMonth == 0 ? 0 : Math.Round(sumDSChiQuayLastMonth / sumDSChiQuayTotalLastMonth * 100, 2, MidpointRounding.ToEven),
                                DSChiNha = sumDSChiNhaTotalLastMonth == 0 ? 0 : Math.Round(sumDSChiNhaLastMonth / sumDSChiNhaTotalLastMonth * 100, 2, MidpointRounding.ToEven),
                                DSCK = sumDSCKTotalLastMonth == 0 ? 0 : Math.Round(sumDSCKLastMonth / sumDSCKTotalLastMonth * 100, 2, MidpointRounding.ToEven),
                                Year = toYear.ToString(),
                                Month = (toMonth - 1).ToString()
                            };
                        }

                        resultConvert.Add(itemDetailtPercent);
                    }
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
                HSReportDAL dal = new HSReportDAL();
                List<ReportDetailtServiceType> result = dal.ReportDetailtCompareMonthForOne(toYear, toMonth, reportTypeID, marketID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        #endregion

        #region Get dữ liệu Hồ sơ chi tiết cho thị trường loại tiền

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
                HSReportDAL dal = new HSReportDAL();
                List<ReportDetailtForTotalMoneyType> result = dal.SearchReportDetailtMTForAll(fromDate, toDate, reportTypeID, marketID);
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
        public List<ReportDetailtForTotalMoneyType> SearchReportDetailtMTForAllForMonth(DateTime fromDate, DateTime toDate, string reportTypeID, string marketID)
        {
            try
            {
                HSReportDAL dal = new HSReportDAL();
                List<ReportDetailtForTotalMoneyType> result = dal.SearchReportDetailtMTForAllForMonth(fromDate, toDate, reportTypeID, marketID);
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
        public List<ReportDetailtForTotalMoneyType> SearchReportDetailtMTForAllForYear(DateTime fromDate, DateTime toDate, string reportTypeID, string marketID)
        {
            try
            {
                HSReportDAL dal = new HSReportDAL();
                List<ReportDetailtForTotalMoneyType> result = dal.SearchReportDetailtMTForAllForYear(fromDate, toDate, reportTypeID, marketID);
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
                HSReportDAL dal = new HSReportDAL();
                List<ReportDetailtForTotalMoneyType> result = dal.SearchReportDetailtMTForOne(fromDate, toDate, reportTypeID, marketID);
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
        public List<ReportDetailtForTotalMoneyType> SearchReportDetailtMTForOneForMonth(DateTime fromDate, DateTime toDate, string reportTypeID, string marketID)
        {
            try
            {
                HSReportDAL dal = new HSReportDAL();
                List<ReportDetailtForTotalMoneyType> result = dal.SearchReportDetailtMTForOneForMonth(fromDate, toDate, reportTypeID, marketID);
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
        public List<ReportDetailtForTotalMoneyType> SearchReportDetailtMTForOneForYear(DateTime fromDate, DateTime toDate, string reportTypeID, string marketID)
        {
            try
            {
                HSReportDAL dal = new HSReportDAL();
                List<ReportDetailtForTotalMoneyType> result = dal.SearchReportDetailtMTForOneForYear(fromDate, toDate, reportTypeID, marketID);
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
        public List<ReportDetailtForTotalMoneyType> ReportDetailtMTGradationCompareForAll(int ToYear, int typeID, string reportTypeID, string marketID)
        {
            try
            {
                HSReportDAL dal = new HSReportDAL();
                List<ReportDetailtForTotalMoneyType> result = dal.ReportDetailtMTGradationCompareForAll(ToYear, typeID, reportTypeID, marketID);
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
        public List<ReportDetailtForTotalMoneyType> ReportDetailtMTGradationCompareForAllPercent(int toYear, int typeID, string reportTypeID, string marketID)
        {
            try
            {
                HSReportDAL dal = new HSReportDAL();
                List<ReportDetailtForTotalMoneyType> result = dal.ReportDetailtMTGradationCompareForAll(toYear, typeID, reportTypeID, marketID);
                List<ReportDetailtForTotalMoneyType> resultConvert = new List<ReportDetailtForTotalMoneyType>();



                // Trường hợp chọn tất cả thị trường
                if (marketID.Equals("0"))
                {
                    List<ReportDetailtForTotalMoneyType> listDataYear = result.Where(x => x.Year == toYear.ToString()).ToList();
                    List<ReportDetailtForTotalMoneyType> listDataLastYear = result.Where(x => x.Year == (toYear - 1).ToString()).ToList();

                    // Year
                    double sumVNDYear = listDataYear.Sum(x => x.VND);
                    double sumUSDYear = listDataYear.Sum(x => x.USD);
                    double sumEURYear = listDataYear.Sum(x => x.EUR);
                    double sumCADYear = listDataYear.Sum(x => x.CAD);
                    double sumAUDYear = listDataYear.Sum(x => x.AUD);
                    double sumGBPYear = listDataYear.Sum(x => x.GBP);

                    double sumTongYear = sumVNDYear + sumUSDYear + sumEURYear + sumCADYear + sumAUDYear + sumGBPYear;

                    // Last Year
                    double sumVNDLastYear = listDataLastYear.Sum(x => x.VND);
                    double sumUSDLastYear = listDataLastYear.Sum(x => x.USD);
                    double sumEURLastYear = listDataLastYear.Sum(x => x.EUR);
                    double sumCADLastYear = listDataLastYear.Sum(x => x.CAD);
                    double sumAUDLastYear = listDataLastYear.Sum(x => x.AUD);
                    double sumGBPLastYear = listDataLastYear.Sum(x => x.GBP);

                    double sumTongLastYear = sumVNDLastYear + sumUSDLastYear + sumEURLastYear + sumCADLastYear + sumAUDLastYear + sumGBPLastYear;

                    foreach (ReportDetailtForTotalMoneyType item in result)
                    {
                        if (item.Year == toYear.ToString())
                        {
                            ReportDetailtForTotalMoneyType dataItem = result.Find(x => x.MarketName == item.MarketName && x.Year == toYear.ToString());
                            resultConvert.Add(
                                new ReportDetailtForTotalMoneyType()
                                {
                                    MarketCode = item.MarketCode,
                                    MarketName = item.MarketName,
                                    VND = sumVNDYear == 0 ? 0 : Math.Round((item.VND / sumVNDYear) * 100, 2, MidpointRounding.ToEven),
                                    USD = sumUSDYear == 0 ? 0 : Math.Round((item.USD / sumUSDYear) * 100, 2, MidpointRounding.ToEven),
                                    EUR = sumEURYear == 0 ? 0 : Math.Round((item.EUR / sumEURYear) * 100, 2, MidpointRounding.ToEven),
                                    CAD = sumCADYear == 0 ? 0 : Math.Round((item.CAD / sumCADYear) * 100, 2, MidpointRounding.ToEven),
                                    AUD = sumAUDYear == 0 ? 0 : Math.Round((item.AUD / sumAUDYear) * 100, 2, MidpointRounding.ToEven),
                                    GBP = sumGBPYear == 0 ? 0 : Math.Round((item.GBP / sumGBPYear) * 100, 2, MidpointRounding.ToEven),
                                    Year = item.Year
                                }
                            );
                        }
                        else
                        {
                            ReportDetailtForTotalMoneyType dataItem = result.Find(x => x.MarketName == item.MarketName && x.Year == (toYear - 1).ToString());
                            resultConvert.Add(
                                new ReportDetailtForTotalMoneyType()
                                {
                                    MarketCode = item.MarketCode,
                                    MarketName = item.MarketName,
                                    VND = sumVNDLastYear == 0 ? 0 : Math.Round((item.VND / sumVNDLastYear) * 100, 2, MidpointRounding.ToEven),
                                    USD = sumUSDLastYear == 0 ? 0 : Math.Round((item.USD / sumUSDLastYear) * 100, 2, MidpointRounding.ToEven),
                                    EUR = sumEURLastYear == 0 ? 0 : Math.Round((item.EUR / sumEURLastYear) * 100, 2, MidpointRounding.ToEven),
                                    CAD = sumCADLastYear == 0 ? 0 : Math.Round((item.CAD / sumCADLastYear) * 100, 2, MidpointRounding.ToEven),
                                    AUD = sumAUDLastYear == 0 ? 0 : Math.Round((item.AUD / sumAUDLastYear) * 100, 2, MidpointRounding.ToEven),
                                    GBP = sumGBPLastYear == 0 ? 0 : Math.Round((item.GBP / sumGBPLastYear) * 100, 2, MidpointRounding.ToEven),
                                    Year = item.Year
                                }
                            );
                        }
                    }
                }
                else
                {
                    List<string> listMarket = new List<string>();

                    foreach (ReportDetailtForTotalMoneyType item in result)
                    {
                        if (!listMarket.Contains(item.MarketName))
                        {
                            listMarket.Add(item.MarketName);
                        }
                    }

                    List<ReportDetailtForTotalMoneyType> listDataTotalYear = result.Where(x => x.Year == toYear.ToString()).ToList();
                    List<ReportDetailtForTotalMoneyType> listDataTotalLastYear = result.Where(x => x.Year == (toYear - 1).ToString()).ToList();

                    // Year
                    double sumVNDTotalYear = listDataTotalYear.Sum(x => x.VND);
                    double sumUSDTotalYear = listDataTotalYear.Sum(x => x.USD);
                    double sumEURTotalYear = listDataTotalYear.Sum(x => x.EUR);
                    double sumCADTotalYear = listDataTotalYear.Sum(x => x.CAD);
                    double sumAUDTotalYear = listDataTotalYear.Sum(x => x.AUD);
                    double sumGBPTotalYear = listDataTotalYear.Sum(x => x.GBP);

                    double sumTongTotalYear = sumVNDTotalYear + sumUSDTotalYear + sumEURTotalYear + sumCADTotalYear + sumAUDTotalYear + sumGBPTotalYear;

                    // Last Year
                    double sumVNDTotalLastYear = listDataTotalLastYear.Sum(x => x.VND);
                    double sumUSDTotalLastYear = listDataTotalLastYear.Sum(x => x.USD);
                    double sumEURTotalLastYear = listDataTotalLastYear.Sum(x => x.EUR);
                    double sumCADTotalLastYear = listDataTotalLastYear.Sum(x => x.CAD);
                    double sumAUDTotalLastYear = listDataTotalLastYear.Sum(x => x.AUD);
                    double sumGBPTotalLastYear = listDataTotalLastYear.Sum(x => x.GBP);

                    double sumTongTotalLastYear = sumVNDTotalLastYear + sumUSDTotalLastYear + sumEURTotalLastYear + sumCADTotalLastYear + sumAUDTotalLastYear + sumGBPTotalLastYear;

                    // Danh sách thị trường
                    foreach (string item in listMarket)
                    {
                        List<ReportDetailtForTotalMoneyType> listDataYear = result.Where(x => x.Year == toYear.ToString() && x.MarketName == item).ToList();
                        List<ReportDetailtForTotalMoneyType> listDataLastYear = result.Where(x => x.Year == (toYear - 1).ToString() && x.MarketName == item).ToList();

                        // Year
                        double sumVNDYear = listDataYear.Sum(x => x.VND);
                        double sumUSDYear = listDataYear.Sum(x => x.USD);
                        double sumEURYear = listDataYear.Sum(x => x.EUR);
                        double sumCADYear = listDataYear.Sum(x => x.CAD);
                        double sumAUDYear = listDataYear.Sum(x => x.AUD);
                        double sumGBPYear = listDataYear.Sum(x => x.GBP);

                        double sumTongYear = sumVNDYear + sumUSDYear + sumEURYear + sumCADYear + sumAUDYear + sumGBPYear;

                        resultConvert.Add(
                                new ReportDetailtForTotalMoneyType()
                                {
                                    MarketName = item,
                                    VND = sumVNDTotalYear == 0 ? 0 : Math.Round((sumVNDYear / sumVNDTotalYear) * 100, 2, MidpointRounding.ToEven),
                                    USD = sumUSDTotalYear == 0 ? 0 : Math.Round((sumUSDYear / sumUSDTotalYear) * 100, 2, MidpointRounding.ToEven),
                                    EUR = sumEURTotalYear == 0 ? 0 : Math.Round((sumEURYear / sumEURTotalYear) * 100, 2, MidpointRounding.ToEven),
                                    CAD = sumCADTotalYear == 0 ? 0 : Math.Round((sumCADYear / sumCADTotalYear) * 100, 2, MidpointRounding.ToEven),
                                    AUD = sumAUDTotalYear == 0 ? 0 : Math.Round((sumAUDYear / sumAUDTotalYear) * 100, 2, MidpointRounding.ToEven),
                                    GBP = sumGBPTotalYear == 0 ? 0 : Math.Round((sumGBPYear / sumGBPTotalYear) * 100, 2, MidpointRounding.ToEven),
                                    Year = toYear.ToString()
                                }
                            );

                        // Last Year
                        double sumVNDLastYear = listDataLastYear.Sum(x => x.VND);
                        double sumUSDLastYear = listDataLastYear.Sum(x => x.USD);
                        double sumEURLastYear = listDataLastYear.Sum(x => x.EUR);
                        double sumCADLastYear = listDataLastYear.Sum(x => x.CAD);
                        double sumAUDLastYear = listDataLastYear.Sum(x => x.AUD);
                        double sumGBPLastYear = listDataLastYear.Sum(x => x.GBP);

                        double sumTongLastYear = sumVNDLastYear + sumUSDLastYear + sumEURLastYear + sumCADLastYear + sumAUDLastYear + sumGBPLastYear;

                        resultConvert.Add(
                                new ReportDetailtForTotalMoneyType()
                                {
                                    MarketName = item,
                                    VND = sumVNDTotalLastYear == 0 ? 0 : Math.Round((sumVNDLastYear / sumVNDTotalLastYear) * 100, 2, MidpointRounding.ToEven),
                                    USD = sumUSDTotalLastYear == 0 ? 0 : Math.Round((sumUSDLastYear / sumUSDTotalLastYear) * 100, 2, MidpointRounding.ToEven),
                                    EUR = sumEURTotalLastYear == 0 ? 0 : Math.Round((sumEURLastYear / sumEURTotalLastYear) * 100, 2, MidpointRounding.ToEven),
                                    CAD = sumCADTotalLastYear == 0 ? 0 : Math.Round((sumCADLastYear / sumCADTotalLastYear) * 100, 2, MidpointRounding.ToEven),
                                    AUD = sumAUDTotalLastYear == 0 ? 0 : Math.Round((sumAUDLastYear / sumAUDTotalLastYear) * 100, 2, MidpointRounding.ToEven),
                                    GBP = sumGBPTotalLastYear == 0 ? 0 : Math.Round((sumGBPLastYear / sumGBPTotalLastYear) * 100, 2, MidpointRounding.ToEven),
                                    Year = (toYear - 1).ToString()
                                }
                            );

                    }

                    // Sau khi data theo thị trường

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
        public List<ReportDetailtForTotalMoneyType> ReportDetailtMTGradationCompareForOne(int toYear, int typeID, string reportTypeID, string marketID)
        {
            try
            {
                HSReportDAL dal = new HSReportDAL();
                List<ReportDetailtForTotalMoneyType> result = dal.ReportDetailtMTGradationCompareForOne(toYear, typeID, reportTypeID, marketID);
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
        public List<ReportDetailtForTotalMoneyType> ReportDetailtMTGradationCompareForOnePercent(int toYear, int typeID, string reportTypeID, string marketID)
        {
            try
            {
                HSReportDAL dal = new HSReportDAL();
                List<ReportDetailtForTotalMoneyType> result = dal.ReportDetailtMTGradationCompareForOne(toYear, typeID, reportTypeID, marketID);
                List<ReportDetailtForTotalMoneyType> resultClone = new List<ReportDetailtForTotalMoneyType>(result);
                List<ReportDetailtForTotalMoneyType> resultConvert = new List<ReportDetailtForTotalMoneyType>();

                if (result.Count > 0)
                {
                    // Trường hợp chọn tất cả thị trường
                    if (!marketID.Equals("005"))
                    {
                        List<string> listPartner = new List<string>();
                        foreach (ReportDetailtForTotalMoneyType item in resultClone)
                        {
                            if (listPartner.Contains(item.PartnerCode))
                            {
                                continue;
                            }
                            ReportDetailtForTotalMoneyType dataItemYear = resultClone.Find(x => x.PartnerName == item.PartnerName && x.Year == toYear.ToString());
                            ReportDetailtForTotalMoneyType dataItemLastYear = resultClone.Find(x => x.PartnerName == item.PartnerName && x.Year == (toYear - 1).ToString());

                            if (dataItemYear == null)
                            {
                                result.Add(
                                 new ReportDetailtForTotalMoneyType()
                                 {
                                     PartnerCode = item.PartnerCode,
                                     PartnerName = item.PartnerName,
                                     MarketCode = item.MarketCode,
                                     MarketName = item.MarketName,
                                     Year = toYear.ToString()
                                 }
                                );
                            }

                            if (dataItemLastYear == null)
                            {
                                result.Add(
                                 new ReportDetailtForTotalMoneyType()
                                 {
                                     PartnerCode = item.PartnerCode,
                                     PartnerName = item.PartnerName,
                                     MarketCode = item.MarketCode,
                                     MarketName = item.MarketName,
                                     Year = (toYear - 1).ToString()
                                 }
                                );
                            }

                            listPartner.Add(item.PartnerCode);
                        }


                        List<ReportDetailtForTotalMoneyType> listDataYear = result.Where(x => x.Year == toYear.ToString()).ToList();
                        List<ReportDetailtForTotalMoneyType> listDataLastYear = result.Where(x => x.Year == (toYear - 1).ToString()).ToList();

                        // Year 
                        double sumVNDTotalYear = listDataYear.Sum(x => x.VND);
                        double sumUSDTotalYear = listDataYear.Sum(x => x.USD);
                        double sumEURTotalYear = listDataYear.Sum(x => x.EUR);
                        double sumCADTotalYear = listDataYear.Sum(x => x.CAD);
                        double sumAUDTotalYear = listDataYear.Sum(x => x.AUD);
                        double sumGBPTotalYear = listDataYear.Sum(x => x.GBP);

                        double sumTongDSTotalYear = listDataYear.Sum(x => x.TongDS);

                        // Last Year 
                        double sumVNDTotalLastYear = listDataLastYear.Sum(x => x.VND);
                        double sumUSDTotalLastYear = listDataLastYear.Sum(x => x.USD);
                        double sumEURTotalLastYear = listDataLastYear.Sum(x => x.EUR);
                        double sumCADTotalLastYear = listDataLastYear.Sum(x => x.CAD);
                        double sumAUDTotalLastYear = listDataLastYear.Sum(x => x.AUD);
                        double sumGBPTotalLastYear = listDataLastYear.Sum(x => x.GBP);

                        double sumTongDSTotalLastYear = listDataLastYear.Sum(x => x.TongDS);

                        // Duyệt danh sách các đối tác
                        foreach (ReportDetailtForTotalMoneyType item in result)
                        {
                            if (item.Year == toYear.ToString())
                            {
                                ReportDetailtForTotalMoneyType itemDetailtPercent = new ReportDetailtForTotalMoneyType()
                                {
                                    MarketCode = item.MarketCode,
                                    MarketName = item.MarketName,
                                    PartnerCode = item.PartnerCode,
                                    PartnerName = item.PartnerName,
                                    VND = sumVNDTotalYear == 0 ? 0 : Math.Round((item.VND / sumVNDTotalYear) * 100, 2, MidpointRounding.ToEven),
                                    USD = sumUSDTotalYear == 0 ? 0 : Math.Round((item.USD / sumUSDTotalYear) * 100, 2, MidpointRounding.ToEven),
                                    EUR = sumEURTotalYear == 0 ? 0 : Math.Round((item.EUR / sumEURTotalYear) * 100, 2, MidpointRounding.ToEven),
                                    CAD = sumCADTotalYear == 0 ? 0 : Math.Round((item.CAD / sumCADTotalYear) * 100, 2, MidpointRounding.ToEven),
                                    AUD = sumAUDTotalYear == 0 ? 0 : Math.Round((item.AUD / sumAUDTotalYear) * 100, 2, MidpointRounding.ToEven),
                                    GBP = sumGBPTotalYear == 0 ? 0 : Math.Round((item.GBP / sumGBPTotalYear) * 100, 2, MidpointRounding.ToEven),
                                    Year = item.Year
                                };

                                resultConvert.Add(itemDetailtPercent);
                            }
                            else
                            {
                                ReportDetailtForTotalMoneyType itemDetailtPercent = new ReportDetailtForTotalMoneyType()
                                {
                                    MarketCode = item.MarketCode,
                                    MarketName = item.MarketName,
                                    PartnerCode = item.PartnerCode,
                                    PartnerName = item.PartnerName,
                                    VND = sumVNDTotalLastYear == 0 ? 0 : Math.Round((item.VND / sumVNDTotalLastYear) * 100, 2, MidpointRounding.ToEven),
                                    USD = sumUSDTotalLastYear == 0 ? 0 : Math.Round((item.USD / sumUSDTotalLastYear) * 100, 2, MidpointRounding.ToEven),
                                    EUR = sumEURTotalLastYear == 0 ? 0 : Math.Round((item.EUR / sumEURTotalLastYear) * 100, 2, MidpointRounding.ToEven),
                                    CAD = sumCADTotalLastYear == 0 ? 0 : Math.Round((item.CAD / sumCADTotalLastYear) * 100, 2, MidpointRounding.ToEven),
                                    AUD = sumAUDTotalLastYear == 0 ? 0 : Math.Round((item.AUD / sumAUDTotalLastYear) * 100, 2, MidpointRounding.ToEven),
                                    GBP = sumGBPTotalLastYear == 0 ? 0 : Math.Round((item.GBP / sumGBPTotalLastYear) * 100, 2, MidpointRounding.ToEven),
                                    Year = item.Year
                                };

                                resultConvert.Add(itemDetailtPercent);
                            }
                        }
                    }
                    else
                    {
                        List<string> listPartner = new List<string>();
                        List<string> listMarket = new List<string>();

                        foreach (ReportDetailtForTotalMoneyType item in resultClone)
                        {
                            if (!listMarket.Contains(item.MarketName))
                            {
                                listMarket.Add(item.MarketName);
                            }
                        }

                        foreach (ReportDetailtForTotalMoneyType item in resultClone)
                        {
                            ReportDetailtForTotalMoneyType dataItemYear = resultClone.Find(x => x.PartnerCode == item.PartnerCode && x.Year == toYear.ToString());
                            ReportDetailtForTotalMoneyType dataItemLastYear = resultClone.Find(x => x.PartnerCode == item.PartnerCode && x.Year == (toYear - 1).ToString());

                            if (dataItemYear == null)
                            {
                                dataItemYear = new ReportDetailtForTotalMoneyType()
                                {
                                    PartnerCode = item.PartnerCode,
                                    PartnerName = item.PartnerName,
                                    MarketCode = item.MarketCode,
                                    MarketName = item.MarketName,
                                    Year = toYear.ToString()
                                };

                                result.Add(dataItemYear);
                            }

                            if (dataItemLastYear == null)
                            {
                                dataItemLastYear = new ReportDetailtForTotalMoneyType()
                                {
                                    PartnerCode = item.PartnerCode,
                                    PartnerName = item.PartnerName,
                                    MarketCode = item.MarketCode,
                                    MarketName = item.MarketName,
                                    Year = (toYear - 1).ToString()
                                };

                                result.Add(dataItemLastYear);
                            }
                        }

                        List<ReportDetailtForTotalMoneyType> listDataYear = result.Where(x => x.Year == toYear.ToString()).ToList();
                        List<ReportDetailtForTotalMoneyType> listDataLastYear = result.Where(x => x.Year == (toYear - 1).ToString()).ToList();

                        // Year
                        double sumVNDTotalYear = listDataYear.Sum(x => x.VND);
                        double sumUSDTotalYear = listDataYear.Sum(x => x.USD);
                        double sumEURTotalYear = listDataYear.Sum(x => x.EUR);
                        double sumCADTotalYear = listDataYear.Sum(x => x.CAD);
                        double sumAUDTotalYear = listDataYear.Sum(x => x.AUD);
                        double sumGBPTotalYear = listDataYear.Sum(x => x.GBP);


                        // Last Year
                        double sumVNDTotalLastYear = listDataLastYear.Sum(x => x.VND);
                        double sumUSDTotalLastYear = listDataLastYear.Sum(x => x.USD);
                        double sumEURTotalLastYear = listDataLastYear.Sum(x => x.EUR);
                        double sumCADTotalLastYear = listDataLastYear.Sum(x => x.CAD);
                        double sumAUDTotalLastYear = listDataLastYear.Sum(x => x.AUD);
                        double sumGBPTotalLastYear = listDataLastYear.Sum(x => x.GBP);



                        foreach (string item in listMarket)
                        {
                            listDataYear = result.Where(x => x.Year == toYear.ToString() && x.MarketName == item).ToList();
                            listDataLastYear = result.Where(x => x.Year == (toYear - 1).ToString() && x.MarketName == item).ToList();

                            // Year
                            double sumVNDYear = listDataYear.Sum(x => x.VND);
                            double sumUSDYear = listDataYear.Sum(x => x.USD);
                            double sumEURYear = listDataYear.Sum(x => x.EUR);
                            double sumCADYear = listDataYear.Sum(x => x.CAD);
                            double sumAUDYear = listDataYear.Sum(x => x.AUD);
                            double sumGBPYear = listDataYear.Sum(x => x.GBP);


                            // Last Year
                            double sumVNDLastYear = listDataLastYear.Sum(x => x.VND);
                            double sumUSDLastYear = listDataLastYear.Sum(x => x.USD);
                            double sumEURLastYear = listDataLastYear.Sum(x => x.EUR);
                            double sumCADLastYear = listDataLastYear.Sum(x => x.CAD);
                            double sumAUDLastYear = listDataLastYear.Sum(x => x.AUD);
                            double sumGBPLastYear = listDataLastYear.Sum(x => x.GBP);

                            // Year
                            resultConvert.Add(
                                new ReportDetailtForTotalMoneyType()
                                {
                                    MarketName = item,
                                    VND = sumVNDTotalYear == 0 ? 0 : Math.Round((sumVNDYear / sumVNDTotalYear) * 100, 2, MidpointRounding.ToEven),
                                    USD = sumUSDTotalYear == 0 ? 0 : Math.Round((sumUSDYear / sumUSDTotalYear) * 100, 2, MidpointRounding.ToEven),
                                    EUR = sumEURTotalYear == 0 ? 0 : Math.Round((sumEURYear / sumEURTotalYear) * 100, 2, MidpointRounding.ToEven),
                                    CAD = sumCADTotalYear == 0 ? 0 : Math.Round((sumCADYear / sumCADTotalYear) * 100, 2, MidpointRounding.ToEven),
                                    AUD = sumAUDTotalYear == 0 ? 0 : Math.Round((sumAUDYear / sumAUDTotalYear) * 100, 2, MidpointRounding.ToEven),
                                    GBP = sumGBPTotalYear == 0 ? 0 : Math.Round((sumGBPYear / sumGBPTotalYear) * 100, 2, MidpointRounding.ToEven),
                                    Year = toYear.ToString()
                                }
                            );

                            // Last Year
                            resultConvert.Add(
                                new ReportDetailtForTotalMoneyType()
                                {
                                    MarketName = item,
                                    VND = sumVNDTotalLastYear == 0 ? 0 : Math.Round((sumVNDLastYear / sumVNDTotalLastYear) * 100, 2, MidpointRounding.ToEven),
                                    USD = sumUSDTotalLastYear == 0 ? 0 : Math.Round((sumUSDLastYear / sumUSDTotalLastYear) * 100, 2, MidpointRounding.ToEven),
                                    EUR = sumEURTotalLastYear == 0 ? 0 : Math.Round((sumEURLastYear / sumEURTotalLastYear) * 100, 2, MidpointRounding.ToEven),
                                    CAD = sumCADTotalLastYear == 0 ? 0 : Math.Round((sumCADLastYear / sumCADTotalLastYear) * 100, 2, MidpointRounding.ToEven),
                                    AUD = sumAUDTotalLastYear == 0 ? 0 : Math.Round((sumAUDLastYear / sumAUDTotalLastYear) * 100, 2, MidpointRounding.ToEven),
                                    GBP = sumGBPTotalLastYear == 0 ? 0 : Math.Round((sumGBPLastYear / sumGBPTotalLastYear) * 100, 2, MidpointRounding.ToEven),
                                    Year = (toYear - 1).ToString()
                                }
                            );
                        }
                    }
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
        public List<ReportDetailtForTotalMoneyType> ReportDetailtMTCompareMonthForAll(int toYear, int toMonth, string reportTypeID, string marketID)
        {
            try
            {
                HSReportDAL dal = new HSReportDAL();
                List<ReportDetailtForTotalMoneyType> result = dal.ReportDetailtMTCompareMonthForAll(toYear, toMonth, reportTypeID, marketID);
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
        public List<ReportDetailtForTotalMoneyType> ColumnChartStackDetailtMTCompareMonthForAllPercent(int toYear, int toMonth, string reportTypeID, string marketID)
        {
            try
            {
                HSReportDAL dal = new HSReportDAL();
                List<ReportDetailtForTotalMoneyType> result = dal.ReportDetailtMTCompareMonthForAll(toYear, toMonth, reportTypeID, marketID);
                List<ReportDetailtForTotalMoneyType> resultConvert = new List<ReportDetailtForTotalMoneyType>();

                foreach (ReportDetailtForTotalMoneyType item in result)
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
        public List<ReportDetailtForTotalMoneyType> ReportDetailtMTCompareMonthForOne(int year, int month, string reportTypeID, string marketID)
        {
            try
            {
                HSReportDAL dal = new HSReportDAL();
                List<ReportDetailtForTotalMoneyType> result = dal.ReportDetailtMTCompareMonthForOne(year, month, reportTypeID, marketID);
                List<ReportDetailtForTotalMoneyType> resultConvert = new List<ReportDetailtForTotalMoneyType>(result);

                List<string> listTemp = new List<string>();

                foreach (ReportDetailtForTotalMoneyType item in resultConvert)
                {
                    // Cùng kì
                    ReportDetailtForTotalMoneyType dataItemLastYear = result.Find(x => x.PartnerCode == item.PartnerCode && x.Month == month.ToString() && x.Year == (year - 1).ToString());
                    ReportDetailtForTotalMoneyType dataItemYear = result.Find(x => x.PartnerCode == item.PartnerCode && x.Month == month.ToString() && x.Year == year.ToString());
                    ReportDetailtForTotalMoneyType dataItemLastMonth = result.Find(x => x.PartnerCode == item.PartnerCode && x.Month == (month - 1).ToString() && x.Year == year.ToString());
                    // Trường hợp tháng 1
                    if (month == 1)
                    {
                        dataItemLastMonth = result.Find(x => x.PartnerCode == item.PartnerCode && x.Month == "12" && x.Year == (year - 1).ToString());
                    }
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
                            if (month == 1)
                            {
                                dataItemLastMonth.Year = (year - 1).ToString();
                                dataItemLastMonth.Month = "12";
                            }
                            else
                            {
                                dataItemLastMonth.Month = (month - 1).ToString();
                                dataItemLastMonth.Year = year.ToString();
                            }
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
                HSReportDAL dal = new HSReportDAL();
                List<ReportDetailtForTotalMoneyType> result = dal.ReportDetailtMTCompareMonthForOne(year, month, reportTypeID, marketID);
                List<ReportDetailtForTotalMoneyType> resultConvert = new List<ReportDetailtForTotalMoneyType>(result);
                List<ReportDetailtForTotalMoneyType> resultConvertPercent = new List<ReportDetailtForTotalMoneyType>();

                List<string> listTemp = new List<string>();

                List<ReportDetailtForTotalMoneyType> listDataItemLastYear = result.Where(x => x.Month == month.ToString() && x.Year == (year - 1).ToString()).ToList();
                List<ReportDetailtForTotalMoneyType> listDataItemYear = result.Where(x => x.Month == month.ToString() && x.Year == year.ToString()).ToList();
                List<ReportDetailtForTotalMoneyType> listDataItemLastMonth = result.Where(x => x.Month == (month - 1).ToString() && x.Year == year.ToString()).ToList();
                // Trường hợp tháng 1
                if (month == 1)
                {
                    listDataItemLastMonth = result.Where(x => x.Month == "12" && x.Year == (year - 1).ToString()).ToList();
                }
                // Cùng kì
                double sumVNDTotalLastYear = listDataItemLastYear.Sum(x => x.VND);
                double sumUSDTotalLastYear = listDataItemLastYear.Sum(x => x.USD);
                double sumEURTotalLastYear = listDataItemLastYear.Sum(x => x.EUR);
                double sumCADTotalLastYear = listDataItemLastYear.Sum(x => x.CAD);
                double sumAUDTotalLastYear = listDataItemLastYear.Sum(x => x.AUD);
                double sumGBPTotalLastYear = listDataItemLastYear.Sum(x => x.GBP);

                // Tháng hiện tại
                double sumVNDTotalYear = listDataItemYear.Sum(x => x.VND);
                double sumUSDTotalYear = listDataItemYear.Sum(x => x.USD);
                double sumEURTotalYear = listDataItemYear.Sum(x => x.EUR);
                double sumCADTotalYear = listDataItemYear.Sum(x => x.CAD);
                double sumAUDTotalYear = listDataItemYear.Sum(x => x.AUD);
                double sumGBPTotalYear = listDataItemYear.Sum(x => x.GBP);

                // Tháng trước
                double sumVNDTotalLastMonth = listDataItemLastMonth.Sum(x => x.VND);
                double sumUSDTotalLastMonth = listDataItemLastMonth.Sum(x => x.USD);
                double sumEURTotalLastMonth = listDataItemLastMonth.Sum(x => x.EUR);
                double sumCADTotalLastMonth = listDataItemLastMonth.Sum(x => x.CAD);
                double sumAUDTotalLastMonth = listDataItemLastMonth.Sum(x => x.AUD);
                double sumGBPTotalLastMonth = listDataItemLastMonth.Sum(x => x.GBP);

                foreach (ReportDetailtForTotalMoneyType item in resultConvert)
                {
                    // Cùng kì
                    ReportDetailtForTotalMoneyType dataItemLastYear = result.Find(x => x.PartnerCode == item.PartnerCode && x.Month == month.ToString() && x.Year == (year - 1).ToString());
                    ReportDetailtForTotalMoneyType dataItemYear = result.Find(x => x.PartnerCode == item.PartnerCode && x.Month == month.ToString() && x.Year == year.ToString());
                    ReportDetailtForTotalMoneyType dataItemLastMonth = result.Find(x => x.PartnerCode == item.PartnerCode && x.Month == (month - 1).ToString() && x.Year == year.ToString());
                    // Trường hợp tháng 1
                    if (month == 1)
                    {
                        dataItemLastMonth = result.Find(x => x.PartnerCode == item.PartnerCode && x.Month == "12" && x.Year == (year - 1).ToString());
                    }
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
                            if (month == 1)
                            {
                                dataItemLastMonth.Year = (year - 1).ToString();
                                dataItemLastMonth.Month = "12";
                            }
                            else
                            {
                                dataItemLastMonth.Year = (month - 1).ToString();
                                dataItemLastMonth.Month = year.ToString();
                            }
                            result.Add(dataItemLastMonth);
                        }

                        // Add partnerCode để kiểm tra
                        listTemp.Add(item.PartnerCode);
                    }
                }

                foreach (ReportDetailtForTotalMoneyType item in result)
                {
                    // Tháng hiện tại
                    if (item.Month == month.ToString() && item.Year == year.ToString())
                    {
                        resultConvertPercent.Add(
                             new ReportDetailtForTotalMoneyType()
                             {
                                 PartnerCode = item.PartnerCode,
                                 PartnerName = item.PartnerName,
                                 MarketCode = item.MarketCode,
                                 MarketName = item.MarketName,
                                 VND = sumVNDTotalYear == 0 ? 0 : Math.Round((item.VND / sumVNDTotalYear) * 100, 2, MidpointRounding.ToEven),
                                 USD = sumUSDTotalYear == 0 ? 0 : Math.Round((item.USD / sumUSDTotalYear) * 100, 2, MidpointRounding.ToEven),
                                 EUR = sumEURTotalYear == 0 ? 0 : Math.Round((item.EUR / sumEURTotalYear) * 100, 2, MidpointRounding.ToEven),
                                 CAD = sumCADTotalYear == 0 ? 0 : Math.Round((item.CAD / sumCADTotalYear) * 100, 2, MidpointRounding.ToEven),
                                 AUD = sumAUDTotalYear == 0 ? 0 : Math.Round((item.AUD / sumAUDTotalYear) * 100, 2, MidpointRounding.ToEven),
                                 GBP = sumGBPTotalYear == 0 ? 0 : Math.Round((item.GBP / sumGBPTotalYear) * 100, 2, MidpointRounding.ToEven),
                                 Month = item.Month,
                                 Year = item.Year
                             }
                        );
                    }

                    // Tháng trước
                    if (month == 1)
                    {
                        if (item.Month == "12" && item.Year == (year - 1).ToString())
                        {
                            resultConvertPercent.Add(
                                 new ReportDetailtForTotalMoneyType()
                                 {
                                     PartnerCode = item.PartnerCode,
                                     PartnerName = item.PartnerName,
                                     MarketCode = item.MarketCode,
                                     MarketName = item.MarketName,
                                     VND = sumVNDTotalLastMonth == 0 ? 0 : Math.Round((item.VND / sumVNDTotalLastMonth) * 100, 2, MidpointRounding.ToEven),
                                     USD = sumUSDTotalLastMonth == 0 ? 0 : Math.Round((item.USD / sumUSDTotalLastMonth) * 100, 2, MidpointRounding.ToEven),
                                     EUR = sumEURTotalLastMonth == 0 ? 0 : Math.Round((item.EUR / sumEURTotalLastMonth) * 100, 2, MidpointRounding.ToEven),
                                     CAD = sumCADTotalLastMonth == 0 ? 0 : Math.Round((item.CAD / sumCADTotalLastMonth) * 100, 2, MidpointRounding.ToEven),
                                     AUD = sumAUDTotalLastMonth == 0 ? 0 : Math.Round((item.AUD / sumAUDTotalLastMonth) * 100, 2, MidpointRounding.ToEven),
                                     GBP = sumGBPTotalLastMonth == 0 ? 0 : Math.Round((item.GBP / sumGBPTotalLastMonth) * 100, 2, MidpointRounding.ToEven),
                                     Month = "12",
                                     Year = (year - 1).ToString()
                                 }
                            );
                        }
                    }
                    else
                    {
                        if (item.Month == (month - 1).ToString() && item.Year == year.ToString())
                        {
                            resultConvertPercent.Add(
                                 new ReportDetailtForTotalMoneyType()
                                 {
                                     PartnerCode = item.PartnerCode,
                                     PartnerName = item.PartnerName,
                                     MarketCode = item.MarketCode,
                                     MarketName = item.MarketName,
                                     VND = sumVNDTotalLastMonth == 0 ? 0 : Math.Round((item.VND / sumVNDTotalLastMonth) * 100, 2, MidpointRounding.ToEven),
                                     USD = sumUSDTotalLastMonth == 0 ? 0 : Math.Round((item.USD / sumUSDTotalLastMonth) * 100, 2, MidpointRounding.ToEven),
                                     EUR = sumEURTotalLastMonth == 0 ? 0 : Math.Round((item.EUR / sumEURTotalLastMonth) * 100, 2, MidpointRounding.ToEven),
                                     CAD = sumCADTotalLastMonth == 0 ? 0 : Math.Round((item.CAD / sumCADTotalLastMonth) * 100, 2, MidpointRounding.ToEven),
                                     AUD = sumAUDTotalLastMonth == 0 ? 0 : Math.Round((item.AUD / sumAUDTotalLastMonth) * 100, 2, MidpointRounding.ToEven),
                                     GBP = sumGBPTotalLastMonth == 0 ? 0 : Math.Round((item.GBP / sumGBPTotalLastMonth) * 100, 2, MidpointRounding.ToEven),
                                     Month = (month - 1).ToString(),
                                     Year = year.ToString()
                                 }
                            );
                        }
                    }


                    // Cùng kì năm trước
                    if (item.Month == month.ToString() && item.Year == (year - 1).ToString())
                    {
                        resultConvertPercent.Add(
                             new ReportDetailtForTotalMoneyType()
                             {
                                 PartnerCode = item.PartnerCode,
                                 PartnerName = item.PartnerName,
                                 MarketCode = item.MarketCode,
                                 MarketName = item.MarketName,
                                 VND = sumVNDTotalLastYear == 0 ? 0 : Math.Round((item.VND / sumVNDTotalLastYear) * 100, 2, MidpointRounding.ToEven),
                                 USD = sumUSDTotalLastYear == 0 ? 0 : Math.Round((item.USD / sumUSDTotalLastYear) * 100, 2, MidpointRounding.ToEven),
                                 EUR = sumEURTotalLastYear == 0 ? 0 : Math.Round((item.EUR / sumEURTotalLastYear) * 100, 2, MidpointRounding.ToEven),
                                 CAD = sumCADTotalLastYear == 0 ? 0 : Math.Round((item.CAD / sumCADTotalLastYear) * 100, 2, MidpointRounding.ToEven),
                                 AUD = sumAUDTotalLastYear == 0 ? 0 : Math.Round((item.AUD / sumAUDTotalLastYear) * 100, 2, MidpointRounding.ToEven),
                                 GBP = sumGBPTotalLastYear == 0 ? 0 : Math.Round((item.GBP / sumGBPTotalLastYear) * 100, 2, MidpointRounding.ToEven),
                                 Month = item.Month,
                                 Year = item.Year
                             }
                        );
                    }
                }

                return resultConvertPercent;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }
        #endregion
    }
}

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
    }
}

// #################################################################
// # Copyright (C) 2010-2011, ASoft JSC.  All Rights Reserved.
// #
// # History：
// #	Date Time		Updated			Content
// #	05/10/2020		Truong Lam		Create New
// ##################################################################

using DongA.API.DataAccess;
using DongA.Core;
using DongA.Entities;
using System;
using System.Collections.Generic;
using static DongA.Core.DongAException;
using System.Linq;

namespace DongA.API.Bussiness
{
    public class HSReportBL : DongABaseDAL
    {
        /// <summary>
        /// List Report từ ngày đến ngày
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<Report> SearchReportDay(DateTime fromDate, DateTime toDate, int reportTypeID)
        {
            try
            {
                HSReportDAL dal = new HSReportDAL();
                List<Report> result = dal.SearchReportDay(fromDate, toDate, reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// List Report từ ngày đến ngày
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<Report> SearchReportMonth(DateTime fromDate, DateTime toDate, int reportTypeID)
        {
            try
            {
                HSReportDAL dal = new HSReportDAL();
                // get first day in fromMonth
                DateTime fromDateRecent = new DateTime(fromDate.Year, fromDate.Month, 1);

                // get last day in toMonth
                int lastDayInToDate = DateTime.DaysInMonth(toDate.Year, toDate.Month);
                DateTime toDateRecent = new DateTime(toDate.Year, toDate.Month, lastDayInToDate);

                List<Report> result = dal.SearchReportMonth(fromDateRecent, toDateRecent, reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// List Report từ ngày đến ngày
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<Report> SearchReportYear(DateTime fromDate, DateTime toDate, int reportTypeID)
        {
            try
            {
                HSReportDAL dal = new HSReportDAL();
                // get first year
                DateTime fromDateRecent = new DateTime(fromDate.Year, 1, 1);

                // Ngày cuối năm
                DateTime toDateRecent = new DateTime(toDate.Year, 12, 31);
                List<Report> result = dal.SearchReportYear(fromDateRecent, toDateRecent, reportTypeID);
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
        public List<Report> SearchReportGradationCompare(int year, int typeID, string reportTypeID)
        {
            try
            {
                HSReportDAL dal = new HSReportDAL();
                List<Report> result = dal.SearchReportGradationCompare(year, typeID, reportTypeID);
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
        public List<Report> SearchReportCompareMonth(int toYear, int toMonth, string reportTypeID)
        {
            try
            {
                HSReportDAL dal = new HSReportDAL();
                List<Report> result = dal.SearchReportCompareMonth(toYear, toMonth, reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        #region Get dữ liệu cho loại tiền
        /// <summary>
        /// Danh sách report theo thị trường
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportForTotalMoneyType> SearchReportForDay(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            try
            {
                HSReportDAL dal = new HSReportDAL();
                List<ReportForTotalMoneyType> result = dal.SearchReportForDay(fromDate, toDate, reportTypeID);
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
        public List<ReportForTotalMoneyType> SearchReportForMonth(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            try
            {
                HSReportDAL dal = new HSReportDAL();
                // get first day in fromMonth
                DateTime fromDateRecent = new DateTime(fromDate.Year, fromDate.Month, 1);

                // get last day in toMonth
                int lastDayInToDate = DateTime.DaysInMonth(toDate.Year, toDate.Month);
                DateTime toDateRecent = new DateTime(toDate.Year, toDate.Month, lastDayInToDate);

                List<ReportForTotalMoneyType> result = dal.SearchReportForMonth(fromDateRecent, toDateRecent, reportTypeID);
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
        public List<ReportForTotalMoneyType> SearchReportForYear(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            try
            {
                HSReportDAL dal = new HSReportDAL();
                // get first year
                DateTime fromDateRecent = new DateTime(fromDate.Year, 1, 1);

                // Ngày cuối năm
                DateTime toDateRecent = new DateTime(toDate.Year, 12, 31);

                List<ReportForTotalMoneyType> result = dal.SearchReportForYear(fromDateRecent, toDateRecent, reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// List Report theo giai đoạn - Nguyên tệ
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportForTotalMoneyType> SearchForGradationCompare(int year, int typeID, string reportTypeID)
        {
            try
            {
                HSReportDAL dal = new HSReportDAL();
                List<ReportForTotalMoneyType> result = dal.SearchForGradationCompare(year, typeID, reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// List Report cho so sánh last month - Nguyên tệ
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportForTotalMoneyType> SearchForCompareMonth(int toYear, int toMonth, string reportTypeID)
        {
            try
            {
                HSReportDAL dal = new HSReportDAL();
                List<ReportForTotalMoneyType> result = dal.SearchForCompareMonth(toYear, toMonth, reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }
        #endregion

        #region Get dữ liệu cho thị trường

        /// <summary>
        /// List Report từ ngày đến ngày
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
        /// List Report
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
                // get first day in fromMonth
                DateTime fromDateRecent = new DateTime(fromDate.Year, fromDate.Month, 1);

                // get last day in toMonth
                int lastDayInToDate = DateTime.DaysInMonth(toDate.Year, toDate.Month);
                DateTime toDateRecent = new DateTime(toDate.Year, toDate.Month, lastDayInToDate);

                List<ReportForMaket> result = dal.SearchReportMaketForMonth(fromDateRecent, toDateRecent, reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// List Report từ ngày đến ngày
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
                // get first year
                DateTime fromDateRecent = new DateTime(fromDate.Year, 1, 1);

                // Ngày cuối năm
                DateTime toDateRecent = new DateTime(toDate.Year, 12, 31);
                List<ReportForMaket> result = dal.SearchReportMaketForYear(fromDateRecent, toDateRecent, reportTypeID);
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
        public List<ReportForMaket> SearchReportMaketForGradationCompare(int ToYear, int typeID, string reportTypeID)
        {
            try
            {
                HSReportDAL dal = new HSReportDAL();
                List<ReportForMaket> result = dal.SearchReportMaketForGradationCompare(ToYear, typeID, reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }


        /// <summary>
        /// List Report cho so sánh theo tháng
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportForMaket> SearchReportMaketForCompareMonth(int ToYear, int ToMonth, string reportTypeID)
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

        #endregion

        #region Get dữ liệu cho tổng hồ sơ

        /// <summary>
        /// List Report từ ngày đến ngày
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportForTotalPayment> SearchReportHSTotalHSForDay(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            try
            {
                HSReportDAL dal = new HSReportDAL();
                List<ReportForTotalPayment> result = dal.SearchReportHSTotalHSForDay(fromDate, toDate, reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }


        /// <summary>
        /// List Report từ ngày đến ngày
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportForTotalPayment> SearchReportHSTotalHSForMonth(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            try
            {
                HSReportDAL dal = new HSReportDAL();
                // get first day in fromMonth
                DateTime fromDateRecent = new DateTime(fromDate.Year, fromDate.Month, 1);

                // get last day in toMonth
                int lastDayInToDate = DateTime.DaysInMonth(toDate.Year, toDate.Month);
                DateTime toDateRecent = new DateTime(toDate.Year, toDate.Month, lastDayInToDate);

                List<ReportForTotalPayment> result = dal.SearchReportHSTotalHSForMonth(fromDateRecent, toDateRecent, reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// List Report từ ngày đến ngày
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportForTotalPayment> SearchReportHSTotalHSForYear(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            try
            {
                HSReportDAL dal = new HSReportDAL();
                // get first year
                DateTime fromDateRecent = new DateTime(fromDate.Year, 1, 1);

                // Ngày cuối năm
                DateTime toDateRecent = new DateTime(toDate.Year, 12, 31);

                List<ReportForTotalPayment> result = dal.SearchReportHSTotalHSForYear(fromDateRecent, toDateRecent, reportTypeID);
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
        public List<ReportForTotalPayment> SearchReportHSTotalHSForGradationCompare(int ToYear, int typeID, string reportTypeID)
        {
            try
            {
                HSReportDAL dal = new HSReportDAL();
                List<ReportForTotalPayment> result = dal.SearchReportHSTotalHSForGradationCompare(ToYear, typeID, reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }
        
        /// <summary>
        /// List Report cho so sánh theo tháng
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportForTotalPayment> SearchReportHSTotalHSForCompareMonth(int ToYear, int ToMonth, string reportTypeID)
        {
            try
            {
                HSReportDAL dal = new HSReportDAL();
                List<ReportForTotalPayment> result = dal.SearchReportHSTotalHSForCompareMonth(ToYear, ToMonth, reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        #endregion

        #region get dữ liệu cho Hồ sơ chi tiết loại hình
        /// <summary>
        /// List Report từ ngày đến ngày
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtSTMarket> SearchDataDetailtMarketForDay(DateTime fromDate, DateTime toDate, string reportTypeID, string marketID)
        {
            try
            {
                HSReportDAL dal = new HSReportDAL();
                List<ReportDetailtSTMarket> result = dal.SearchDataReportDetailtDay(fromDate, toDate, reportTypeID, marketID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// List Report từ ngày đến ngày
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtSTMarket> SearchDataDetailtMarketForMonth(DateTime fromDate, DateTime toDate, string reportTypeID, string marketID)
        {
            try
            {
                HSReportDAL dal = new HSReportDAL();
                // get first day in fromMonth
                DateTime fromDateRecent = new DateTime(fromDate.Year, fromDate.Month, 1);

                // get last day in toMonth
                int lastDayInToDate = DateTime.DaysInMonth(toDate.Year, toDate.Month);
                DateTime toDateRecent = new DateTime(toDate.Year, toDate.Month, lastDayInToDate);

                List<ReportDetailtSTMarket> result = dal.SearchDataReportDetailtDay(fromDateRecent, toDateRecent, reportTypeID, marketID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// List Report từ ngày đến ngày
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtSTMarket> SearchDataDetailtMarketForYear(DateTime fromDate, DateTime toDate, string reportTypeID, string marketID)
        {
            try
            {
                HSReportDAL dal = new HSReportDAL();

                // get first year
                DateTime fromDateRecent = new DateTime(fromDate.Year, 1, 1);

                // Ngày cuối năm
                DateTime toDateRecent = new DateTime(toDate.Year, 12, 31);

                List<ReportDetailtSTMarket> result = dal.SearchDataReportDetailtDay(fromDateRecent, toDateRecent, reportTypeID, marketID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// List Report từ ngày đến ngày
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtServiceType> SearchDataReportDetailtForOneMarketForDay(DateTime fromDate, DateTime toDate, string reportTypeID, string marketID)
        {
            try
            {
                HSReportDAL dal = new HSReportDAL();
                List<ReportDetailtServiceType> result = dal.SearchDataReportDetailtForOneMarketForDay(fromDate, toDate, reportTypeID, marketID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// List Report từ ngày đến ngày
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtServiceType> SearchDataReportDetailtForOneMarketForMonth(DateTime fromDate, DateTime toDate, string reportTypeID, string marketID)
        {
            try
            {
                HSReportDAL dal = new HSReportDAL();

                // get first day in fromMonth
                DateTime fromDateRecent = new DateTime(fromDate.Year, fromDate.Month, 1);

                // get last day in toMonth
                int lastDayInToDate = DateTime.DaysInMonth(toDate.Year, toDate.Month);
                DateTime toDateRecent = new DateTime(toDate.Year, toDate.Month, lastDayInToDate);

                List<ReportDetailtServiceType> result = dal.SearchDataReportDetailtForOneMarketForMonth(fromDateRecent, toDateRecent, reportTypeID, marketID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// List Report từ ngày đến ngày
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtServiceType> SearchDataReportDetailtForOneMarketForYear(DateTime fromDate, DateTime toDate, string reportTypeID, string marketID)
        {
            try
            {
                HSReportDAL dal = new HSReportDAL();

                // get first year
                DateTime fromDateRecent = new DateTime(fromDate.Year, 1, 1);

                // Ngày cuối năm
                DateTime toDateRecent = new DateTime(toDate.Year, 12, 31);

                List<ReportDetailtServiceType> result = dal.SearchDataReportDetailtForOneMarketForYear(fromDateRecent, toDateRecent, reportTypeID, marketID);
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
        public List<ReportDetailtServiceType> SearchDataReportDetailtGradationForAll(int toYear, int typeID, string reportTypeID, string marketID)
        {
            try
            {
                HSReportDAL dal = new HSReportDAL();
                List<ReportDetailtServiceType> result = dal.SearchDataReportDetailtGradationForAll(toYear, typeID, reportTypeID, marketID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// List Report cho so sánh giai đoạn từng thị trường theo đối tác
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtServiceType> SearchDataReportDetailtGradationForOne(int ToYear, int typeID, string reportTypeID, string marketID)
        {
            try
            {
                HSReportDAL dal = new HSReportDAL();
                List<ReportDetailtServiceType> result = dal.SearchDataReportDetailtGradationForOne(ToYear, typeID, reportTypeID, marketID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// List Report detailt của tất cả thị trường theo tháng
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/08/2020]
        /// </history>
        public List<ReportDetailtSTMarket> SearchDataReportDetailtCompareMonthForAll(int ToYear, int ToMonth, string reportTypeID, string marketID)
        {
            try
            {
                HSReportDAL dal = new HSReportDAL();
                List<ReportDetailtSTMarket> result = dal.SearchDataReportDetailtCompareMonthForAll(ToYear, ToMonth, reportTypeID, marketID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// List Report detailt của từng trường theo tháng
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/08/2020]
        /// </history>
        public List<ReportDetailtServiceType> SearchDataReportDetailtCompareMonthForOne(int toYear, int toMonth, string reportTypeID, string marketID)
        {
            try
            {
                HSReportDAL dal = new HSReportDAL();
                List<ReportDetailtServiceType> result = dal.SearchDataReportDetailtCompareMonthForOne(toYear, toMonth, reportTypeID, marketID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        #endregion

        #region Get dữ liệu hồ sơ thị trường theo loại tiền

        /// <summary>
        /// List Report từ ngày đến ngày
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtForTotalMoneyType> SearchDataReportDetailtHSLTForDay(DateTime fromDate, DateTime toDate, string reportTypeID, string marketID)
        {
            try
            {
                HSReportDAL dal = new HSReportDAL();
                List<ReportDetailtForTotalMoneyType> result = dal.SearchDataReportDetailtHSLTForDay(fromDate, toDate, reportTypeID, marketID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// List Report từ ngày đến ngày
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtForTotalMoneyType> SearchDataReportDetailtHSLTForMonth(DateTime fromDate, DateTime toDate, string reportTypeID, string marketID)
        {
            try
            {
                HSReportDAL dal = new HSReportDAL();

                // get first day in fromMonth
                DateTime fromDateRecent = new DateTime(fromDate.Year, fromDate.Month, 1);

                // get last day in toMonth
                int lastDayInToDate = DateTime.DaysInMonth(toDate.Year, toDate.Month);
                DateTime toDateRecent = new DateTime(toDate.Year, toDate.Month, lastDayInToDate);


                List<ReportDetailtForTotalMoneyType> result = dal.SearchDataReportDetailtHSLTForDay(fromDateRecent, toDateRecent, reportTypeID, marketID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// List Report từ ngày đến ngày
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtForTotalMoneyType> SearchDataReportDetailtHSLTForYear(DateTime fromDate, DateTime toDate, string reportTypeID, string marketID)
        {
            try
            {
                HSReportDAL dal = new HSReportDAL();

                // get first year
                DateTime fromDateRecent = new DateTime(fromDate.Year, 1, 1);

                // Ngày cuối năm
                DateTime toDateRecent = new DateTime(toDate.Year, 12, 31);

                List<ReportDetailtForTotalMoneyType> result = dal.SearchDataReportDetailtHSLTForDay(fromDateRecent, toDateRecent, reportTypeID, marketID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }


        /// <summary>
        /// List Report từ ngày đến ngày
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtForTotalMoneyType> SearchDataReportDetailtHSLTForOneMarketForDay(DateTime fromDate, DateTime toDate, string reportTypeID, string marketID)
        {
            try
            {
                HSReportDAL dal = new HSReportDAL();
                List<ReportDetailtForTotalMoneyType> result = dal.SearchDataReportDetailtHSLTForOneMarketForDay(fromDate, toDate, reportTypeID, marketID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// List Report từ ngày đến ngày
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtForTotalMoneyType> SearchDataReportDetailtHSLTForOneMarketForMonth(DateTime fromDate, DateTime toDate, string reportTypeID, string marketID)
        {
            try
            {
                HSReportDAL dal = new HSReportDAL();

                // get first day in fromMonth
                DateTime fromDateRecent = new DateTime(fromDate.Year, fromDate.Month, 1);

                // get last day in toMonth
                int lastDayInToDate = DateTime.DaysInMonth(toDate.Year, toDate.Month);
                DateTime toDateRecent = new DateTime(toDate.Year, toDate.Month, lastDayInToDate);

                List<ReportDetailtForTotalMoneyType> result = dal.SearchDataReportDetailtHSLTForOneMarketForDay(fromDateRecent, toDateRecent, reportTypeID, marketID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// List Report từ ngày đến ngày
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtForTotalMoneyType> SearchDataReportDetailtHSLTForOneMarketForYear(DateTime fromDate, DateTime toDate, string reportTypeID, string marketID)
        {
            try
            {
                HSReportDAL dal = new HSReportDAL();

                // get first year
                DateTime fromDateRecent = new DateTime(fromDate.Year, 1, 1);

                // Ngày cuối năm
                DateTime toDateRecent = new DateTime(toDate.Year, 12, 31);

                List<ReportDetailtForTotalMoneyType> result = dal.SearchDataReportDetailtHSLTForOneMarketForDay(fromDateRecent, toDateRecent, reportTypeID, marketID);
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
        public List<ReportDetailtForTotalMoneyType> SearchDataReportDetailtMTGradationForAll(int ToYear, int typeID, string reportTypeID, string marketID)
        {
            try
            {
                HSReportDAL dal = new HSReportDAL();
                List<ReportDetailtForTotalMoneyType> result = dal.SearchDataReportDetailtMTGradationForAll(ToYear, typeID, reportTypeID, marketID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }


        /// <summary>
        /// List Report cho so sánh giai đoạn từng thị trường theo đối tác
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtForTotalMoneyType> SearchDataReportDetailtMTGradationForOne(int ToYear, int typeID, string reportTypeID, string marketID)
        {
            try
            {
                HSReportDAL dal = new HSReportDAL();
                List<ReportDetailtForTotalMoneyType> result = dal.SearchDataReportDetailtMTGradationForOne(ToYear, typeID, reportTypeID, marketID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }


        /// <summary>
        /// List Report detailt của tất cả thị trường theo tháng
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/08/2020]
        /// </history>
        public List<ReportDetailtForTotalMoneyType> SearchDataReportDetailtMTCompareMonthForAll(int ToYear, int ToMonth, string reportTypeID, string marketID)
        {
            try
            {
                HSReportDAL dal = new HSReportDAL();
                List<ReportDetailtForTotalMoneyType> result = dal.SearchDataReportDetailtMTCompareMonthForAll(ToYear, ToMonth, reportTypeID, marketID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }


        /// <summary>
        /// List Report detailt của từng trường theo tháng
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/08/2020]
        /// </history>
        public List<ReportDetailtForTotalMoneyType> SearchDataReportDetailtMTCompareMonthForOne(int toYear, int toMonth, string reportTypeID, string marketID)
        {
            try
            {
                HSReportDAL dal = new HSReportDAL();
                List<ReportDetailtForTotalMoneyType> result = dal.SearchDataReportDetailtMTCompareMonthForOne(toYear, toMonth, reportTypeID, marketID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        #endregion

        #region Get dữ liệu chi tiết đối tác loại hình

        /// <summary>
        /// List Report từ ngày đến ngày
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtForPartner> SearchDataDetailtForPartnerForAll(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            try
            {
                HSReportDAL dal = new HSReportDAL();
                List<ReportDetailtForPartner> result = dal.SearchDataDetailtForPartnerForAll(fromDate, toDate, reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// List Report từ ngày đến ngày
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtForPartner> SearchDataDetailtForPartnerForAllForMonth(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            try
            {
                HSReportDAL dal = new HSReportDAL();

                // get first day in fromMonth
                DateTime fromDateRecent = new DateTime(fromDate.Year, fromDate.Month, 1);

                // get last day in toMonth
                int lastDayInToDate = DateTime.DaysInMonth(toDate.Year, toDate.Month);
                DateTime toDateRecent = new DateTime(toDate.Year, toDate.Month, lastDayInToDate);

                List<ReportDetailtForPartner> result = dal.SearchDataDetailtForPartnerForAll(fromDateRecent, toDateRecent, reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// List Report từ ngày đến ngày
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtForPartner> SearchDataDetailtForPartnerForAllForYear(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            try
            {
                HSReportDAL dal = new HSReportDAL();

                // get first year
                DateTime fromDateRecent = new DateTime(fromDate.Year, 1, 1);

                // Ngày cuối năm
                DateTime toDateRecent = new DateTime(toDate.Year, 12, 31);

                List<ReportDetailtForPartner> result = dal.SearchDataDetailtForPartnerForAll(fromDateRecent, toDateRecent, reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }


        /// <summary>
        /// List Report từ ngày đến ngày
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtForPartner> SearchDataDetailtForOnePartnerForDay(DateTime fromDate, DateTime toDate, string reportTypeID, string partnerID)
        {
            try
            {
                HSReportDAL dal = new HSReportDAL();
                List<ReportDetailtForPartner> result = dal.SearchDataDetailtForOnePartnerForDay(fromDate, toDate, reportTypeID, partnerID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// List Report từ ngày đến ngày
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtForPartner> SearchDataDetailtForOnePartnerForMonth(DateTime fromDate, DateTime toDate, string reportTypeID, string partnerID)
        {
            try
            {
                HSReportDAL dal = new HSReportDAL();

                // get first day in fromMonth
                DateTime fromDateRecent = new DateTime(fromDate.Year, fromDate.Month, 1);

                // get last day in toMonth
                int lastDayInToDate = DateTime.DaysInMonth(toDate.Year, toDate.Month);
                DateTime toDateRecent = new DateTime(toDate.Year, toDate.Month, lastDayInToDate);

                List<ReportDetailtForPartner> result = dal.SearchDataDetailtForOnePartnerForMonth(fromDateRecent, toDateRecent, reportTypeID, partnerID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }


        /// <summary>
        /// List Report từ ngày đến ngày
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtForPartner> SearchDataDetailtForOnePartnerForYear(DateTime fromDate, DateTime toDate, string reportTypeID, string partnerID)
        {
            try
            {
                HSReportDAL dal = new HSReportDAL();
                // get first year
                DateTime fromDateRecent = new DateTime(fromDate.Year, 1, 1);

                // Ngày cuối năm
                DateTime toDateRecent = new DateTime(toDate.Year, 12, 31);

                List<ReportDetailtForPartner> result = dal.SearchDataDetailtForOnePartnerForYear(fromDateRecent, toDateRecent, reportTypeID, partnerID);
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
        public List<ReportDetailtForPartner> SearchDataReportDetailtGradationForPartnerForAll(int toYear, int typeID, string reportTypeID)
        {
            try
            {
                HSReportDAL dal = new HSReportDAL();
                List<ReportDetailtForPartner> result = dal.SearchDataReportDetailtGradationForPartnerForAll(toYear, typeID, reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }


        /// <summary>
        /// List Report cho so sánh giai đoạn từng thị trường theo đối tác
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtForPartner> SearchDataReportDetailtGradationForPartnerForOne(int ToYear, int typeID, string reportTypeID, string partnerID)
        {
            try
            {
                HSReportDAL dal = new HSReportDAL();
                List<ReportDetailtForPartner> result = dal.SearchDataReportDetailtGradationForPartnerForOne(ToYear, typeID, reportTypeID, partnerID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }


        /// <summary>
        /// List Report detailt của tất cả thị trường theo tháng
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/08/2020]
        /// </history>
        public List<ReportDetailtForPartner> SearchDataReportDetailtCompareMonthForPartnerForAll(int ToYear, int ToMonth, string reportTypeID)
        {
            try
            {
                HSReportDAL dal = new HSReportDAL();
                List<ReportDetailtForPartner> result = dal.SearchDataReportDetailtCompareMonthForPartnerForAll(ToYear, ToMonth, reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// List Report detailt của từng trường theo tháng
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/08/2020]
        /// </history>
        public List<ReportDetailtForPartner> SearchDataReportDetailtCompareMonthForPartnerForOne(int toYear, int toMonth, string reportTypeID, string partnerID)
        {
            try
            {
                HSReportDAL dal = new HSReportDAL();
                List<ReportDetailtForPartner> result = dal.SearchDataReportDetailtCompareMonthForPartnerForOne(toYear, toMonth, reportTypeID, partnerID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }
        #endregion

        #region API cho báo cáo chi tiết theo đối tác loại tiền
        /// <summary>
        /// List Report từ ngày đến ngày
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtForTotalMoneyType> SearchDataDetailtForPartnerLTForAll(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            try
            {
                HSReportDAL dal = new HSReportDAL();
                List<ReportDetailtForTotalMoneyType> result = dal.SearchDataDetailtForPartnerLTForAll(fromDate, toDate, reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// List Report từ ngày đến ngày
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtForTotalMoneyType> SearchDataDetailtForPartnerLTForAllForMonth(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            try
            {
                HSReportDAL dal = new HSReportDAL();
                // get first day in fromMonth
                DateTime fromDateRecent = new DateTime(fromDate.Year, fromDate.Month, 1);

                // get last day in toMonth
                int lastDayInToDate = DateTime.DaysInMonth(toDate.Year, toDate.Month);
                DateTime toDateRecent = new DateTime(toDate.Year, toDate.Month, lastDayInToDate);

                List<ReportDetailtForTotalMoneyType> result = dal.SearchDataDetailtForPartnerLTForAll(fromDateRecent, toDateRecent, reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }


        /// <summary>
        /// List Report từ ngày đến ngày
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtForTotalMoneyType> SearchDataDetailtForPartnerLTForAllForYear(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            try
            {
                HSReportDAL dal = new HSReportDAL();

                // get first year
                DateTime fromDateRecent = new DateTime(fromDate.Year, 1, 1);

                // Ngày cuối năm
                DateTime toDateRecent = new DateTime(toDate.Year, 12, 31);

                List<ReportDetailtForTotalMoneyType> result = dal.SearchDataDetailtForPartnerLTForAll(fromDateRecent, toDateRecent, reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }


        /// <summary>
        /// List Report từ ngày đến ngày
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtForTotalMoneyType> SearchDataDetailtForOnePartnerLTForDay(DateTime fromDate, DateTime toDate, string reportTypeID, string partnerID)
        {
            try
            {
                HSReportDAL dal = new HSReportDAL();
                List<ReportDetailtForTotalMoneyType> result = dal.SearchDataDetailtForOnePartnerLTForDay(fromDate, toDate, reportTypeID, partnerID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }


        /// <summary>
        /// List Report từ ngày đến ngày
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtForTotalMoneyType> SearchDataDetailtForOnePartnerLTForMonth(DateTime fromDate, DateTime toDate, string reportTypeID, string partnerID)
        {
            try
            {
                HSReportDAL dal = new HSReportDAL();
                // get first day in fromMonth
                DateTime fromDateRecent = new DateTime(fromDate.Year, fromDate.Month, 1);

                // get last day in toMonth
                int lastDayInToDate = DateTime.DaysInMonth(toDate.Year, toDate.Month);
                DateTime toDateRecent = new DateTime(toDate.Year, toDate.Month, lastDayInToDate);

                List<ReportDetailtForTotalMoneyType> result = dal.SearchDataDetailtForOnePartnerLTForMonth(fromDateRecent, toDateRecent, reportTypeID, partnerID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }


        /// <summary>
        /// List Report từ ngày đến ngày
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtForTotalMoneyType> SearchDataDetailtForOnePartnerLTForYear(DateTime fromDate, DateTime toDate, string reportTypeID, string partnerID)
        {
            try
            {
                HSReportDAL dal = new HSReportDAL();
                // get first year
                DateTime fromDateRecent = new DateTime(fromDate.Year, 1, 1);

                // Ngày cuối năm
                DateTime toDateRecent = new DateTime(toDate.Year, 12, 31);

                List<ReportDetailtForTotalMoneyType> result = dal.SearchDataDetailtForOnePartnerLTForYear(fromDateRecent, toDateRecent, reportTypeID, partnerID);
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
        public List<ReportDetailtForTotalMoneyType> SearchDataReportDetailtGradationForPartnerLTForAll(int toYear, int typeID, string reportTypeID)
        {
            try
            {
                HSReportDAL dal = new HSReportDAL();
                List<ReportDetailtForTotalMoneyType> result = dal.SearchDataReportDetailtGradationForPartnerLTForAll(toYear, typeID, reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }


        /// <summary>
        /// List Report cho so sánh giai đoạn từng thị trường theo đối tác
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtForTotalMoneyType> SearchDataReportDetailtGradationForPartnerLTForOne(int ToYear, int typeID, string reportTypeID, string partnerID)
        {
            try
            {
                HSReportDAL dal = new HSReportDAL();
                List<ReportDetailtForTotalMoneyType> result = dal.SearchDataReportDetailtGradationForPartnerLTForOne(ToYear, typeID, reportTypeID, partnerID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }


        /// <summary>
        /// List Report detailt của tất cả thị trường theo tháng
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/08/2020]
        /// </history>
        public List<ReportDetailtForTotalMoneyType> SearchDataReportDetailtCompareMonthForPartnerLTForAll(int ToYear, int ToMonth, string reportTypeID)
        {
            try
            {
                HSReportDAL dal = new HSReportDAL();
                List<ReportDetailtForTotalMoneyType> result = dal.SearchDataReportDetailtCompareMonthForPartnerLTForAll(ToYear, ToMonth, reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }


        /// <summary>
        /// List Report detailt của từng trường theo tháng
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/08/2020]
        /// </history>
        public List<ReportDetailtForTotalMoneyType> SearchDataReportDetailtCompareMonthForPartnerLTForOne(int toYear, int toMonth, string reportTypeID, string partnerID)
        {
            try
            {
                HSReportDAL dal = new HSReportDAL();
                List<ReportDetailtForTotalMoneyType> result = dal.SearchDataReportDetailtCompareMonthForPartnerLTForOne(toYear, toMonth, reportTypeID, partnerID);
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

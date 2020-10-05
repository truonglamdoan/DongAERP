// #################################################################
// # Copyright (C) 2010-2011, ASoft JSC.  All Rights Reserved.
// #
// # History：
// #	Date Time		Updated			Content
// #	10/06/2020		Truong Lam		Create New
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
    public class ReportBL : DongABaseDAL
    {
        /// <summary>
        /// List
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<Account> GetListAccount()
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
        public List<Report> DataReportDay(DateTime now, int reportTypeID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<Report> result = dal.DataReportDay(now, reportTypeID);
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
                ReportDAL dal = new ReportDAL();
                List<Report> result = dal.SearchReportDay(fromDate, toDate, reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// List Report theo tháng theo now
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
        /// List Report từ ngày đến ngày
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<Report> SearchReportMonth(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
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
        /// List Report theo năm theo now
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<Report> GetListReportYear(DateTime now, string reportTypeID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<Report> result = dal.GetListReportYear(now, reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// List Report từ năm đến năm
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<Report> SearchReportYear(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<Report> result = dal.SearchReportYear(fromDate, toDate, reportTypeID);
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
        /// List Report cho so sánh theo tháng
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<Report> ListDataCompareMonth(int ToYear, int ToMonth, string reportTypeID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<Report> result = dal.ListDataCompareMonth(ToYear, ToMonth, reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// List Report cho ngày - Nguyên tệ
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportForTotalMoneyType> DataReportForDay(DateTime now, string reportTypeID)
        {
            try
            {
                List<ReportForTotalMoneyType> result = new ReportDAL().DataReportForDay(now, reportTypeID);
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
        public List<ReportForTotalMoneyType> SearchReportForDay(DateTime fromDay, DateTime toDay, string reportTypeID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<ReportForTotalMoneyType> result = dal.SearchReportForDay(fromDay, toDay, reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// List Report cho ngày - USD
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportForTotalMoneyType> DataReportForDayConvert(DateTime now, string reportTypeID)
        {
            try
            {
                List<ReportForTotalMoneyType> result = new ReportDAL().DataReportForDayConvert(now, reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// Danh sách report theo ngày - USD
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportForTotalMoneyType> SearchReportForDayConvert(DateTime fromDay, DateTime toDay, string reportTypeID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<ReportForTotalMoneyType> result = dal.SearchReportForDayConvert(fromDay, toDay, reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// List Report cho tháng - Nguyên tệ
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportForTotalMoneyType> DataReportypeForMonth(DateTime now, string reportTypeID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<ReportForTotalMoneyType> result = dal.DataReportypeForMonth(now, reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// List Report theo tháng - USD
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportForTotalMoneyType> DataReportypeForMonthConvert(DateTime now, string reportTypeID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<ReportForTotalMoneyType> result = dal.DataReportypeForMonthConvert(now, reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// Danh sách report tìm kiếm theo tháng - Nguyên tệ
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportForTotalMoneyType> SearchReportForMonth(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();

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
        /// Danh sách report theo tháng - Nguyên tệ
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportForTotalMoneyType> SearchReportForMonthConvert(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();

                // get first day in fromMonth
                DateTime fromDateRecent = new DateTime(fromDate.Year, fromDate.Month, 1);

                // get last day in toMonth
                int lastDayInToDate = DateTime.DaysInMonth(toDate.Year, toDate.Month);
                DateTime toDateRecent = new DateTime(toDate.Year, toDate.Month, lastDayInToDate);

                List<ReportForTotalMoneyType> result = dal.SearchReportForMonthConvert(fromDateRecent, toDateRecent, reportTypeID);
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
        public List<ReportForTotalMoneyType> DataReportypeForYear(DateTime now, string reportTypeID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<ReportForTotalMoneyType> result = dal.DataReportypeForYear(now, reportTypeID);
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
        public List<ReportForTotalMoneyType> DataReportypeForYearConvert(DateTime now, string reportTypeID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<ReportForTotalMoneyType> result = dal.DataReportypeForYearConvert(now, reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// Danh sách report tìm kiếm theo năm - Nguyên tệ
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportForTotalMoneyType> SearchReportForYear(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
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
        /// Danh sách report tìm kiếm theo năm - USD
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportForTotalMoneyType> SearchReportForYearConvert(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                // get first year
                DateTime fromDateRecent = new DateTime(fromDate.Year, 1, 1);

                // Ngày cuối năm
                DateTime toDateRecent = new DateTime(toDate.Year, 12, 31);
                List<ReportForTotalMoneyType> result = dal.SearchReportForYearConvert(fromDateRecent, toDateRecent, reportTypeID);
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
        public List<ReportForTotalMoneyType> DataReportypeForGradationCompare(int year, int typeID, string reportTypeID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<ReportForTotalMoneyType> result = dal.DataReportypeForGradationCompare(year, typeID, reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// List Report theo giai đoạn - USD
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportForTotalMoneyType> DataReportypeForGradationCompareConvert(int year, int typeID, string reportTypeID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<ReportForTotalMoneyType> result = dal.DataReportypeForGradationCompareConvert(year, typeID, reportTypeID);
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
        public List<ReportForTotalMoneyType> DataReportypeCompareForMonth(int toYear, int toMonth, string reportTypeID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<ReportForTotalMoneyType> result = dal.DataReportypeCompareForMonth(toYear, toMonth, reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// List Report cho so sánh last month - USD
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportForTotalMoneyType> DataReportypeCompareForMonthConvert(int toYear, int toMonth, string reportTypeID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<ReportForTotalMoneyType> result = dal.DataReportypeCompareForMonthConvert(toYear, toMonth, reportTypeID);
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
        public List<ReportForTotalPayment> ReportTPForDay(DateTime now, string reportTypeID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<ReportForTotalPayment> result = dal.ReportTPForDay(now, reportTypeID);
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
        /// List Report theo tháng theo now
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportForTotalPayment> GetListReportTPForMonth(DateTime now, string reportTypeID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<ReportForTotalPayment> result = dal.GetListReportTPForMonth(now, reportTypeID);
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
        public List<ReportForTotalPayment> SearchReportTPForMonth(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();

                // get first day in fromMonth
                DateTime fromDateRecent = new DateTime(fromDate.Year, fromDate.Month, 1);

                // get last day in toMonth
                int lastDayInToDate = DateTime.DaysInMonth(toDate.Year, toDate.Month);
                DateTime toDateRecent = new DateTime(toDate.Year, toDate.Month, lastDayInToDate);

                List<ReportForTotalPayment> result = dal.SearchReportTPForMonth(fromDateRecent, toDateRecent, reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// List Report theo năm theo now
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportForTotalPayment> GetListReportTPForYear(DateTime now, string reportTypeID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<ReportForTotalPayment> result = dal.GetListReportTPForYear(now, reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// List Report từ năm đến năm
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
                // get first year
                DateTime fromDateRecent = new DateTime(fromDate.Year, 1, 1);

                // Ngày cuối năm
                DateTime toDateRecent = new DateTime(toDate.Year, 12, 31);
                List<ReportForTotalPayment> result = dal.SearchReportTPForYear(fromDateRecent, toDateRecent, reportTypeID);
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
        public List<ReportForTotalPayment> ListDataTPForGradationCompare(int ToYear, int typeID, string reportTypeID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<ReportForTotalPayment> result = dal.ListDataTPForGradationCompare(ToYear, typeID, reportTypeID);
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
        public List<ReportForTotalPayment> ListDataTPForCompareMonth(int ToYear, int ToMonth, string reportTypeID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<ReportForTotalPayment> result = dal.ListDataTPForCompareMonth(ToYear, ToMonth, reportTypeID);
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
        public List<ReportForMaket> ReportMaketForDay(DateTime now, string reportTypeID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<ReportForMaket> result = dal.ReportMaketForDay(now, reportTypeID);
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
        /// List Report theo tháng theo now
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportForMaket> GetListReportMarketForMonth(DateTime now, string reportTypeID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<ReportForMaket> result = dal.GetListReportMarketForMonth(now, reportTypeID);
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
        public List<ReportForMaket> SearchReportMaketForMonth(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();

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
        /// List Report theo năm theo now
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportForMaket> GetListReportMaketForYear(DateTime now, string reportTypeID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();

                List<ReportForMaket> result = dal.GetListReportMaketForYear(now, reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// List Report từ năm đến năm
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
        public List<ReportForMaket> ListDataMaketForGradationCompare(int ToYear, int typeID, string reportTypeID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<ReportForMaket> result = dal.ListDataMaketForGradationCompare(ToYear, typeID, reportTypeID);
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
        public List<ReportForMaket> ListDataMaketForCompareMonth(int ToYear, int ToMonth, string reportTypeID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<ReportForMaket> result = dal.ListDataMaketForCompareMonth(ToYear, ToMonth, reportTypeID);
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
        /// List Report deatailt cho ngày
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/08/2020]
        /// </history>
        public List<ReportDetailtServiceType> DataReportDetailtDay(DateTime now, int reportTypeID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<ReportDetailtServiceType> result = dal.DataReportDetailtDay(now, reportTypeID);
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
        public List<ReportDetailtServiceType> SearchReportDetailtDay(DateTime fromDate, DateTime toDate, int reportTypeID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<ReportDetailtServiceType> result = dal.SearchReportDetailtDay(fromDate, toDate, reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// List Report theo tháng theo now
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtServiceType> GetListReportDetailtMonth(DateTime now, string reportTypeID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<ReportDetailtServiceType> result = dal.GetListReportDetailtMonth(now, reportTypeID);
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
        public List<ReportDetailtServiceType> SearchReportDetailtMonth(DateTime fromDate, DateTime toDate, string partnerCode, string reportTypeID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                // get first day in fromMonth
                DateTime fromDateRecent = new DateTime(fromDate.Year, fromDate.Month, 1);

                // get last day in toMonth
                int lastDayInToDate = DateTime.DaysInMonth(toDate.Year, toDate.Month);
                DateTime toDateRecent = new DateTime(toDate.Year, toDate.Month, lastDayInToDate);

                List<ReportDetailtServiceType> result = dal.SearchReportDetailtMonth(fromDateRecent, toDateRecent, partnerCode, reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// List Report theo năm theo now
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtServiceType> GetListReportDetailtYear(DateTime now, string reportTypeID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<ReportDetailtServiceType> result = dal.GetListReportDetailtYear(now, reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// List Report từ năm đến năm
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtServiceType> SearchReportDetailtYear(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<ReportDetailtServiceType> result = dal.SearchReportDetailtYear(fromDate, toDate, reportTypeID);
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
        public List<ReportDetailtServiceType> ListDataDetailtGradationCompare(int ToYear, int typeID, string reportTypeID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<ReportDetailtServiceType> result = dal.ListDataDetailtGradationCompare(ToYear, typeID, reportTypeID);
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
        public List<ReportDetailtServiceType> ListDataDetailtCompareMonth(int ToYear, int ToMonth, string reportTypeID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<ReportDetailtServiceType> result = dal.ListDataDetailtCompareMonth(ToYear, ToMonth, reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// List các đối tác
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/08/2020]
        /// </history>
        public List<Partner> ListPartNer()
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<Partner> result = dal.ListPartNer();
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }

        }

        /// <summary>
        /// List các đối tác
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/08/2020]
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
        /// List các đối tác
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/08/2020]
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
        /// List Report deatailt cho ngày
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/08/2020]
        /// </history>
        public List<ReportDetailtSTMarket> DataReportDetailtMarKetForDay(DateTime now, int reportTypeID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<ReportDetailtSTMarket> result = dal.DataReportDetailtMarKetForDay(now, reportTypeID);
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
        public List<ReportDetailtSTMarket> SearchDataDetailtMarketForDay(DateTime fromDate, DateTime toDate, string reportTypeID, string marketID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<ReportDetailtSTMarket> result = dal.SearchDataReportDetailtDay(fromDate, toDate, reportTypeID, marketID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// List Report từ tháng đến tháng
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtSTMarket> SearchDataReportDetailtMonth(DateTime fromDate, DateTime toDate, string reportTypeID, string marketID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                
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
        /// List Report từ năm đến năm
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtSTMarket> SearchDataReportDetailtYear(DateTime fromDate, DateTime toDate, string reportTypeID, string marketID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();

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
        /// List Report deatailt cho ngày
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/08/2020]
        /// </history>
        public List<ReportDetailtServiceType> DataReportDetailtForOneMarket(DateTime now, int reportTypeID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<ReportDetailtServiceType> result = dal.DataReportDetailtForOneMarket(now, reportTypeID);
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
        public List<ReportDetailtServiceType> SearchDataReportDetailtForOneMarket(DateTime fromDate, DateTime toDate, string reportTypeID, string marketID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<ReportDetailtServiceType> result = dal.SearchDataReportDetailtForOneMarket(fromDate, toDate, reportTypeID, marketID);
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
                ReportDAL dal = new ReportDAL();

                // get first day in fromMonth
                DateTime fromDateRecent = new DateTime(fromDate.Year, fromDate.Month, 1);

                // get last day in toMonth
                int lastDayInToDate = DateTime.DaysInMonth(toDate.Year, toDate.Month);
                DateTime toDateRecent = new DateTime(toDate.Year, toDate.Month, lastDayInToDate);

                List<ReportDetailtServiceType> result = dal.SearchDataReportDetailtForOneMarket(fromDateRecent, toDateRecent, reportTypeID, marketID);

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
                ReportDAL dal = new ReportDAL();
                // get first year
                DateTime fromDateRecent = new DateTime(fromDate.Year, 1, 1);

                // Ngày cuối năm
                DateTime toDateRecent = new DateTime(toDate.Year, 12, 31);

                List<ReportDetailtServiceType> result = dal.SearchDataReportDetailtForOneMarket(fromDateRecent, toDateRecent, reportTypeID, marketID);

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
                ReportDAL dal = new ReportDAL();
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
                ReportDAL dal = new ReportDAL();
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
                ReportDAL dal = new ReportDAL();
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
                ReportDAL dal = new ReportDAL();
                List<ReportDetailtServiceType> result = dal.SearchDataReportDetailtCompareMonthForOne(toYear, toMonth, reportTypeID, marketID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        #region Báo cáo chi tiết thị trường theo loại tiền chi trả
        /// <summary>
        /// List Report detailt cho thị trường theo loại tiền chi trả
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/08/2020]
        /// </history>
        public List<ReportDetailtForTotalMoneyType> SearchDataReportDetailtMTForAll(DateTime fromDate, DateTime toDate, string reportTypeID, string marketID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<ReportDetailtForTotalMoneyType> result = dal.SearchDataReportDetailtMTForAll(fromDate, toDate, reportTypeID, marketID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// List Report detailt cho thị trường theo loại tiền chi trả Quy USD
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/08/2020]
        /// </history>
        public List<ReportDetailtForTotalMoneyType> SearchDataReportDetailtMTForAllConvert(DateTime fromDate, DateTime toDate, string reportTypeID, string marketID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<ReportDetailtForTotalMoneyType> result = dal.SearchDataReportDetailtMTForAllConvert(fromDate, toDate, reportTypeID, marketID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// List Report detailt cho thị trường theo loại tiền chi trả
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/08/2020]
        /// </history>
        public List<ReportDetailtForTotalMoneyType> SearchDataReportDetailtMTForOne(DateTime fromDate, DateTime toDate, string reportTypeID, string marketID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<ReportDetailtForTotalMoneyType> result = dal.SearchDataReportDetailtMTForOne(fromDate, toDate, reportTypeID, marketID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// List Report detailt cho thị trường theo loại tiền chi trả Quy USD
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/08/2020]
        /// </history>
        public List<ReportDetailtForTotalMoneyType> SearchDataReportDetailtMTForOneConvert(DateTime fromDate, DateTime toDate, string reportTypeID, string marketID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                // get first day in fromMonth
                DateTime fromDateRecent = new DateTime(fromDate.Year, fromDate.Month, 1);

                // get last day in toMonth
                int lastDayInToDate = DateTime.DaysInMonth(toDate.Year, toDate.Month);
                DateTime toDateRecent = new DateTime(toDate.Year, toDate.Month, lastDayInToDate);

                List<ReportDetailtForTotalMoneyType> result = dal.SearchDataReportDetailtMTForOneConvert(fromDateRecent, toDateRecent, reportTypeID, marketID);
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
        public List<ReportDetailtForTotalMoneyType> SearchDataReportDetailtMTGradationForAll(int ToYear, int typeID, string reportTypeID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<ReportDetailtForTotalMoneyType> result = dal.SearchDataReportDetailtMTGradationForAll(ToYear, typeID, reportTypeID);
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
        public List<ReportDetailtForTotalMoneyType> SearchDataReportDetailtMTGradationForAllConvert(int ToYear, int typeID, string reportTypeID, string marketID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<ReportDetailtForTotalMoneyType> result = dal.SearchDataReportDetailtMTGradationForAllConvert(ToYear, typeID, reportTypeID, marketID);
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
                ReportDAL dal = new ReportDAL();
                List<ReportDetailtForTotalMoneyType> result = dal.SearchDataReportDetailtMTGradationForOne(ToYear, typeID, reportTypeID, marketID);
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
        public List<ReportDetailtForTotalMoneyType> SearchDataReportDetailtMTGradationForOneConvert(int ToYear, int typeID, string reportTypeID, string marketID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<ReportDetailtForTotalMoneyType> result = dal.SearchDataReportDetailtMTGradationForOneConvert(ToYear, typeID, reportTypeID, marketID);
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
        public List<ReportDetailtForTotalMoneyType> SearchDataReportDetailtMTCompareMonthForAll(int ToYear, int ToMonth, string reportTypeID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<ReportDetailtForTotalMoneyType> result = dal.SearchDataReportDetailtMTCompareMonthForAll(ToYear, ToMonth, reportTypeID);
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
        public List<ReportDetailtForTotalMoneyType> SearchDataReportDetailtMTCompareMonthForAllConvert(int ToYear, int ToMonth, string reportTypeID, string marketID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<ReportDetailtForTotalMoneyType> result = dal.SearchDataReportDetailtMTCompareMonthForAllConvert(ToYear, ToMonth, reportTypeID, marketID);
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
                ReportDAL dal = new ReportDAL();
                List<ReportDetailtForTotalMoneyType> result = dal.SearchDataReportDetailtMTCompareMonthForOne(toYear, toMonth, reportTypeID, marketID);
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
        public List<ReportDetailtForTotalMoneyType> SearchDataReportDetailtMTCompareMonthForOneConvert(int toYear, int toMonth, string reportTypeID, string marketID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<ReportDetailtForTotalMoneyType> result = dal.SearchDataReportDetailtMTCompareMonthForOneConvert(toYear, toMonth, reportTypeID, marketID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        #endregion

        #region Báo cáo chi tiết cho Thị trường theo đối tác

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
                ReportDAL dal = new ReportDAL();
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
        public List<ReportDetailtForPartner> SearchDataDetailtForOnePartnerForDay(DateTime fromDate, DateTime toDate, string reportTypeID, string partnerID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<ReportDetailtForPartner> result = dal.SearchDataDetailtForOnePartnerForDay(fromDate, toDate, reportTypeID, partnerID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// List Report từ tháng đến tháng
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtForPartner> SearchDataDetailtForOnePartnerForMonth(DateTime fromDate, DateTime toDate, string reportTypeID, string partnerID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
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
        /// List Report từ năm đến năm
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtForPartner> SearchDataDetailtForOnePartnerForYear(DateTime fromDate, DateTime toDate, string reportTypeID, string partnerID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
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
                ReportDAL dal = new ReportDAL();
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
                ReportDAL dal = new ReportDAL();
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
                ReportDAL dal = new ReportDAL();
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
                ReportDAL dal = new ReportDAL();
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
                ReportDAL dal = new ReportDAL();
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
        public List<ReportDetailtForTotalMoneyType> SearchDataDetailtForPartnerLTForAllConvert(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<ReportDetailtForTotalMoneyType> result = dal.SearchDataDetailtForPartnerLTForAllConvert(fromDate, toDate, reportTypeID);
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
                ReportDAL dal = new ReportDAL();
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
        public List<ReportDetailtForTotalMoneyType> SearchDataDetailtForOnePartnerLTForDayConvert(DateTime fromDate, DateTime toDate, string reportTypeID, string partnerID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<ReportDetailtForTotalMoneyType> result = dal.SearchDataDetailtForOnePartnerLTForDayConvert(fromDate, toDate, reportTypeID, partnerID);
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
                ReportDAL dal = new ReportDAL();
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
        public List<ReportDetailtForTotalMoneyType> SearchDataDetailtForOnePartnerLTForMonthConvert(DateTime fromDate, DateTime toDate, string reportTypeID, string partnerID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                // get first day in fromMonth
                DateTime fromDateRecent = new DateTime(fromDate.Year, fromDate.Month, 1);

                // get last day in toMonth
                int lastDayInToDate = DateTime.DaysInMonth(toDate.Year, toDate.Month);
                DateTime toDateRecent = new DateTime(toDate.Year, toDate.Month, lastDayInToDate);

                List<ReportDetailtForTotalMoneyType> result = dal.SearchDataDetailtForOnePartnerLTForMonthConvert(fromDateRecent, toDateRecent, reportTypeID, partnerID);
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
                ReportDAL dal = new ReportDAL();
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
        /// List Report từ ngày đến ngày
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtForTotalMoneyType> SearchDataDetailtForOnePartnerLTForYearConvert(DateTime fromDate, DateTime toDate, string reportTypeID, string partnerID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                // get first year
                DateTime fromDateRecent = new DateTime(fromDate.Year, 1, 1);

                // Ngày cuối năm
                DateTime toDateRecent = new DateTime(toDate.Year, 12, 31);

                List<ReportDetailtForTotalMoneyType> result = dal.SearchDataDetailtForOnePartnerLTForYearConvert(fromDateRecent, toDateRecent, reportTypeID, partnerID);
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
                ReportDAL dal = new ReportDAL();
                List<ReportDetailtForTotalMoneyType> result = dal.SearchDataReportDetailtGradationForPartnerLTForAll(toYear, typeID, reportTypeID);
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
        public List<ReportDetailtForTotalMoneyType> SearchDataReportDetailtGradationForPartnerLTForAllConvert(int toYear, int typeID, string reportTypeID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<ReportDetailtForTotalMoneyType> result = dal.SearchDataReportDetailtGradationForPartnerLTForAllConvert(toYear, typeID, reportTypeID);
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
                ReportDAL dal = new ReportDAL();
                List<ReportDetailtForTotalMoneyType> result = dal.SearchDataReportDetailtGradationForPartnerLTForOne(ToYear, typeID, reportTypeID, partnerID);
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
        public List<ReportDetailtForTotalMoneyType> SearchDataReportDetailtGradationForPartnerLTForOneConvert(int ToYear, int typeID, string reportTypeID, string partnerID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<ReportDetailtForTotalMoneyType> result = dal.SearchDataReportDetailtGradationForPartnerLTForOneConvert(ToYear, typeID, reportTypeID, partnerID);
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
                ReportDAL dal = new ReportDAL();
                List<ReportDetailtForTotalMoneyType> result = dal.SearchDataReportDetailtCompareMonthForPartnerLTForAll(ToYear, ToMonth, reportTypeID);
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
        public List<ReportDetailtForTotalMoneyType> SearchDataReportDetailtCompareMonthForPartnerLTForAllConvert(int ToYear, int ToMonth, string reportTypeID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<ReportDetailtForTotalMoneyType> result = dal.SearchDataReportDetailtCompareMonthForPartnerLTForAllConvert(ToYear, ToMonth, reportTypeID);
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
                ReportDAL dal = new ReportDAL();
                List<ReportDetailtForTotalMoneyType> result = dal.SearchDataReportDetailtCompareMonthForPartnerLTForOne(toYear, toMonth, reportTypeID, partnerID);
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
        public List<ReportDetailtForTotalMoneyType> SearchDataReportDetailtCompareMonthForPartnerLTForOneConvert(int toYear, int toMonth, string reportTypeID, string partnerID)
        {
            try
            {
                ReportDAL dal = new ReportDAL();
                List<ReportDetailtForTotalMoneyType> result = dal.SearchDataReportDetailtCompareMonthForPartnerLTForOneConvert(toYear, toMonth, reportTypeID, partnerID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }
        #endregion

        #region Báo cáo hồ sơ


        #endregion
    }
}

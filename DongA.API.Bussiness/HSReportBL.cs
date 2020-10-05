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
    }
}

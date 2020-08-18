// #################################################################
// # Copyright (C) 2010-2011, ASoft JSC.  All Rights Reserved.
// #
// # History：
// #	Date Time		Updated			Content
// #	04/08/2020		Truong Lam		Create New
// ##################################################################

using DongA.API.Bussiness;
using DongA.Core;
using DongA.Entities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using static DongA.Core.DongAException;

namespace DongA.API.Controllers
{
    public class ReportDetailtDataForMarketController : ApiController
    {
        /// <summary>
        /// Danh sách thị trường
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/08/2020]
        /// </history>
        [HttpGet]
        public List<Market> ListMarket()
        {
            try
            {
                List<Market> listData = new ReportBL().ListMarket();
                return listData;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        [HttpGet]
        public List<ReportDetailtSTMarket> DataReportDetailtDay(int reportTypeID)
        {
            try
            {
                DateTime now = DateTime.Now;
                List<ReportDetailtSTMarket> listReport = new ReportBL().DataReportDetailtMarKetForDay(now, reportTypeID);
                return listReport;
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
        [HttpGet]
        public List<ReportDetailtSTMarket> SearchDataReportDetailtDay(DateTime fromDate, DateTime toDate, string reportTypeID, string marketID)
        {
            try
            {
                List<ReportDetailtSTMarket> listReport = new ReportBL().SearchDataDetailtMarketForDay(fromDate, toDate, reportTypeID, marketID);
                return listReport;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        [HttpGet]
        public List<ReportDetailtServiceType> DataReportDetailtForOneMarket(int reportTypeID)
        {
            try
            {
                DateTime now = DateTime.Now;
                List<ReportDetailtServiceType> listReport = new ReportBL().DataReportDetailtForOneMarket(now, reportTypeID);
                return listReport;
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
        [HttpGet]
        public List<ReportDetailtServiceType> SearchDataReportDetailtForOneMarket(DateTime fromDate, DateTime toDate, string reportTypeID, string marketID)
        {
            try
            {
                List<ReportDetailtServiceType> listReport = new ReportBL().SearchDataReportDetailtForOneMarket(fromDate, toDate, reportTypeID, marketID);
                return listReport;
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
        [HttpGet]
        public List<ReportDetailtServiceType> SearchDataReportDetailtGradationForAll(int toYear, int typeID, string reportTypeID)
        {
            try
            {
                List<ReportDetailtServiceType> result = new ReportBL().SearchDataReportDetailtGradationForAll(toYear, typeID, reportTypeID);
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
        [HttpGet]
        public List<ReportDetailtServiceType> SearchDataReportDetailtGradationForOne(int toYear, int typeID, string reportTypeID, string marketID)
        {
            try
            {
                List<ReportDetailtServiceType> result = new ReportBL().SearchDataReportDetailtGradationForOne(toYear, typeID, reportTypeID, marketID);
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
        [HttpGet]
        public List<ReportDetailtSTMarket> SearchDataReportDetailtCompareMonthForAll(int toYear, int toMonth, string reportTypeID)
        {
            try
            {
                List<ReportDetailtSTMarket> result = new ReportBL().SearchDataReportDetailtCompareMonthForAll(toYear, toMonth, reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// List Report detailt của từng thị trường theo tháng và đối tác
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/08/2020]
        /// </history>
        [HttpGet]
        public List<ReportDetailtServiceType> SearchDataReportDetailtCompareMonthForOne(int toYear, int toMonth, string reportTypeID, string marketID)
        {
            try
            {
                List<ReportDetailtServiceType> result = new ReportBL().SearchDataReportDetailtCompareMonthForOne(toYear, toMonth, reportTypeID, marketID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }
    }
}

// #################################################################
// # Copyright (C) 2010-2011, ASoft JSC.  All Rights Reserved.
// #
// # History：
// #	Date Time		Updated			Content
// #	09/10/2020		Truong Lam		Create New
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
    public class ReportHSDetailtMarketLTController : ApiController
    {

        /// <summary>
        /// List Report từ ngày đến ngày
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpGet]
        public List<ReportDetailtForTotalMoneyType> SearchDataReportDetailtDay(DateTime fromDate, DateTime toDate, string reportTypeID, string marketID)
        {
            try
            {
                List<ReportDetailtForTotalMoneyType> listReport = new HSReportBL().SearchDataReportDetailtHSLTForDay(fromDate, toDate, reportTypeID, marketID);
                return listReport;
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
        [HttpGet]
        public List<ReportDetailtForTotalMoneyType> SearchDataReportDetailtMonth(DateTime fromDate, DateTime toDate, string reportTypeID, string marketID)
        {
            try
            {
                List<ReportDetailtForTotalMoneyType> listReport = new HSReportBL().SearchDataReportDetailtHSLTForMonth(fromDate, toDate, reportTypeID, marketID);
                return listReport;
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
        [HttpGet]
        public List<ReportDetailtForTotalMoneyType> SearchDataReportDetailtYear(DateTime fromDate, DateTime toDate, string reportTypeID, string marketID)
        {
            try
            {
                List<ReportDetailtForTotalMoneyType> listReport = new HSReportBL().SearchDataReportDetailtHSLTForYear(fromDate, toDate, reportTypeID, marketID);
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
        public List<ReportDetailtForTotalMoneyType> SearchDataReportDetailtForOneMarketForDay(DateTime fromDate, DateTime toDate, string reportTypeID, string marketID)
        {
            try
            {
                List<ReportDetailtForTotalMoneyType> listReport = new HSReportBL().SearchDataReportDetailtHSLTForOneMarketForDay(fromDate, toDate, reportTypeID, marketID);
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
        public List<ReportDetailtForTotalMoneyType> SearchDataReportDetailtForOneMarketForMonth(DateTime fromDate, DateTime toDate, string reportTypeID, string marketID)
        {
            try
            {
                List<ReportDetailtForTotalMoneyType> listReport = new HSReportBL().SearchDataReportDetailtHSLTForOneMarketForMonth(fromDate, toDate, reportTypeID, marketID);
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
        public List<ReportDetailtForTotalMoneyType> SearchDataReportDetailtForOneMarketForYear(DateTime fromDate, DateTime toDate, string reportTypeID, string marketID)
        {
            try
            {
                List<ReportDetailtForTotalMoneyType> listReport = new HSReportBL().SearchDataReportDetailtHSLTForOneMarketForYear(fromDate, toDate, reportTypeID, marketID);
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
        public List<ReportDetailtForTotalMoneyType> SearchDataReportDetailtMTGradationForAll(int toYear, int typeID, string reportTypeID, string marketID)
        {
            try
            {
                List<ReportDetailtForTotalMoneyType> result = new HSReportBL().SearchDataReportDetailtMTGradationForAll(toYear, typeID, reportTypeID, marketID);
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
        public List<ReportDetailtForTotalMoneyType> SearchDataReportDetailtMTGradationForOne(int toYear, int typeID, string reportTypeID, string marketID)
        {
            try
            {
                List<ReportDetailtForTotalMoneyType> result = new HSReportBL().SearchDataReportDetailtMTGradationForOne(toYear, typeID, reportTypeID, marketID);
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
        public List<ReportDetailtForTotalMoneyType> SearchDataReportDetailtMTCompareMonthForAll(int toYear, int toMonth, string reportTypeID, string marketID)
        {
            try
            {
                List<ReportDetailtForTotalMoneyType> result = new HSReportBL().SearchDataReportDetailtMTCompareMonthForAll(toYear, toMonth, reportTypeID, marketID);
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
        public List<ReportDetailtForTotalMoneyType> SearchDataReportDetailtMTCompareMonthForOne(int toYear, int toMonth, string reportTypeID, string marketID)
        {
            try
            {
                List<ReportDetailtForTotalMoneyType> result = new HSReportBL().SearchDataReportDetailtMTCompareMonthForOne(toYear, toMonth, reportTypeID, marketID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }
    }
}

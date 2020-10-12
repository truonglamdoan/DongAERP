// #################################################################
// # Copyright (C) 2010-2011, ASoft JSC.  All Rights Reserved.
// #
// # History：
// #	Date Time		Updated			Content
// #	09/09/2020		Truong Lam		Create New
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
    public class ReportHSDetailtPartnerLTController : ApiController
    {
        /// <summary>
        /// List Report từ ngày đến ngày
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpGet]
        public List<ReportDetailtForTotalMoneyType> SearchDataReportDetailtForPartnerForAll(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            try
            {
                List<ReportDetailtForTotalMoneyType> listReport = new HSReportBL().SearchDataDetailtForPartnerLTForAll(fromDate, toDate, reportTypeID);
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
        public List<ReportDetailtForTotalMoneyType> SearchDataReportDetailtForPartnerForAllForMonth(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            try
            {
                List<ReportDetailtForTotalMoneyType> listReport = new HSReportBL().SearchDataDetailtForPartnerLTForAllForMonth(fromDate, toDate, reportTypeID);
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
        public List<ReportDetailtForTotalMoneyType> SearchDataReportDetailtForPartnerForAllForYear(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            try
            {
                List<ReportDetailtForTotalMoneyType> listReport = new HSReportBL().SearchDataDetailtForPartnerLTForAllForYear(fromDate, toDate, reportTypeID);
                return listReport;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// List Report từ ngày đến ngày cho từng đối tác
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpGet]
        public List<ReportDetailtForTotalMoneyType> SearchDataReportDetailtForOneForDay(DateTime fromDate, DateTime toDate, string reportTypeID, string partnerID)
        {
            try
            {
                List<ReportDetailtForTotalMoneyType> listReport = new HSReportBL().SearchDataDetailtForOnePartnerLTForDay(fromDate, toDate, reportTypeID, partnerID);
                return listReport;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }



        /// <summary>
        /// List Report từ ngày đến ngày cho từng đối tác
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpGet]
        public List<ReportDetailtForTotalMoneyType> SearchDataReportDetailtForOneForMonth(DateTime fromDate, DateTime toDate, string reportTypeID, string partnerID)
        {
            try
            {
                List<ReportDetailtForTotalMoneyType> listReport = new HSReportBL().SearchDataDetailtForOnePartnerLTForMonth(fromDate, toDate, reportTypeID, partnerID);
                return listReport;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }
        
        /// <summary>
        /// List Report từ ngày đến ngày cho từng đối tác
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpGet]
        public List<ReportDetailtForTotalMoneyType> SearchDataReportDetailtForOneForYear(DateTime fromDate, DateTime toDate, string reportTypeID, string partnerID)
        {
            try
            {
                List<ReportDetailtForTotalMoneyType> listReport = new HSReportBL().SearchDataDetailtForOnePartnerLTForYear(fromDate, toDate, reportTypeID, partnerID);
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
        public List<ReportDetailtForTotalMoneyType> SearchDataReportDetailtGradationForAll(int toYear, int typeID, string reportTypeID)
        {
            try
            {
                List<ReportDetailtForTotalMoneyType> result = new HSReportBL().SearchDataReportDetailtGradationForPartnerLTForAll(toYear, typeID, reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }


        /// <summary>
        /// List Report cho so sánh giai đoạn cho từng đối tác
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpGet]
        public List<ReportDetailtForTotalMoneyType> SearchDataReportDetailtGradationForOne(int toYear, int typeID, string reportTypeID, string partnerID)
        {
            try
            {
                List<ReportDetailtForTotalMoneyType> result = new HSReportBL().SearchDataReportDetailtGradationForPartnerLTForOne(toYear, typeID, reportTypeID, partnerID);
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
        public List<ReportDetailtForTotalMoneyType> SearchDataReportDetailtCompareMonthForAll(int toYear, int toMonth, string reportTypeID)
        {
            try
            {
                List<ReportDetailtForTotalMoneyType> result = new HSReportBL().SearchDataReportDetailtCompareMonthForPartnerLTForAll(toYear, toMonth, reportTypeID);
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
        public List<ReportDetailtForTotalMoneyType> SearchDataReportDetailtCompareMonthForOne(int toYear, int toMonth, string reportTypeID, string partnerID)
        {
            try
            {
                List<ReportDetailtForTotalMoneyType> result = new HSReportBL().SearchDataReportDetailtCompareMonthForPartnerLTForOne(toYear, toMonth, reportTypeID, partnerID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

    }
}

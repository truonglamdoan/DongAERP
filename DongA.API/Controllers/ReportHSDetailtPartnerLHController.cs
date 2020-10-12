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
    public class ReportHSDetailtPartnerLHController : ApiController
    {

        /// <summary>
        /// List Report từ ngày đến ngày
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpGet]
        public List<ReportDetailtForPartner> SearchDataReportDetailtForPartnerForAll(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            try
            {
                List<ReportDetailtForPartner> listReport = new HSReportBL().SearchDataDetailtForPartnerForAll(fromDate, toDate, reportTypeID);
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
        public List<ReportDetailtForPartner> SearchDataReportDetailtForPartnerForAllForMonth(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            try
            {
                List<ReportDetailtForPartner> listReport = new HSReportBL().SearchDataDetailtForPartnerForAllForMonth(fromDate, toDate, reportTypeID);
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
        public List<ReportDetailtForPartner> SearchDataReportDetailtForPartnerForAllforYear(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            try
            {
                List<ReportDetailtForPartner> listReport = new HSReportBL().SearchDataDetailtForPartnerForAllForYear(fromDate, toDate, reportTypeID);
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
        public List<ReportDetailtForPartner> SearchDataReportDetailtForOneForDay(DateTime fromDate, DateTime toDate, string reportTypeID, string partnerID)
        {
            try
            {
                List<ReportDetailtForPartner> listReport = new HSReportBL().SearchDataDetailtForOnePartnerForDay(fromDate, toDate, reportTypeID, partnerID);
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
        public List<ReportDetailtForPartner> SearchDataReportDetailtForOneForMonth(DateTime fromDate, DateTime toDate, string reportTypeID, string partnerID)
        {
            try
            {
                List<ReportDetailtForPartner> listReport = new HSReportBL().SearchDataDetailtForOnePartnerForMonth(fromDate, toDate, reportTypeID, partnerID);
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
        public List<ReportDetailtForPartner> SearchDataReportDetailtForOneForYear(DateTime fromDate, DateTime toDate, string reportTypeID, string partnerID)
        {
            try
            {
                List<ReportDetailtForPartner> listReport = new HSReportBL().SearchDataDetailtForOnePartnerForYear(fromDate, toDate, reportTypeID, partnerID);
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
        public List<ReportDetailtForPartner> SearchDataReportDetailtGradationForAll(int toYear, int typeID, string reportTypeID)
        {
            try
            {
                List<ReportDetailtForPartner> result = new HSReportBL().SearchDataReportDetailtGradationForPartnerForAll(toYear, typeID, reportTypeID);
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
        public List<ReportDetailtForPartner> SearchDataReportDetailtGradationForOne(int toYear, int typeID, string reportTypeID, string partnerID)
        {
            try
            {
                List<ReportDetailtForPartner> result = new HSReportBL().SearchDataReportDetailtGradationForPartnerForOne(toYear, typeID, reportTypeID, partnerID);
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
        public List<ReportDetailtForPartner> SearchDataReportDetailtCompareMonthForAll(int toYear, int toMonth, string reportTypeID)
        {
            try
            {
                List<ReportDetailtForPartner> result = new HSReportBL().SearchDataReportDetailtCompareMonthForPartnerForAll(toYear, toMonth, reportTypeID);
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
        public List<ReportDetailtForPartner> SearchDataReportDetailtCompareMonthForOne(int toYear, int toMonth, string reportTypeID, string partnerID)
        {
            try
            {
                List<ReportDetailtForPartner> result = new HSReportBL().SearchDataReportDetailtCompareMonthForPartnerForOne(toYear, toMonth, reportTypeID, partnerID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }
    }
}

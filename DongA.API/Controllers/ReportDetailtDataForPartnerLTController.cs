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
    public class ReportDetailtDataForPartnerLTController : ApiController
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
                List<ReportDetailtForTotalMoneyType> listReport = new ReportBL().SearchDataDetailtForPartnerLTForAll(fromDate, toDate, reportTypeID);
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
        public List<ReportDetailtForTotalMoneyType> SearchDataReportDetailtForPartnerForAllConvert(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            try
            {
                List<ReportDetailtForTotalMoneyType> listReport = new ReportBL().SearchDataDetailtForPartnerLTForAllConvert(fromDate, toDate, reportTypeID);
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
                List<ReportDetailtForTotalMoneyType> listReport = new ReportBL().SearchDataDetailtForOnePartnerLTForDay(fromDate, toDate, reportTypeID, partnerID);
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
        public List<ReportDetailtForTotalMoneyType> SearchDataReportDetailtForOneForDayConvert(DateTime fromDate, DateTime toDate, string reportTypeID, string partnerID)
        {
            try
            {
                List<ReportDetailtForTotalMoneyType> listReport = new ReportBL().SearchDataDetailtForOnePartnerLTForDayConvert(fromDate, toDate, reportTypeID, partnerID);
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
                List<ReportDetailtForTotalMoneyType> listReport = new ReportBL().SearchDataDetailtForOnePartnerLTForMonth(fromDate, toDate, reportTypeID, partnerID);
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
        public List<ReportDetailtForTotalMoneyType> SearchDataReportDetailtForOneForMonthConvert(DateTime fromDate, DateTime toDate, string reportTypeID, string partnerID)
        {
            try
            {
                List<ReportDetailtForTotalMoneyType> listReport = new ReportBL().SearchDataDetailtForOnePartnerLTForMonthConvert(fromDate, toDate, reportTypeID, partnerID);
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
                List<ReportDetailtForTotalMoneyType> listReport = new ReportBL().SearchDataDetailtForOnePartnerLTForYear(fromDate, toDate, reportTypeID, partnerID);
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
        public List<ReportDetailtForTotalMoneyType> SearchDataReportDetailtForOneForYearConvert(DateTime fromDate, DateTime toDate, string reportTypeID, string partnerID)
        {
            try
            {
                List<ReportDetailtForTotalMoneyType> listReport = new ReportBL().SearchDataDetailtForOnePartnerLTForYearConvert(fromDate, toDate, reportTypeID, partnerID);
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
                List<ReportDetailtForTotalMoneyType> result = new ReportBL().SearchDataReportDetailtGradationForPartnerLTForAll(toYear, typeID, reportTypeID);
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
        public List<ReportDetailtForTotalMoneyType> SearchDataReportDetailtGradationForAllConvert(int toYear, int typeID, string reportTypeID)
        {
            try
            {
                List<ReportDetailtForTotalMoneyType> result = new ReportBL().SearchDataReportDetailtGradationForPartnerLTForAllConvert(toYear, typeID, reportTypeID);
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
                List<ReportDetailtForTotalMoneyType> result = new ReportBL().SearchDataReportDetailtGradationForPartnerLTForOne(toYear, typeID, reportTypeID, partnerID);
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
        public List<ReportDetailtForTotalMoneyType> SearchDataReportDetailtGradationForOneConvert(int toYear, int typeID, string reportTypeID, string partnerID)
        {
            try
            {
                List<ReportDetailtForTotalMoneyType> result = new ReportBL().SearchDataReportDetailtGradationForPartnerLTForOneConvert(toYear, typeID, reportTypeID, partnerID);
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
                List<ReportDetailtForTotalMoneyType> result = new ReportBL().SearchDataReportDetailtCompareMonthForPartnerLTForAll(toYear, toMonth, reportTypeID);
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
        public List<ReportDetailtForTotalMoneyType> SearchDataReportDetailtCompareMonthForAllConvert(int toYear, int toMonth, string reportTypeID)
        {
            try
            {
                List<ReportDetailtForTotalMoneyType> result = new ReportBL().SearchDataReportDetailtCompareMonthForPartnerLTForAllConvert(toYear, toMonth, reportTypeID);
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
                List<ReportDetailtForTotalMoneyType> result = new ReportBL().SearchDataReportDetailtCompareMonthForPartnerLTForOne(toYear, toMonth, reportTypeID, partnerID);
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
        public List<ReportDetailtForTotalMoneyType> SearchDataReportDetailtCompareMonthForOneConvert(int toYear, int toMonth, string reportTypeID, string partnerID)
        {
            try
            {
                List<ReportDetailtForTotalMoneyType> result = new ReportBL().SearchDataReportDetailtCompareMonthForPartnerLTForOneConvert(toYear, toMonth, reportTypeID, partnerID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }
    }
}

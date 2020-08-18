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
    public class ReportDetailtDataController : ApiController
    {
        [HttpGet]
        public List<ReportDetailtServiceType> DataReportDetailtDay(int reportTypeID)
        {
            try
            {
                DateTime now = DateTime.Now;
                List<ReportDetailtServiceType> listReport = new ReportBL().DataReportDetailtDay(now, reportTypeID);
                return listReport;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// List Report detailt từ ngày đến ngày
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpGet]
        public List<ReportDetailtServiceType> SearchReportDay(DateTime fromDate, DateTime toDate, int reportTypeID)
        {
            try
            {
                List<ReportDetailtServiceType> listReport = new ReportBL().SearchReportDetailtDay(fromDate, toDate, reportTypeID);
                return listReport;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// List Report theo tháng
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpGet]
        public List<ReportDetailtServiceType> MarketForPartner(string reportTypeID)
        {
            try
            {
                DateTime now = DateTime.Now;
                List<ReportDetailtServiceType> listReport = new ReportBL().GetListReportDetailtMonth(now, reportTypeID);
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
        public List<ReportDetailtServiceType> SearchMarketForPartner(DateTime fromDate, DateTime toDate, string partnerCode, string reportTypeID)
        {
            try
            {
                List<ReportDetailtServiceType> listReport = new ReportBL().SearchReportDetailtMonth(fromDate, toDate, partnerCode, reportTypeID);
                return listReport;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// List Report theo tháng
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpGet]
        public List<ReportDetailtServiceType> ReportYear(string reportTypeID)
        {
            try
            {
                DateTime now = DateTime.Now;
                List<ReportDetailtServiceType> listReport = new ReportBL().GetListReportDetailtYear(now, reportTypeID);
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
        public List<ReportDetailtServiceType> SearchReportYear(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            try
            {
                List<ReportDetailtServiceType> listReport = new ReportBL().SearchReportDetailtYear(fromDate, toDate, reportTypeID);
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
        public List<ReportDetailtServiceType> ListDataGradationCompare(int toYear, int typeID, string reportTypeID)
        {
            try
            {
                List<ReportDetailtServiceType> result = new ReportBL().ListDataDetailtGradationCompare(toYear, typeID, reportTypeID);
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
        public List<ReportDetailtServiceType> ListDataCompareMonth(int toYear, int toMonth, string reportTypeID)
        {
            try
            {
                List<ReportDetailtServiceType> result = new ReportBL().ListDataDetailtCompareMonth(toYear, toMonth, reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// Danh sách đối tác
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/08/2020]
        /// </history>
        [HttpGet]
        public List<Partner> ListPartNer()
        {
            try
            {
                List<Partner> listData = new ReportBL().ListPartNer();
                return listData;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

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
    }
}

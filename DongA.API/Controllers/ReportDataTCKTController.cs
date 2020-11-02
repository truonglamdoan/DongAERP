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
    public class ReportDataTCKTController : ApiController
    {
        /// <summary>
        /// List Report từ ngày đến ngày
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpGet]
        public List<ReportDetailtSTMarket> SearchReportDay(DateTime fromDate, DateTime toDate, int reportTypeID)
        {
            try
            {
                List<ReportDetailtSTMarket> listReport = new ReportBL().SearchTCKTReportDay(fromDate, toDate, reportTypeID);
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
        public List<ReportDetailtSTMarket> SearchReportMonth(DateTime fromDate, DateTime toDate, int reportTypeID)
        {
            try
            {
                List<ReportDetailtSTMarket> listReport = new ReportBL().SearchTCKTReportMonth(fromDate, toDate, reportTypeID);
                return listReport;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }
    }
}

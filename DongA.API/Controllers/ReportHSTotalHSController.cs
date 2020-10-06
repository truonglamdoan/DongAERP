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
    public class ReportHSTotalHSController : ApiController
    {

        /// <summary>
        /// List Report từ ngày đến ngày
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpGet]
        public List<ReportForTotalPayment> SearchReportDay(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            try
            {
                List<ReportForTotalPayment> listReport = new HSReportBL().SearchReportHSTotalHSForDay(fromDate, toDate, reportTypeID);
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
        public List<ReportForTotalPayment> SearchReportMonth(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            try
            {
                List<ReportForTotalPayment> listReport = new HSReportBL().SearchReportHSTotalHSForMonth(fromDate, toDate, reportTypeID);
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
        public List<ReportForTotalPayment> SearchReportYear(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            try
            {
                List<ReportForTotalPayment> listReport = new HSReportBL().SearchReportHSTotalHSForYear(fromDate, toDate, reportTypeID);
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
        public List<ReportForTotalPayment> ListDataGradationCompare(int toYear, int typeID, string reportTypeID)
        {
            try
            {
                List<ReportForTotalPayment> result = new HSReportBL().SearchReportHSTotalHSForGradationCompare(toYear, typeID, reportTypeID);
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
        public List<ReportForTotalPayment> ListDataCompareMonth(int toYear, int toMonth, string reportTypeID)
        {
            try
            {
                List<ReportForTotalPayment> result = new HSReportBL().SearchReportHSTotalHSForCompareMonth(toYear, toMonth, reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }
    }
}

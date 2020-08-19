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
    public class ReportMaketDataController : ApiController
    {

        [HttpGet]
        public List<ReportForMaket> CreateDataMarket()
        {
            try
            {
                List<ReportForMaket> listReport = new ReportBL().CreateDataMarket();
                return listReport;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        [HttpGet]
        public List<ReportForMaket> ReportDay(string reportTypeID)
        {
            try
            {
                DateTime now = DateTime.Now;
                List<ReportForMaket> listReport = new ReportBL().ReportMaketForDay(now, reportTypeID);
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
        public List<ReportForMaket> SearchReportDay(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            try
            {
                List<ReportForMaket> listReport = new ReportBL().SearchReportMaketForDay(fromDate, toDate, reportTypeID);
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
        public List<ReportForMaket> ReportMonth(string reportTypeID)
        {
            try
            {
                DateTime now = DateTime.Now;
                List<ReportForMaket> listReport = new ReportBL().GetListReportMarketForMonth(now, reportTypeID);
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
        public List<ReportForMaket> SearchReportMonth(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            try
            {
                List<ReportForMaket> listReport = new ReportBL().SearchReportMaketForMonth(fromDate, toDate, reportTypeID);
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
        public List<ReportForMaket> ReportYear(string reportTypeID)
        {
            try
            {
                DateTime now = DateTime.Now;
                List<ReportForMaket> listReport = new ReportBL().GetListReportMaketForYear(now, reportTypeID);
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
        public List<ReportForMaket> SearchReportYear(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            try
            {
                List<ReportForMaket> listReport = new ReportBL().SearchReportMaketForYear(fromDate, toDate, reportTypeID);
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
        public List<ReportForMaket> ListDataGradationCompare(int toYear, int typeID, string reportTypeID)
        {
            try
            {
                List<ReportForMaket> result = new ReportBL().ListDataMaketForGradationCompare(toYear, typeID, reportTypeID);
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
        public List<ReportForMaket> ListDataCompareMonth(int toYear, int toMonth, string reportTypeID)
        {
            try
            {
                List<ReportForMaket> result = new ReportBL().ListDataMaketForCompareMonth(toYear, toMonth, reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }
    }
}

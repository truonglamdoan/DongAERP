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
    public class ReportTMTDataController : ApiController
    {

        DateTime now = DateTime.Now;

        /// <summary>
        /// List Report theo ngày - Nguyên tệ
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpGet]
        public List<ReportForTotalMoneyType> DataReportForDay(string reportTypeID)
        {
            try
            {
                List<ReportForTotalMoneyType> result = new ReportBL().DataReportForDay(now, reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// Danh sách report tìm kiếm theo ngày - Nguyên tệ
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpGet]
        public List<ReportForTotalMoneyType> SearchReportForDay(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            try
            {
                List<ReportForTotalMoneyType> result = new ReportBL().SearchReportForDay(fromDate, toDate, reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// List Report theo ngày - USD
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpGet]
        public List<ReportForTotalMoneyType> DataReportForDayConvert(string reportTypeID)
        {
            try
            {
                List<ReportForTotalMoneyType> result = new ReportBL().DataReportForDayConvert(now, reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// Danh sách report tìm kiếm theo ngày - USD
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpGet]
        public List<ReportForTotalMoneyType> SearchReportForDayConvert(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            try
            {
                List<ReportForTotalMoneyType> result = new ReportBL().SearchReportForDayConvert(fromDate, toDate, reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// List Report theo tháng - Nguyên tệ
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpGet]
        public List<ReportForTotalMoneyType> DataReportypeForMonth(string reportTypeID)
        {
            try
            {
                List<ReportForTotalMoneyType> result = new ReportBL().DataReportypeForMonth(now, reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// List Report theo tháng - USD
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpGet]
        public List<ReportForTotalMoneyType> DataReportypeForMonthConvert(string reportTypeID)
        {
            try
            {
                List<ReportForTotalMoneyType> result = new ReportBL().DataReportypeForMonthConvert(now, reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// Danh sách report tìm kiếm theo tháng - Nguyên tệ
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpGet]
        public List<ReportForTotalMoneyType> SearchReportForMonth(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            try
            {
                List<ReportForTotalMoneyType> result = new ReportBL().SearchReportForMonth(fromDate, toDate, reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// Danh sách report theo tháng - Nguyên tệ
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpGet]
        public List<ReportForTotalMoneyType> SearchReportForMonthConvert(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            try
            {
                List<ReportForTotalMoneyType> result = new ReportBL().SearchReportForMonthConvert(fromDate, toDate, reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// List Report theo năm - Nguyên tệ
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpGet]
        public List<ReportForTotalMoneyType> DataReportypeForYear(string reportTypeID)
        {
            try
            {
                List<ReportForTotalMoneyType> result = new ReportBL().DataReportypeForYear(now, reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// List Report theo năm - USD
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpGet]
        public List<ReportForTotalMoneyType> DataReportypeForYearConvert(string reportTypeID)
        {
            try
            {
                List<ReportForTotalMoneyType> result = new ReportBL().DataReportypeForYearConvert(now, reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// Danh sách report tìm kiếm theo năm - Nguyên tệ
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpGet]
        public List<ReportForTotalMoneyType> SearchReportForYear(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            try
            {
                List<ReportForTotalMoneyType> result = new ReportBL().SearchReportForYear(fromDate, toDate, reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// Danh sách report tìm kiếm theo năm - USD
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpGet]
        public List<ReportForTotalMoneyType> SearchReportForYearConvert(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            try
            {
                List<ReportForTotalMoneyType> result = new ReportBL().SearchReportForYearConvert(fromDate, toDate, reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// List Report theo giai đoạn - Nguyên tệ
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpGet]
        public List<ReportForTotalMoneyType> DataReportypeForGradationCompare(int year, int typeID, string reportTypeID)
        {
            try
            {
                List<ReportForTotalMoneyType> result = new ReportBL().DataReportypeForGradationCompare(year, typeID, reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// List Report theo giai đoạn - USD
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpGet]
        public List<ReportForTotalMoneyType> DataReportypeForGradationCompareConvert(int year, int typeID, string reportTypeID)
        {
            try
            {
                List<ReportForTotalMoneyType> result = new ReportBL().DataReportypeForGradationCompareConvert(year, typeID, reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// List Report cho so sánh last month - Nguyên tệ
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpGet]
        public List<ReportForTotalMoneyType> DataReportypeCompareForMonth(int toYear, int toMonth, string reportTypeID)
        {
            try
            {
                List<ReportForTotalMoneyType> result = new ReportBL().DataReportypeCompareForMonth(toYear, toMonth, reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// List Report cho so sánh last month - USD
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpGet]
        public List<ReportForTotalMoneyType> DataReportypeCompareForMonthConvert(int toYear, int toMonth, string reportTypeID)
        {
            try
            {
                List<ReportForTotalMoneyType> result = new ReportBL().DataReportypeCompareForMonthConvert(toYear, toMonth, reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }
    }
}

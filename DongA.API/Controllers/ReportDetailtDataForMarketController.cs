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

        /// <summary>
        /// Danh sách thị trường
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/08/2020]
        /// </history>
        [HttpGet]
        public List<Partner> ListPartner()
        {
            try
            {
                List<Partner> listData = new ReportBL().ListPartner();
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

        /// <summary>
        /// List Report từ tháng đến tháng
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpGet]
        public List<ReportDetailtSTMarket> SearchDataReportDetailtMonth(DateTime fromDate, DateTime toDate, string reportTypeID, string marketID)
        {
            try
            {
                List<ReportDetailtSTMarket> listReport = new ReportBL().SearchDataReportDetailtMonth(fromDate, toDate, reportTypeID, marketID);
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
        public List<ReportDetailtSTMarket> SearchDataReportDetailtYear(DateTime fromDate, DateTime toDate, string reportTypeID, string marketID)
        {
            try
            {
                List<ReportDetailtSTMarket> listReport = new ReportBL().SearchDataReportDetailtYear(fromDate, toDate, reportTypeID, marketID);
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
        /// List Report detailt từ tháng đến tháng cho báo cáo chi tiết
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpGet]
        public List<ReportDetailtServiceType> SearchDataReportDetailtForOneMarketForMonth(DateTime fromDate, DateTime toDate, string reportTypeID, string marketID)
        {
            try
            {
                List<ReportDetailtServiceType> listReport = new ReportBL().SearchDataReportDetailtForOneMarketForMonth(fromDate, toDate, reportTypeID, marketID);
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
        public List<ReportDetailtServiceType> SearchDataReportDetailtForOneMarketForYear(DateTime fromDate, DateTime toDate, string reportTypeID, string marketID)
        {
            try
            {
                List<ReportDetailtServiceType> listReport = new ReportBL().SearchDataReportDetailtForOneMarketForYear(fromDate, toDate, reportTypeID, marketID);
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
        public List<ReportDetailtServiceType> SearchDataReportDetailtGradationForAll(int toYear, int typeID, string reportTypeID, string marketID)
        {
            try
            {
                List<ReportDetailtServiceType> result = new ReportBL().SearchDataReportDetailtGradationForAll(toYear, typeID, reportTypeID, marketID);
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
        public List<ReportDetailtSTMarket> SearchDataReportDetailtCompareMonthForAll(int toYear, int toMonth, string reportTypeID, string marketID)
        {
            try
            {
                List<ReportDetailtSTMarket> result = new ReportBL().SearchDataReportDetailtCompareMonthForAll(toYear, toMonth, reportTypeID, marketID);
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

        #region Báo cáo chi tiết thị trường theo loại tiền chi trả
        /// <summary>
        /// Get doanh số chi trả  theo thị trường cho loại tiền chi trả
        /// </summary>
        /// <param name="reportTypeID"></param>
        /// <returns></returns>
        [HttpGet]
        public List<ReportDetailtForTotalMoneyType> DataReportDetailtMTForAll(int reportTypeID)
        {
            try
            {
                DateTime now = DateTime.Now;
                string marketID = "0";
                List<ReportDetailtForTotalMoneyType> listReport = new ReportBL().SearchDataReportDetailtMTForAll(now, now, reportTypeID.ToString(), marketID);
                return listReport;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        ///  Get doanh số chi trả  theo thị trường cho loại tiền chi trả
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpGet]
        public List<ReportDetailtForTotalMoneyType> SearchDataReportDetailtMTForAll(DateTime fromDate, DateTime toDate, string reportTypeID, string marketID)
        {
            try
            {
                List<ReportDetailtForTotalMoneyType> listReport = new ReportBL().SearchDataReportDetailtMTForAll(fromDate, toDate, reportTypeID, marketID);
                return listReport;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// Get doanh số chi trả  theo thị trường cho loại tiền chi trả Quy USD
        /// </summary>
        /// <param name="reportTypeID"></param>
        /// <returns></returns>
        [HttpGet]
        public List<ReportDetailtForTotalMoneyType> DataReportDetailtMTForAllConvert(int reportTypeID)
        {
            try
            {
                DateTime now = DateTime.Now;
                string marketID = "0";
                List<ReportDetailtForTotalMoneyType> listReport = new ReportBL().SearchDataReportDetailtMTForAllConvert(now, now, reportTypeID.ToString(), marketID);
                return listReport;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        ///  Get doanh số chi trả  theo thị trường cho loại tiền chi trả Quy USD
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpGet]
        public List<ReportDetailtForTotalMoneyType> SearchDataReportDetailtMTForAllConvert(DateTime fromDate, DateTime toDate, string reportTypeID, string marketID)
        {
            try
            {
                List<ReportDetailtForTotalMoneyType> listReport = new ReportBL().SearchDataReportDetailtMTForAllConvert(fromDate, toDate, reportTypeID, marketID);
                return listReport;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// Get doanh số chi trả  theo thị trường cho loại tiền chi trả
        /// </summary>
        /// <param name="reportTypeID"></param>
        /// <returns></returns>
        [HttpGet]
        public List<ReportDetailtForTotalMoneyType> DataReportDetailtMTForOne(int reportTypeID)
        {
            try
            {
                DateTime now = DateTime.Now;
                string marketID = "0";
                List<ReportDetailtForTotalMoneyType> listReport = new ReportBL().SearchDataReportDetailtMTForOne(now, now, reportTypeID.ToString(), marketID);
                return listReport;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        ///  Get doanh số chi trả  theo thị trường cho loại tiền chi trả
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpGet]
        public List<ReportDetailtForTotalMoneyType> SearchDataReportDetailtMTForOne(DateTime fromDate, DateTime toDate, string reportTypeID, string marketID)
        {
            try
            {
                List<ReportDetailtForTotalMoneyType> listReport = new ReportBL().SearchDataReportDetailtMTForOne(fromDate, toDate, reportTypeID, marketID);
                return listReport;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// Get doanh số chi trả  theo thị trường cho loại tiền chi trả Quy USD
        /// </summary>
        /// <param name="reportTypeID"></param>
        /// <returns></returns>
        [HttpGet]
        public List<ReportDetailtForTotalMoneyType> DataReportDetailtMTForOneConvert(int reportTypeID)
        {
            try
            {
                DateTime now = DateTime.Now;
                string marketID = "0";
                List<ReportDetailtForTotalMoneyType> listReport = new ReportBL().SearchDataReportDetailtMTForOneConvert(now, now, reportTypeID.ToString(), marketID);
                return listReport;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        ///  Get doanh số chi trả  theo thị trường cho loại tiền chi trả Quy USD
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        [HttpGet]
        public List<ReportDetailtForTotalMoneyType> SearchDataReportDetailtMTForOneConvert(DateTime fromDate, DateTime toDate, string reportTypeID, string marketID)
        {
            try
            {
                List<ReportDetailtForTotalMoneyType> listReport = new ReportBL().SearchDataReportDetailtMTForOneConvert(fromDate, toDate, reportTypeID, marketID);
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
        public List<ReportDetailtForTotalMoneyType> SearchDataReportDetailtMTGradationForAll(int toYear, int typeID, string reportTypeID)
        {
            try
            {
                List<ReportDetailtForTotalMoneyType> result = new ReportBL().SearchDataReportDetailtMTGradationForAll(toYear, typeID, reportTypeID);
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
        public List<ReportDetailtForTotalMoneyType> SearchDataReportDetailtMTGradationForAllConvert(int toYear, int typeID, string reportTypeID)
        {
            try
            {
                List<ReportDetailtForTotalMoneyType> result = new ReportBL().SearchDataReportDetailtMTGradationForAllConvert(toYear, typeID, reportTypeID);
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
                List<ReportDetailtForTotalMoneyType> result = new ReportBL().SearchDataReportDetailtMTGradationForOne(toYear, typeID, reportTypeID, marketID);
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
        public List<ReportDetailtForTotalMoneyType> SearchDataReportDetailtMTGradationForOneConvert(int toYear, int typeID, string reportTypeID, string marketID)
        {
            try
            {
                List<ReportDetailtForTotalMoneyType> result = new ReportBL().SearchDataReportDetailtMTGradationForOneConvert(toYear, typeID, reportTypeID, marketID);
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
        public List<ReportDetailtForTotalMoneyType> SearchDataReportDetailtMTCompareMonthForAll(int toYear, int toMonth, string reportTypeID)
        {
            try
            {
                List<ReportDetailtForTotalMoneyType> result = new ReportBL().SearchDataReportDetailtMTCompareMonthForAll(toYear, toMonth, reportTypeID);
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
        public List<ReportDetailtForTotalMoneyType> SearchDataReportDetailtMTCompareMonthForAllConvert(int toYear, int toMonth, string reportTypeID)
        {
            try
            {
                List<ReportDetailtForTotalMoneyType> result = new ReportBL().SearchDataReportDetailtMTCompareMonthForAllConvert(toYear, toMonth, reportTypeID);
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
                List<ReportDetailtForTotalMoneyType> result = new ReportBL().SearchDataReportDetailtMTCompareMonthForOne(toYear, toMonth, reportTypeID, marketID);
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
        public List<ReportDetailtForTotalMoneyType> SearchDataReportDetailtMTCompareMonthForOneConvert(int toYear, int toMonth, string reportTypeID, string marketID)
        {
            try
            {
                List<ReportDetailtForTotalMoneyType> result = new ReportBL().SearchDataReportDetailtMTCompareMonthForOneConvert(toYear, toMonth, reportTypeID, marketID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }
        #endregion
    }
}

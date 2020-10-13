using DongA.Core;
using DongA.Entities;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Web.Configuration;
using System.Web.ModelBinding;
using Newtonsoft.Json;

namespace DongA.DataAccess
{
    public class HSReportDAL : DongABaseDAL
    {
        /// <summary>
        /// List Report
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<Report> SearchDay(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<Report>();
                result = DongADatabase.ToDataAPIObject<Report>("ReportHSType", "SearchReportDay", "fromDate", fromDate, "toDate", toDate, "reportTypeID", reportTypeID);

                return result;
            }
            catch (Exception ex)
            {
                throw DongAException.FromCommand(command, ex);
            }
        }

        /// <summary>
        /// List Report
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<Report> SearchMonth(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<Report>();
                result = DongADatabase.ToDataAPIObject<Report>("ReportHSType", "SearchReportMonth", "fromDate", fromDate, "toDate", toDate, "reportTypeID", reportTypeID);

                return result;
            }
            catch (Exception ex)
            {
                throw DongAException.FromCommand(command, ex);
            }
        }

        /// <summary>
        /// List Report
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<Report> SearchYear(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<Report>();
                result = DongADatabase.ToDataAPIObject<Report>("ReportHSType", "SearchReportYear", "fromDate", fromDate, "toDate", toDate, "reportTypeID", reportTypeID);

                return result;
            }
            catch (Exception ex)
            {
                throw DongAException.FromCommand(command, ex);
            }
        }

        /// <summary>
        /// List Report cho so sánh giai đoạn
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<Report> SearchGradationCompare(int toYear, int typeID, string reportTypeID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<Report>();
                result = DongADatabase.ToDataAPIObject<Report>("ReportHSType", "SearchReportGradationCompare", "toYear", toYear, "typeID", typeID, "reportTypeID", reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw DongAException.FromCommand(command, ex);
            }
        }

        /// <summary>
        /// List Report cho so sánh giai đoạn
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<Report> SearchMonthCompareGrid(int toYear, int toMonth, string reportTypeID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<Report>();
                result = DongADatabase.ToDataAPIObject<Report>("ReportHSType", "SearchReportCompareMonth", "toYear", toYear, "toMonth", toMonth, "reportTypeID", reportTypeID);

                return result;
            }
            catch (Exception ex)
            {
                throw DongAException.FromCommand(command, ex);
            }
        }

        #region Get dữ liệu cho loại tiền
        /// <summary>
        /// List Report
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportForTotalMoneyType> SearchReportHSTotalMoneyTypeForDay(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportForTotalMoneyType>();
                result = DongADatabase.ToDataAPIObject<ReportForTotalMoneyType>("ReportHSMoneyType", "SearchReportForDay", "fromDate", fromDate, "toDate", toDate, "reportTypeID", reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw DongAException.FromCommand(command, ex);
            }
        }

        /// <summary>
        /// List Report
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportForTotalMoneyType> SearchReportHSTotalMoneyTypeForMonth(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportForTotalMoneyType>();
                result = DongADatabase.ToDataAPIObject<ReportForTotalMoneyType>("ReportHSMoneyType", "SearchReportForMonth", "fromDate", fromDate, "toDate", toDate, "reportTypeID", reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw DongAException.FromCommand(command, ex);
            }
        }

        /// <summary>
        /// List Report
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportForTotalMoneyType> SearchReportHSTotalMoneyTypeForYear(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportForTotalMoneyType>();
                result = DongADatabase.ToDataAPIObject<ReportForTotalMoneyType>("ReportHSMoneyType", "SearchReportForYear", "fromDate", fromDate, "toDate", toDate, "reportTypeID", reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw DongAException.FromCommand(command, ex);
            }
        }


        /// <summary>
        /// List Report theo ngày
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportForTotalMoneyType> SearchReportHSTotalMoneyTypeForGradationCompare(int year, int typeID, string reportTypeID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportForTotalMoneyType>();
                result = DongADatabase.ToDataAPIObject<ReportForTotalMoneyType>("ReportHSMoneyType", "SearchForGradationCompare", "year", year, "typeID", typeID, "reportTypeID", reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw DongAException.FromCommand(command, ex);
            }
        }


        /// <summary>
        /// List Report cho so sánh theo tháng - Nguyên tệ
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportForTotalMoneyType> SearchReportHSTotalMoneyTypeForCompareMonth(int toYear, int toMonth, string reportTypeID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportForTotalMoneyType>();
                result = DongADatabase.ToDataAPIObject<ReportForTotalMoneyType>("ReportHSMoneyType", "SearchForCompareMonth", "toYear", toYear, "toMonth", toMonth, "reportTypeID", reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw DongAException.FromCommand(command, ex);
            }
        }
        #endregion

        #region Get dữ liệu cho thị trường

        /// <summary>
        /// List Report
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportForMaket> SearchReportMaketForDay(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportForMaket>();
                result = DongADatabase.ToDataAPIObject<ReportForMaket>("ReportHSMarket", "SearchReportDay", "fromDate", fromDate, "toDate", toDate, "reportTypeID", reportTypeID);

                return result;
            }
            catch (Exception ex)
            {
                throw DongAException.FromCommand(command, ex);
            }
        }

        /// <summary>
        /// List Report
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportForMaket> SearchReportMaketForMonth(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportForMaket>();
                result = DongADatabase.ToDataAPIObject<ReportForMaket>("ReportHSMarket", "SearchReportMonth", "fromDate", fromDate, "toDate", toDate, "reportTypeID", reportTypeID);

                return result;
            }
            catch (Exception ex)
            {
                throw DongAException.FromCommand(command, ex);
            }
        }

        /// <summary>
        /// List Report
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportForMaket> SearchReportMaketForYear(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportForMaket>();
                result = DongADatabase.ToDataAPIObject<ReportForMaket>("ReportHSMarket", "SearchReportYear", "fromDate", fromDate, "toDate", toDate, "reportTypeID", reportTypeID);

                return result;
            }
            catch (Exception ex)
            {
                throw DongAException.FromCommand(command, ex);
            }
        }


        /// <summary>
        /// List Report theo ngày
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportForMaket> SearchReportMaketForGradationCompare(int toYear, int typeID, string reportTypeID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportForMaket>();
                result = DongADatabase.ToDataAPIObject<ReportForMaket>("ReportHSMarket", "ListDataGradationCompare", "toYear", toYear, "typeID", typeID, "reportTypeID", reportTypeID);

                return result;
            }
            catch (Exception ex)
            {
                throw DongAException.FromCommand(command, ex);
            }
        }


        /// <summary>
        /// List Report cho so sánh giai đoạn
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportForMaket> SearchReportMaketForCompareMonth(int toYear, int toMonth, string reportTypeID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportForMaket>();
                result = DongADatabase.ToDataAPIObject<ReportForMaket>("ReportHSMarket", "ListDataCompareMonth", "toYear", toYear, "toMonth", toMonth, "reportTypeID", reportTypeID);

                return result;
            }
            catch (Exception ex)
            {
                throw DongAException.FromCommand(command, ex);
            }
        }
        #endregion

        #region Get dữ liệu cho Tổng hồ sơ

        /// <summary>
        /// List Report
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportForTotalPayment> SearchReportTotalHSForDay(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportForTotalPayment>();
                result = DongADatabase.ToDataAPIObject<ReportForTotalPayment>("ReportHSTotalHS", "SearchReportDay", "fromDate", fromDate, "toDate", toDate, "reportTypeID", reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw DongAException.FromCommand(command, ex);
            }
        }


        /// <summary>
        /// List Report
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportForTotalPayment> SearchReportTotalHSForMonth(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportForTotalPayment>();
                result = DongADatabase.ToDataAPIObject<ReportForTotalPayment>("ReportHSTotalHS", "SearchReportMonth", "fromDate", fromDate, "toDate", toDate, "reportTypeID", reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw DongAException.FromCommand(command, ex);
            }
        }

        /// <summary>
        /// List Report
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportForTotalPayment> SearchReportTotalHSForYear(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportForTotalPayment>();
                result = DongADatabase.ToDataAPIObject<ReportForTotalPayment>("ReportHSTotalHS", "SearchReportYear", "fromDate", fromDate, "toDate", toDate, "reportTypeID", reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw DongAException.FromCommand(command, ex);
            }
        }


        /// <summary>
        /// List Report theo ngày
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportForTotalPayment> SearchReportTotalHSForGradationCompare(int toYear, int typeID, string reportTypeID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportForTotalPayment>();
                result = DongADatabase.ToDataAPIObject<ReportForTotalPayment>("ReportHSTotalHS", "ListDataGradationCompare", "toYear", toYear, "typeID", typeID, "reportTypeID", reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw DongAException.FromCommand(command, ex);
            }
        }


        /// <summary>
        /// List Report theo ngày
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportForTotalPayment> SearchReportTotalHSForCompareMonth(int toYear, int toMonth, string reportTypeID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportForTotalPayment>();
                result = DongADatabase.ToDataAPIObject<ReportForTotalPayment>("ReportHSTotalHS", "ListDataCompareMonth", "toYear", toYear, "toMonth", toMonth, "reportTypeID", reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw DongAException.FromCommand(command, ex);
            }
        }
        #endregion

        #region Get dữ liệu hồ sơ chi tiết loại hình

        /// <summary>
        /// List Report
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtSTMarket> SearchMarketForTotalForDay(DateTime fromDate, DateTime toDate, string reportTypeID, string marketID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportDetailtSTMarket>();
                result = DongADatabase.ToDataAPIObject<ReportDetailtSTMarket>("ReportHSDetailtTypeLH", "SearchDataReportDetailtDay", "fromDate", fromDate, "toDate", toDate, "reportTypeID", reportTypeID, "marketID", marketID);

                return result;
            }
            catch (Exception ex)
            {
                throw DongAException.FromCommand(command, ex);
            }
        }

        /// <summary>
        /// List Report
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtSTMarket> SearchMarketForTotalForMonth(DateTime fromDate, DateTime toDate, string reportTypeID, string marketID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportDetailtSTMarket>();
                result = DongADatabase.ToDataAPIObject<ReportDetailtSTMarket>("ReportHSDetailtTypeLH", "SearchDataReportDetailtMonth", "fromDate", fromDate, "toDate", toDate, "reportTypeID", reportTypeID, "marketID", marketID);

                return result;
            }
            catch (Exception ex)
            {
                throw DongAException.FromCommand(command, ex);
            }
        }

        /// <summary>
        /// List Report
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtSTMarket> SearchMarketForTotalForYear(DateTime fromDate, DateTime toDate, string reportTypeID, string marketID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportDetailtSTMarket>();
                result = DongADatabase.ToDataAPIObject<ReportDetailtSTMarket>("ReportHSDetailtTypeLH", "SearchDataReportDetailtYear", "fromDate", fromDate, "toDate", toDate, "reportTypeID", reportTypeID, "marketID", marketID);

                return result;
            }
            catch (Exception ex)
            {
                throw DongAException.FromCommand(command, ex);
            }
        }
        
        /// <summary>
        /// List Report
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtServiceType> SearchMarketForOne(DateTime fromDate, DateTime toDate, string reportTypeID, string marketID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportDetailtServiceType>();
                result = DongADatabase.ToDataAPIObject<ReportDetailtServiceType>("ReportHSDetailtTypeLH", "SearchDataReportDetailtForOneMarketForDay", "fromDate", fromDate, "toDate", toDate, "reportTypeID", reportTypeID, "marketID", marketID);

                return result;
            }
            catch (Exception ex)
            {
                throw DongAException.FromCommand(command, ex);
            }
        }

        /// <summary>
        /// List Report
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtServiceType> SearchMarketForOneForMonth(DateTime fromDate, DateTime toDate, string reportTypeID, string marketID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportDetailtServiceType>();
                result = DongADatabase.ToDataAPIObject<ReportDetailtServiceType>("ReportHSDetailtTypeLH", "SearchDataReportDetailtForOneMarketForMonth", "fromDate", fromDate, "toDate", toDate, "reportTypeID", reportTypeID, "marketID", marketID);

                return result;
            }
            catch (Exception ex)
            {
                throw DongAException.FromCommand(command, ex);
            }
        }

        /// <summary>
        /// List Report
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtServiceType> SearchMarketForOneForYear(DateTime fromDate, DateTime toDate, string reportTypeID, string marketID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportDetailtServiceType>();
                result = DongADatabase.ToDataAPIObject<ReportDetailtServiceType>("ReportHSDetailtTypeLH", "SearchDataReportDetailtForOneMarketForYear", "fromDate", fromDate, "toDate", toDate, "reportTypeID", reportTypeID, "marketID", marketID);

                return result;
            }
            catch (Exception ex)
            {
                throw DongAException.FromCommand(command, ex);
            }
        }

        /// <summary>
        /// List Report cho so sánh giai đoạn
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtServiceType> ReportDetailtGradationCompareForAll(int toYear, int typeID, string reportTypeID, string marketID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportDetailtServiceType>();
                result = DongADatabase.ToDataAPIObject<ReportDetailtServiceType>("ReportHSDetailtTypeLH", "SearchDataReportDetailtGradationForAll", "toYear", toYear, "typeID", typeID, "reportTypeID", reportTypeID, "marketID", marketID);
                return result;
            }
            catch (Exception ex)
            {
                throw DongAException.FromCommand(command, ex);
            }
        }


        /// <summary>
        /// List Report cho so sánh giai đoạn
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtServiceType> ReportDetailtGradationCompareForOne(int toYear, int typeID, string reportTypeID, string marketID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportDetailtServiceType>();
                result = DongADatabase.ToDataAPIObject<ReportDetailtServiceType>("ReportHSDetailtTypeLH", "SearchDataReportDetailtGradationForOne", "toYear", toYear, "typeID", typeID, "reportTypeID", reportTypeID, "marketID", marketID);
                return result;
            }
            catch (Exception ex)
            {
                throw DongAException.FromCommand(command, ex);
            }
        }


        /// <summary>
        /// List Report cho so sánh giai đoạn
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtSTMarket> ReportDetailtCompareMonthForAll(int toYear, int toMonth, string reportTypeID, string marketID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportDetailtSTMarket>();
                result = DongADatabase.ToDataAPIObject<ReportDetailtSTMarket>("ReportHSDetailtTypeLH", "SearchDataReportDetailtCompareMonthForAll", "toYear", toYear, "toMonth", toMonth, "reportTypeID", reportTypeID, "marketID", marketID);
                return result;
            }
            catch (Exception ex)
            {
                throw DongAException.FromCommand(command, ex);
            }
        }


        /// <summary>
        /// List Report cho so sánh giai đoạn
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtServiceType> ReportDetailtCompareMonthForOne(int toYear, int toMonth, string reportTypeID, string marketID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportDetailtServiceType>();
                result = DongADatabase.ToDataAPIObject<ReportDetailtServiceType>("ReportHSDetailtTypeLH", "SearchDataReportDetailtCompareMonthForOne", "toYear", toYear, "toMonth", toMonth, "reportTypeID", reportTypeID, "marketID", marketID);
                return result;
            }
            catch (Exception ex)
            {
                throw DongAException.FromCommand(command, ex);
            }
        }

        #endregion

        #region Get dữ liệu hồ sơ chi tiết cho thị trường loại tiền

        /// <summary>
        /// List Report
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtForTotalMoneyType> SearchReportDetailtMTForAll(DateTime fromDate, DateTime toDate, string reportTypeID, string marketID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportDetailtForTotalMoneyType>();
                result = DongADatabase.ToDataAPIObject<ReportDetailtForTotalMoneyType>("ReportHSDetailtMarketLT", "SearchDataReportDetailtDay", "fromDate", fromDate, "toDate", toDate, "reportTypeID", reportTypeID, "marketID", marketID);

                return result;
            }
            catch (Exception ex)
            {
                throw DongAException.FromCommand(command, ex);
            }
        }

        /// <summary>
        /// List Report
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtForTotalMoneyType> SearchReportDetailtMTForAllForMonth(DateTime fromDate, DateTime toDate, string reportTypeID, string marketID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportDetailtForTotalMoneyType>();
                result = DongADatabase.ToDataAPIObject<ReportDetailtForTotalMoneyType>("ReportHSDetailtMarketLT", "SearchDataReportDetailtMonth", "fromDate", fromDate, "toDate", toDate, "reportTypeID", reportTypeID, "marketID", marketID);

                return result;
            }
            catch (Exception ex)
            {
                throw DongAException.FromCommand(command, ex);
            }
        }

        /// <summary>
        /// List Report
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtForTotalMoneyType> SearchReportDetailtMTForAllForYear(DateTime fromDate, DateTime toDate, string reportTypeID, string marketID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportDetailtForTotalMoneyType>();
                result = DongADatabase.ToDataAPIObject<ReportDetailtForTotalMoneyType>("ReportHSDetailtMarketLT", "SearchDataReportDetailtYear", "fromDate", fromDate, "toDate", toDate, "reportTypeID", reportTypeID, "marketID", marketID);

                return result;
            }
            catch (Exception ex)
            {
                throw DongAException.FromCommand(command, ex);
            }
        }
        
        /// <summary>
        /// List Report
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtForTotalMoneyType> SearchReportDetailtMTForOne(DateTime fromDate, DateTime toDate, string reportTypeID, string marketID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportDetailtForTotalMoneyType>();
                result = DongADatabase.ToDataAPIObject<ReportDetailtForTotalMoneyType>("ReportHSDetailtMarketLT", "SearchDataReportDetailtForOneMarketForDay", "fromDate", fromDate, "toDate", toDate, "reportTypeID", reportTypeID, "marketID", marketID);

                return result;
            }
            catch (Exception ex)
            {
                throw DongAException.FromCommand(command, ex);
            }
        }

        /// <summary>
        /// List Report
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtForTotalMoneyType> SearchReportDetailtMTForOneForMonth(DateTime fromDate, DateTime toDate, string reportTypeID, string marketID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportDetailtForTotalMoneyType>();
                result = DongADatabase.ToDataAPIObject<ReportDetailtForTotalMoneyType>("ReportHSDetailtMarketLT", "SearchDataReportDetailtForOneMarketForMonth", "fromDate", fromDate, "toDate", toDate, "reportTypeID", reportTypeID, "marketID", marketID);

                return result;
            }
            catch (Exception ex)
            {
                throw DongAException.FromCommand(command, ex);
            }
        }

        /// <summary>
        /// List Report
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtForTotalMoneyType> SearchReportDetailtMTForOneForYear(DateTime fromDate, DateTime toDate, string reportTypeID, string marketID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportDetailtForTotalMoneyType>();
                result = DongADatabase.ToDataAPIObject<ReportDetailtForTotalMoneyType>("ReportHSDetailtMarketLT", "SearchDataReportDetailtForOneMarketForYear", "fromDate", fromDate, "toDate", toDate, "reportTypeID", reportTypeID, "marketID", marketID);

                return result;
            }
            catch (Exception ex)
            {
                throw DongAException.FromCommand(command, ex);
            }
        }


        /// <summary>
        /// List Report cho so sánh giai đoạn
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtForTotalMoneyType> ReportDetailtMTGradationCompareForAll(int toYear, int typeID, string reportTypeID, string marketID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportDetailtForTotalMoneyType>();
                result = DongADatabase.ToDataAPIObject<ReportDetailtForTotalMoneyType>("ReportHSDetailtMarketLT", "SearchDataReportDetailtMTGradationForAll", "toYear", toYear, "typeID", typeID, "reportTypeID", reportTypeID, "marketID", marketID);
                return result;
            }
            catch (Exception ex)
            {
                throw DongAException.FromCommand(command, ex);
            }
        }


        /// <summary>
        /// List Report cho so sánh giai đoạn
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtForTotalMoneyType> ReportDetailtMTGradationCompareForOne(int toYear, int typeID, string reportTypeID, string marketID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportDetailtForTotalMoneyType>();
                result = DongADatabase.ToDataAPIObject<ReportDetailtForTotalMoneyType>("ReportHSDetailtMarketLT", "SearchDataReportDetailtMTGradationForOne", "toYear", toYear, "typeID", typeID, "reportTypeID", reportTypeID, "marketID", marketID);
                return result;
            }
            catch (Exception ex)
            {
                throw DongAException.FromCommand(command, ex);
            }
        }


        /// <summary>
        /// List Report cho so sánh giai đoạn
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtForTotalMoneyType> ReportDetailtMTCompareMonthForAll(int toYear, int toMonth, string reportTypeID, string marketID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportDetailtForTotalMoneyType>();
                result = DongADatabase.ToDataAPIObject<ReportDetailtForTotalMoneyType>("ReportHSDetailtMarketLT", "SearchDataReportDetailtMTCompareMonthForAll", "toYear", toYear, "toMonth", toMonth, "reportTypeID", reportTypeID, "marketID", marketID);
                return result;
            }
            catch (Exception ex)
            {
                throw DongAException.FromCommand(command, ex);
            }
        }


        /// <summary>
        /// List Report cho so sánh giai đoạn
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtForTotalMoneyType> ReportDetailtMTCompareMonthForOne(int toYear, int toMonth, string reportTypeID, string marketID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportDetailtForTotalMoneyType>();
                result = DongADatabase.ToDataAPIObject<ReportDetailtForTotalMoneyType>("ReportHSDetailtMarketLT", "SearchDataReportDetailtMTCompareMonthForOne", "toYear", toYear, "toMonth", toMonth, "reportTypeID", reportTypeID, "marketID", marketID);
                return result;
            }
            catch (Exception ex)
            {
                throw DongAException.FromCommand(command, ex);
            }
        }

        #endregion

        #region Báo cáo chi tiết hồ sơ cho đối tác loại hình
        /// <summary>
        /// List Report detailt cho đối tác
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtForPartner> SearchReportDetailtPartnerForDay(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportDetailtForPartner>();
                result = DongADatabase.ToDataAPIObject<ReportDetailtForPartner>("ReportHSDetailtPartnerLH", "SearchDataReportDetailtForPartnerForAll", "fromDate", fromDate, "toDate", toDate, "reportTypeID", reportTypeID);

                return result;
            }
            catch (Exception ex)
            {
                throw DongAException.FromCommand(command, ex);
            }
        }

        /// <summary>
        /// List Report detailt cho đối tác
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtForPartner> SearchReportDetailtPartnerForMonth(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportDetailtForPartner>();
                result = DongADatabase.ToDataAPIObject<ReportDetailtForPartner>("ReportHSDetailtPartnerLH", "SearchDataReportDetailtForPartnerForAllForMonth", "fromDate", fromDate, "toDate", toDate, "reportTypeID", reportTypeID);

                return result;
            }
            catch (Exception ex)
            {
                throw DongAException.FromCommand(command, ex);
            }
        }

        /// <summary>
        /// List Report detailt cho đối tác
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtForPartner> SearchReportDetailtPartnerForYear(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportDetailtForPartner>();
                result = DongADatabase.ToDataAPIObject<ReportDetailtForPartner>("ReportHSDetailtPartnerLH", "SearchDataReportDetailtForPartnerForAllForYear", "fromDate", fromDate, "toDate", toDate, "reportTypeID", reportTypeID);

                return result;
            }
            catch (Exception ex)
            {
                throw DongAException.FromCommand(command, ex);
            }
        }

        /// <summary>
        /// List Report
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtForPartner> SearchPartnerForOneForDay(DateTime fromDate, DateTime toDate, string reportTypeID, string partnerID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportDetailtForPartner>();
                result = DongADatabase.ToDataAPIObject<ReportDetailtForPartner>("ReportHSDetailtPartnerLH", "SearchDataReportDetailtForOneForDay", "fromDate", fromDate, "toDate", toDate, "reportTypeID", reportTypeID, "partnerID", partnerID);

                return result;
            }
            catch (Exception ex)
            {
                throw DongAException.FromCommand(command, ex);
            }
        }


        /// <summary>
        /// List Report
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtForPartner> SearchPartnerForOneForMonth(DateTime fromDate, DateTime toDate, string reportTypeID, string partnerID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportDetailtForPartner>();
                result = DongADatabase.ToDataAPIObject<ReportDetailtForPartner>("ReportHSDetailtPartnerLH", "SearchDataReportDetailtForOneForMonth", "fromDate", fromDate, "toDate", toDate, "reportTypeID", reportTypeID, "partnerID", partnerID);

                return result;
            }
            catch (Exception ex)
            {
                throw DongAException.FromCommand(command, ex);
            }
        }


        /// <summary>
        /// List Report
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtForPartner> SearchPartnerForOneforYear(DateTime fromDate, DateTime toDate, string reportTypeID, string partnerID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportDetailtForPartner>();
                result = DongADatabase.ToDataAPIObject<ReportDetailtForPartner>("ReportHSDetailtPartnerLH", "SearchDataReportDetailtForOneForYear", "fromDate", fromDate, "toDate", toDate, "reportTypeID", reportTypeID, "partnerID", partnerID);

                return result;
            }
            catch (Exception ex)
            {
                throw DongAException.FromCommand(command, ex);
            }
        }


        /// <summary>
        /// List Report cho so sánh giai đoạn
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtForPartner> ReportDetailtPartnerGradationCompareForAll(int toYear, int typeID, string reportTypeID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportDetailtForPartner>();
                result = DongADatabase.ToDataAPIObject<ReportDetailtForPartner>("ReportHSDetailtPartnerLH", "SearchDataReportDetailtGradationForAll", "toYear", toYear, "typeID", typeID, "reportTypeID", reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw DongAException.FromCommand(command, ex);
            }
        }


        /// <summary>
        /// List Report cho so sánh giai đoạn
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtForPartner> ReportDetailtPartnerGradationCompareForOne(int toYear, int typeID, string reportTypeID, string partnerID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportDetailtForPartner>();
                result = DongADatabase.ToDataAPIObject<ReportDetailtForPartner>("ReportHSDetailtPartnerLH", "SearchDataReportDetailtGradationForOne", "toYear", toYear, "typeID", typeID, "reportTypeID", reportTypeID, "partnerID", partnerID);
                return result;
            }
            catch (Exception ex)
            {
                throw DongAException.FromCommand(command, ex);
            }
        }


        /// <summary>
        /// List Report cho so sánh theo tháng
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtForPartner> ReportDetailtPartnerCompareMonthForAll(int toYear, int toMonth, string reportTypeID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportDetailtForPartner>();
                result = DongADatabase.ToDataAPIObject<ReportDetailtForPartner>("ReportHSDetailtPartnerLH", "SearchDataReportDetailtCompareMonthForAll", "toYear", toYear, "toMonth", toMonth, "reportTypeID", reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw DongAException.FromCommand(command, ex);
            }
        }


        /// <summary>
        /// List Report cho so sánh theo tháng
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtForPartner> ReportDetailtPartnerCompareMonthForOne(int toYear, int toMonth, string reportTypeID, string partnerID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportDetailtForPartner>();
                result = DongADatabase.ToDataAPIObject<ReportDetailtForPartner>("ReportHSDetailtPartnerLH", "SearchDataReportDetailtCompareMonthForOne", "toYear", toYear, "toMonth", toMonth, "reportTypeID", reportTypeID, "partnerID", partnerID);
                return result;
            }
            catch (Exception ex)
            {
                throw DongAException.FromCommand(command, ex);
            }
        }

        #endregion

        #region Báo cáo chi tiết - đối tác - loại tiền chi trả
        /// <summary>
        /// List Report detailt cho đối tác Nguyên tệ
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtForTotalMoneyType> SearchReportDetailtPartnerLTForDay(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportDetailtForTotalMoneyType>();
                result = DongADatabase.ToDataAPIObject<ReportDetailtForTotalMoneyType>("ReportHSDetailtPartnerLT", "SearchDataReportDetailtForPartnerForAll", "fromDate", fromDate, "toDate", toDate, "reportTypeID", reportTypeID);

                return result;
            }
            catch (Exception ex)
            {
                throw DongAException.FromCommand(command, ex);
            }
        }

        /// <summary>
        /// List Report detailt cho đối tác Nguyên tệ
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtForTotalMoneyType> SearchReportDetailtPartnerLTForMonth(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportDetailtForTotalMoneyType>();
                result = DongADatabase.ToDataAPIObject<ReportDetailtForTotalMoneyType>("ReportHSDetailtPartnerLT", "SearchDataReportDetailtForPartnerForAllForMonth", "fromDate", fromDate, "toDate", toDate, "reportTypeID", reportTypeID);

                return result;
            }
            catch (Exception ex)
            {
                throw DongAException.FromCommand(command, ex);
            }
        }

        /// <summary>
        /// List Report detailt cho đối tác Nguyên tệ
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtForTotalMoneyType> SearchReportDetailtPartnerLTForYear(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportDetailtForTotalMoneyType>();
                result = DongADatabase.ToDataAPIObject<ReportDetailtForTotalMoneyType>("ReportHSDetailtPartnerLT", "SearchDataReportDetailtForPartnerForAllForYear", "fromDate", fromDate, "toDate", toDate, "reportTypeID", reportTypeID);

                return result;
            }
            catch (Exception ex)
            {
                throw DongAException.FromCommand(command, ex);
            }
        }


        /// <summary>
        /// List Report
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtForTotalMoneyType> SearchPartnerLTForOne(DateTime fromDate, DateTime toDate, string reportTypeID, string partnerID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportDetailtForTotalMoneyType>();
                result = DongADatabase.ToDataAPIObject<ReportDetailtForTotalMoneyType>("ReportHSDetailtPartnerLT", "SearchDataReportDetailtForOneForDay", "fromDate", fromDate, "toDate", toDate, "reportTypeID", reportTypeID, "partnerID", partnerID);

                return result;
            }
            catch (Exception ex)
            {
                throw DongAException.FromCommand(command, ex);
            }
        }


        /// <summary>
        /// List Report
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtForTotalMoneyType> SearchPartnerLTForOneForMonth(DateTime fromDate, DateTime toDate, string reportTypeID, string partnerID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportDetailtForTotalMoneyType>();
                result = DongADatabase.ToDataAPIObject<ReportDetailtForTotalMoneyType>("ReportHSDetailtPartnerLT", "SearchDataReportDetailtForOneForMonth", "fromDate", fromDate, "toDate", toDate, "reportTypeID", reportTypeID, "partnerID", partnerID);

                return result;
            }
            catch (Exception ex)
            {
                throw DongAException.FromCommand(command, ex);
            }
        }


        /// <summary>
        /// List Report
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtForTotalMoneyType> SearchPartnerLTForOneForYear(DateTime fromDate, DateTime toDate, string reportTypeID, string partnerID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportDetailtForTotalMoneyType>();
                result = DongADatabase.ToDataAPIObject<ReportDetailtForTotalMoneyType>("ReportHSDetailtPartnerLT", "SearchDataReportDetailtForOneForYear", "fromDate", fromDate, "toDate", toDate, "reportTypeID", reportTypeID, "partnerID", partnerID);

                return result;
            }
            catch (Exception ex)
            {
                throw DongAException.FromCommand(command, ex);
            }
        }


        /// <summary>
        /// List Report cho so sánh giai đoạn
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtForTotalMoneyType> ReportDetailtPartnerLTGradationCompareForAll(int toYear, int typeID, string reportTypeID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportDetailtForTotalMoneyType>();
                result = DongADatabase.ToDataAPIObject<ReportDetailtForTotalMoneyType>("ReportHSDetailtPartnerLT", "SearchDataReportDetailtGradationForAll", "toYear", toYear, "typeID", typeID, "reportTypeID", reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw DongAException.FromCommand(command, ex);
            }
        }


        /// <summary>
        /// List Report cho so sánh giai đoạn
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtForTotalMoneyType> ReportDetailtPartnerLTGradationCompareForOne(int toYear, int typeID, string reportTypeID, string partnerID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportDetailtForTotalMoneyType>();
                result = DongADatabase.ToDataAPIObject<ReportDetailtForTotalMoneyType>("ReportHSDetailtPartnerLT", "SearchDataReportDetailtGradationForOne", "toYear", toYear, "typeID", typeID, "reportTypeID", reportTypeID, "partnerID", partnerID);
                return result;
            }
            catch (Exception ex)
            {
                throw DongAException.FromCommand(command, ex);
            }
        }


        /// <summary>
        /// List Report cho so sánh theo tháng
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtForTotalMoneyType> ReportDetailtPartnerLTCompareMonthForAll(int toYear, int toMonth, string reportTypeID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportDetailtForTotalMoneyType>();
                result = DongADatabase.ToDataAPIObject<ReportDetailtForTotalMoneyType>("ReportHSDetailtPartnerLT", "SearchDataReportDetailtCompareMonthForAll", "toYear", toYear, "toMonth", toMonth, "reportTypeID", reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw DongAException.FromCommand(command, ex);
            }
        }

        /// <summary>
        /// List Report cho so sánh theo tháng
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtForTotalMoneyType> ReportDetailtPartnerLTCompareMonthForOne(int toYear, int toMonth, string reportTypeID, string partnerID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportDetailtForTotalMoneyType>();
                result = DongADatabase.ToDataAPIObject<ReportDetailtForTotalMoneyType>("ReportHSDetailtPartnerLT", "SearchDataReportDetailtCompareMonthForOne", "toYear", toYear, "toMonth", toMonth, "reportTypeID", reportTypeID, "partnerID", partnerID);
                return result;
            }
            catch (Exception ex)
            {
                throw DongAException.FromCommand(command, ex);
            }
        }
        #endregion
    }
}

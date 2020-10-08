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
    }
}

// #################################################################
// # Copyright (C) 2010-2011, ASoft JSC.  All Rights Reserved.
// #
// # History：
// #	Date Time		Updated			Content
// #	05/10/2020		Truong Lam		Create New
// ##################################################################

using DongA.Core;
using DongA.Entities;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Oracle.ManagedDataAccess.Client;
using System.Data.Entity;
using Oracle.ManagedDataAccess.Dynamic;
using Dapper;
using Oracle.ManagedDataAccess.Types;
namespace DongA.API.DataAccess
{
    public class HSReportDAL : DongABaseDAL
    {
        /// <summary>
        /// List Report theo ngày
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<Report> SearchReportDay(DateTime fromDate, DateTime toDate, int reportTypeID)
        {
            OracleCommand command = null;
            try
            {
                var result = new List<Report>();

                using (command = DongADatabase.GetStoredProcCommandOracle("HS_TOTAL_TYPE.SEARCH_REPORT_D"))
                {
                    command.CommandType = CommandType.StoredProcedure;
                    DongADatabase.AddInOracleParameter(command, "pFromDate", OracleDbType.Date, ParameterDirection.Input, fromDate);
                    DongADatabase.AddInOracleParameter(command, "pToDate", OracleDbType.Date, ParameterDirection.Input, toDate);
                    DongADatabase.AddInOracleParameter(command, "pReportType", OracleDbType.Int32, ParameterDirection.Input, reportTypeID);

                    // Bind the parameters
                    // p1 is the RETURN REF CURSOR bound to SELECT * FROM EMP;
                    OracleParameter output = command.Parameters.Add("p_cur", OracleDbType.RefCursor);
                    output.Direction = ParameterDirection.Output;

                    using (var reader = DongADatabase.ExecuteReader(command, this))
                    {
                        result = DongADatabase.ToList<Report>(reader);
                    }
                }

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
        public List<Report> SearchReportMonth(DateTime fromDate, DateTime toDate, int reportTypeID)
        {
            OracleCommand command = null;
            try
            {
                var result = new List<Report>();

                using (command = DongADatabase.GetStoredProcCommandOracle("HS_TOTAL_TYPE.SEARCH_REPORT_M"))
                {
                    command.CommandType = CommandType.StoredProcedure;
                    DongADatabase.AddInOracleParameter(command, "pFromDate", OracleDbType.Date, ParameterDirection.Input, fromDate);
                    DongADatabase.AddInOracleParameter(command, "pToDate", OracleDbType.Date, ParameterDirection.Input, toDate);
                    DongADatabase.AddInOracleParameter(command, "pReportType", OracleDbType.Int32, ParameterDirection.Input, reportTypeID);

                    // Bind the parameters
                    // p1 is the RETURN REF CURSOR bound to SELECT * FROM EMP;
                    OracleParameter output = command.Parameters.Add("p_cur", OracleDbType.RefCursor);
                    output.Direction = ParameterDirection.Output;

                    using (var reader = DongADatabase.ExecuteReader(command, this))
                    {
                        result = DongADatabase.ToList<Report>(reader);
                    }
                }

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
        public List<Report> SearchReportYear(DateTime fromDate, DateTime toDate, int reportTypeID)
        {
            OracleCommand command = null;
            try
            {
                var result = new List<Report>();

                using (command = DongADatabase.GetStoredProcCommandOracle("HS_TOTAL_TYPE.SEARCH_REPORT_Y"))
                {
                    command.CommandType = CommandType.StoredProcedure;
                    DongADatabase.AddInOracleParameter(command, "pFromDate", OracleDbType.Date, ParameterDirection.Input, fromDate);
                    DongADatabase.AddInOracleParameter(command, "pToDate", OracleDbType.Date, ParameterDirection.Input, toDate);
                    DongADatabase.AddInOracleParameter(command, "pReportType", OracleDbType.Int32, ParameterDirection.Input, reportTypeID);

                    // Bind the parameters
                    // p1 is the RETURN REF CURSOR bound to SELECT * FROM EMP;
                    OracleParameter output = command.Parameters.Add("p_cur", OracleDbType.RefCursor);
                    output.Direction = ParameterDirection.Output;

                    using (var reader = DongADatabase.ExecuteReader(command, this))
                    {
                        result = DongADatabase.ToList<Report>(reader);
                    }
                }

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
        public List<Report> SearchReportGradationCompare(int year, int typeID, string reportTypeID)
        {
            OracleCommand command = null;
            try
            {
                var result = new List<Report>();
                using (command = DongADatabase.GetStoredProcCommandOracle("HS_TOTAL_TYPE.REPORT_GRADATION"))
                {
                    command.CommandType = CommandType.StoredProcedure;
                    DongADatabase.AddInOracleParameter(command, "pYear", OracleDbType.Int32, ParameterDirection.Input, year);
                    DongADatabase.AddInOracleParameter(command, "pGradation", OracleDbType.Int32, ParameterDirection.Input, typeID);
                    DongADatabase.AddInOracleParameter(command, "pReportType", OracleDbType.Int32, ParameterDirection.Input, reportTypeID);
                    DongADatabase.AddInOracleParameterCursor(command, "p_cur", OracleDbType.RefCursor, ParameterDirection.Output);

                    using (var reader = DongADatabase.ExecuteReader(command, this))
                    {
                        result = DongADatabase.ToList<Report>(reader);
                    }
                }
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
        public List<Report> SearchReportCompareMonth(int toYear, int toMonth, string reportTypeID)
        {
            OracleCommand command = null;
            try
            {
                var result = new List<Report>();
                using (command = DongADatabase.GetStoredProcCommandOracle("HS_TOTAL_TYPE.REPORT_COMPARE_MONTH"))
                {
                    command.CommandType = CommandType.StoredProcedure;
                    DongADatabase.AddInOracleParameter(command, "pToMonth", OracleDbType.Int32, ParameterDirection.Input, toMonth);
                    DongADatabase.AddInOracleParameter(command, "pToYear", OracleDbType.Int32, ParameterDirection.Input, toYear);
                    DongADatabase.AddInOracleParameter(command, "pReportType", OracleDbType.Int32, ParameterDirection.Input, reportTypeID);

                    DongADatabase.AddInOracleParameterCursor(command, "p_cur", OracleDbType.RefCursor, ParameterDirection.Output);

                    using (var reader = DongADatabase.ExecuteReader(command, this))
                    {
                        result = DongADatabase.ToList<Report>(reader);
                    }
                }
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
        public List<ReportForTotalMoneyType> SearchReportForDay(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            OracleCommand command = null;
            try
            {
                var result = new List<ReportForTotalMoneyType>();
                using (command = DongADatabase.GetStoredProcCommandOracle("HS_TOTAL_MONEYTYPE.SEARCH_REPORT_D"))
                {
                    command.CommandType = CommandType.StoredProcedure;
                    DongADatabase.AddInOracleParameter(command, "pFromDate", OracleDbType.Date, ParameterDirection.Input, fromDate);
                    DongADatabase.AddInOracleParameter(command, "pToDate", OracleDbType.Date, ParameterDirection.Input, toDate);
                    DongADatabase.AddInOracleParameter(command, "pReportType", OracleDbType.Int32, ParameterDirection.Input, reportTypeID);
                    DongADatabase.AddInOracleParameterCursor(command, "p_cur", OracleDbType.RefCursor, ParameterDirection.Output);

                    using (var reader = DongADatabase.ExecuteReader(command, this))
                    {
                        result = DongADatabase.ToList<ReportForTotalMoneyType>(reader);
                    }
                }
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
        public List<ReportForTotalMoneyType> SearchReportForMonth(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            OracleCommand command = null;
            try
            {
                var result = new List<ReportForTotalMoneyType>();
                using (command = DongADatabase.GetStoredProcCommandOracle("HS_TOTAL_MONEYTYPE.SEARCH_REPORT_M"))
                {
                    command.CommandType = CommandType.StoredProcedure;
                    DongADatabase.AddInOracleParameter(command, "pFromDate", OracleDbType.Date, ParameterDirection.Input, fromDate);
                    DongADatabase.AddInOracleParameter(command, "pToDate", OracleDbType.Date, ParameterDirection.Input, toDate);
                    DongADatabase.AddInOracleParameter(command, "pReportType", OracleDbType.Int32, ParameterDirection.Input, reportTypeID);
                    DongADatabase.AddInOracleParameterCursor(command, "p_cur", OracleDbType.RefCursor, ParameterDirection.Output);

                    using (var reader = DongADatabase.ExecuteReader(command, this))
                    {
                        result = DongADatabase.ToList<ReportForTotalMoneyType>(reader);
                    }
                }
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
        public List<ReportForTotalMoneyType> SearchReportForYear(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            OracleCommand command = null;
            try
            {
                var result = new List<ReportForTotalMoneyType>();
                using (command = DongADatabase.GetStoredProcCommandOracle("HS_TOTAL_MONEYTYPE.SEARCH_REPORT_Y"))
                {
                    command.CommandType = CommandType.StoredProcedure;
                    DongADatabase.AddInOracleParameter(command, "pFromDate", OracleDbType.Date, ParameterDirection.Input, fromDate);
                    DongADatabase.AddInOracleParameter(command, "pToDate", OracleDbType.Date, ParameterDirection.Input, toDate);
                    DongADatabase.AddInOracleParameter(command, "pReportType", OracleDbType.Int32, ParameterDirection.Input, reportTypeID);
                    DongADatabase.AddInOracleParameterCursor(command, "p_cur", OracleDbType.RefCursor, ParameterDirection.Output);

                    using (var reader = DongADatabase.ExecuteReader(command, this))
                    {
                        result = DongADatabase.ToList<ReportForTotalMoneyType>(reader);
                    }
                }
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
        public List<ReportForTotalMoneyType> SearchForGradationCompare(int year, int typeID, string reportTypeID)
        {
            OracleCommand command = null;
            try
            {
                var result = new List<ReportForTotalMoneyType>();
                using (command = DongADatabase.GetStoredProcCommandOracle("HS_TOTAL_MONEYTYPE.REPORT_GRADATION"))
                {
                    command.CommandType = CommandType.StoredProcedure;
                    DongADatabase.AddInOracleParameter(command, "pYear", OracleDbType.Int32, ParameterDirection.Input, year);
                    DongADatabase.AddInOracleParameter(command, "pGradation", OracleDbType.Int32, ParameterDirection.Input, typeID);
                    DongADatabase.AddInOracleParameter(command, "pReportType", OracleDbType.Int32, ParameterDirection.Input, reportTypeID);
                    DongADatabase.AddInOracleParameterCursor(command, "p_cur", OracleDbType.RefCursor, ParameterDirection.Output);

                    using (var reader = DongADatabase.ExecuteReader(command, this))
                    {
                        result = DongADatabase.ToList<ReportForTotalMoneyType>(reader);
                    }
                }
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
        public List<ReportForTotalMoneyType> SearchForCompareMonth(int ToYear, int ToMonth, string reportTypeID)
        {
            OracleCommand command = null;
            try
            {
                var result = new List<ReportForTotalMoneyType>();
                using (command = DongADatabase.GetStoredProcCommandOracle("HS_TOTAL_MONEYTYPE.REPORT_COMPARE_MONTH"))
                {
                    command.CommandType = CommandType.StoredProcedure;
                    DongADatabase.AddInOracleParameter(command, "pToMonth", OracleDbType.Int32, ParameterDirection.Input, ToMonth);
                    DongADatabase.AddInOracleParameter(command, "pToYear", OracleDbType.Int32, ParameterDirection.Input, ToYear);
                    DongADatabase.AddInOracleParameter(command, "pReportType", OracleDbType.Int32, ParameterDirection.Input, reportTypeID);

                    DongADatabase.AddInOracleParameterCursor(command, "p_cur", OracleDbType.RefCursor, ParameterDirection.Output);

                    using (var reader = DongADatabase.ExecuteReader(command, this))
                    {
                        result = DongADatabase.ToList<ReportForTotalMoneyType>(reader);
                    }
                }
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
        /// List Report theo ngày
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportForMaket> SearchReportMaketForDay(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            OracleCommand command = null;
            try
            {
                var result = new List<ReportForMaket>();

                using (command = DongADatabase.GetStoredProcCommandOracle("HS_TOTAL_MARKET.SEARCH_REPORT_D"))
                {
                    command.CommandType = CommandType.StoredProcedure;
                    DongADatabase.AddInOracleParameter(command, "pFromDate", OracleDbType.Date, ParameterDirection.Input, fromDate);
                    DongADatabase.AddInOracleParameter(command, "pToDate", OracleDbType.Date, ParameterDirection.Input, toDate);
                    DongADatabase.AddInOracleParameter(command, "pReportType", OracleDbType.Int32, ParameterDirection.Input, reportTypeID);

                    DongADatabase.AddInOracleParameterCursor(command, "p_cur", OracleDbType.RefCursor, ParameterDirection.Output);

                    using (var reader = DongADatabase.ExecuteReader(command, this))
                    {
                        result = DongADatabase.ToList<ReportForMaket>(reader);
                    }
                }

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
        public List<ReportForMaket> SearchReportMaketForMonth(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            OracleCommand command = null;
            try
            {
                var result = new List<ReportForMaket>();

                using (command = DongADatabase.GetStoredProcCommandOracle("HS_TOTAL_MARKET.SEARCH_REPORT_M"))
                {
                    command.CommandType = CommandType.StoredProcedure;
                    DongADatabase.AddInOracleParameter(command, "pFromDate", OracleDbType.Date, ParameterDirection.Input, fromDate);
                    DongADatabase.AddInOracleParameter(command, "pToDate", OracleDbType.Date, ParameterDirection.Input, toDate);
                    DongADatabase.AddInOracleParameter(command, "pReportType", OracleDbType.Int32, ParameterDirection.Input, reportTypeID);

                    DongADatabase.AddInOracleParameterCursor(command, "p_cur", OracleDbType.RefCursor, ParameterDirection.Output);

                    using (var reader = DongADatabase.ExecuteReader(command, this))
                    {
                        result = DongADatabase.ToList<ReportForMaket>(reader);
                    }
                }

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
        public List<ReportForMaket> SearchReportMaketForYear(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            OracleCommand command = null;
            try
            {
                var result = new List<ReportForMaket>();

                using (command = DongADatabase.GetStoredProcCommandOracle("HS_TOTAL_MARKET.SEARCH_REPORT_Y"))
                {
                    command.CommandType = CommandType.StoredProcedure;
                    DongADatabase.AddInOracleParameter(command, "pFromDate", OracleDbType.Date, ParameterDirection.Input, fromDate);
                    DongADatabase.AddInOracleParameter(command, "pToDate", OracleDbType.Date, ParameterDirection.Input, toDate);
                    DongADatabase.AddInOracleParameter(command, "pReportType", OracleDbType.Int32, ParameterDirection.Input, reportTypeID);

                    DongADatabase.AddInOracleParameterCursor(command, "p_cur", OracleDbType.RefCursor, ParameterDirection.Output);

                    using (var reader = DongADatabase.ExecuteReader(command, this))
                    {
                        result = DongADatabase.ToList<ReportForMaket>(reader);
                    }
                }

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
        public List<ReportForMaket> SearchReportMaketForGradationCompare(int ToYear, int typeID, string reportTypeID)
        {
            OracleCommand command = null;
            try
            {
                var result = new List<ReportForMaket>();

                using (command = DongADatabase.GetStoredProcCommandOracle("HS_TOTAL_MARKET.REPORT_GRADATION"))
                {
                    command.CommandType = CommandType.StoredProcedure;
                    DongADatabase.AddInOracleParameter(command, "pYear", OracleDbType.Int32, ParameterDirection.Input, ToYear);
                    DongADatabase.AddInOracleParameter(command, "pGradation", OracleDbType.Int32, ParameterDirection.Input, typeID);
                    DongADatabase.AddInOracleParameter(command, "pReportType", OracleDbType.Int32, ParameterDirection.Input, reportTypeID);

                    DongADatabase.AddInOracleParameterCursor(command, "p_cur", OracleDbType.RefCursor, ParameterDirection.Output);

                    using (var reader = DongADatabase.ExecuteReader(command, this))
                    {
                        result = DongADatabase.ToList<ReportForMaket>(reader);
                    }
                }

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
        public List<ReportForMaket> SearchReportMaketForCompareMonth(int ToYear, int ToMonth, string reportTypeID)
        {
            OracleCommand command = null;
            try
            {
                var result = new List<ReportForMaket>();
                using (command = DongADatabase.GetStoredProcCommandOracle("HS_TOTAL_MARKET.REPORT_COMPARE_MONTH"))
                {
                    command.CommandType = CommandType.StoredProcedure;
                    DongADatabase.AddInOracleParameter(command, "pToMonth", OracleDbType.Int32, ParameterDirection.Input, ToMonth);
                    DongADatabase.AddInOracleParameter(command, "pToYear", OracleDbType.Int32, ParameterDirection.Input, ToYear);
                    DongADatabase.AddInOracleParameter(command, "pReportType", OracleDbType.Int32, ParameterDirection.Input, reportTypeID);

                    DongADatabase.AddInOracleParameterCursor(command, "p_cur", OracleDbType.RefCursor, ParameterDirection.Output);

                    using (var reader = DongADatabase.ExecuteReader(command, this))
                    {
                        result = DongADatabase.ToList<ReportForMaket>(reader);
                    }
                }
                return result;
            }
            catch (Exception ex)
            {
                throw DongAException.FromCommand(command, ex);
            }
        }

        #endregion

        #region Get dữ liệu cho tổng hồ sơ

        /// <summary>
        /// List Report theo ngày
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportForTotalPayment> SearchReportHSTotalHSForDay(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            OracleCommand command = null;
            try
            {
                var result = new List<ReportForTotalPayment>();

                using (command = DongADatabase.GetStoredProcCommandOracle("HS_TOTAL_TOTALHS.SEARCH_REPORT_D"))
                {
                    command.CommandType = CommandType.StoredProcedure;
                    DongADatabase.AddInOracleParameter(command, "pFromDate", OracleDbType.Date, ParameterDirection.Input, fromDate);
                    DongADatabase.AddInOracleParameter(command, "pToDate", OracleDbType.Date, ParameterDirection.Input, toDate);
                    DongADatabase.AddInOracleParameter(command, "pReportType", OracleDbType.Int32, ParameterDirection.Input, reportTypeID);

                    DongADatabase.AddInOracleParameterCursor(command, "p_cur", OracleDbType.RefCursor, ParameterDirection.Output);

                    using (var reader = DongADatabase.ExecuteReader(command, this))
                    {
                        result = DongADatabase.ToList<ReportForTotalPayment>(reader);
                    }
                }

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
        public List<ReportForTotalPayment> SearchReportHSTotalHSForMonth(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            OracleCommand command = null;
            try
            {
                var result = new List<ReportForTotalPayment>();

                using (command = DongADatabase.GetStoredProcCommandOracle("HS_TOTAL_TOTALHS.SEARCH_REPORT_M"))
                {
                    command.CommandType = CommandType.StoredProcedure;
                    DongADatabase.AddInOracleParameter(command, "pFromDate", OracleDbType.Date, ParameterDirection.Input, fromDate);
                    DongADatabase.AddInOracleParameter(command, "pToDate", OracleDbType.Date, ParameterDirection.Input, toDate);
                    DongADatabase.AddInOracleParameter(command, "pReportType", OracleDbType.Int32, ParameterDirection.Input, reportTypeID);

                    DongADatabase.AddInOracleParameterCursor(command, "p_cur", OracleDbType.RefCursor, ParameterDirection.Output);

                    using (var reader = DongADatabase.ExecuteReader(command, this))
                    {
                        result = DongADatabase.ToList<ReportForTotalPayment>(reader);
                    }
                }

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
        public List<ReportForTotalPayment> SearchReportHSTotalHSForYear(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            OracleCommand command = null;
            try
            {
                var result = new List<ReportForTotalPayment>();

                using (command = DongADatabase.GetStoredProcCommandOracle("HS_TOTAL_TOTALHS.SEARCH_REPORT_Y"))
                {
                    command.CommandType = CommandType.StoredProcedure;
                    DongADatabase.AddInOracleParameter(command, "pFromDate", OracleDbType.Date, ParameterDirection.Input, fromDate);
                    DongADatabase.AddInOracleParameter(command, "pToDate", OracleDbType.Date, ParameterDirection.Input, toDate);
                    DongADatabase.AddInOracleParameter(command, "pReportType", OracleDbType.Int32, ParameterDirection.Input, reportTypeID);

                    DongADatabase.AddInOracleParameterCursor(command, "p_cur", OracleDbType.RefCursor, ParameterDirection.Output);

                    using (var reader = DongADatabase.ExecuteReader(command, this))
                    {
                        result = DongADatabase.ToList<ReportForTotalPayment>(reader);
                    }
                }

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
        public List<ReportForTotalPayment> SearchReportHSTotalHSForGradationCompare(int ToYear, int typeID, string reportTypeID)
        {
            OracleCommand command = null;
            try
            {
                var result = new List<ReportForTotalPayment>();

                using (command = DongADatabase.GetStoredProcCommandOracle("HS_TOTAL_TOTALHS.REPORT_GRADATION"))
                {
                    command.CommandType = CommandType.StoredProcedure;
                    DongADatabase.AddInOracleParameter(command, "pYear", OracleDbType.Int32, ParameterDirection.Input, ToYear);
                    DongADatabase.AddInOracleParameter(command, "pGradation", OracleDbType.Int32, ParameterDirection.Input, typeID);
                    DongADatabase.AddInOracleParameter(command, "pReportType", OracleDbType.Int32, ParameterDirection.Input, reportTypeID);

                    DongADatabase.AddInOracleParameterCursor(command, "p_cur", OracleDbType.RefCursor, ParameterDirection.Output);

                    using (var reader = DongADatabase.ExecuteReader(command, this))
                    {
                        result = DongADatabase.ToList<ReportForTotalPayment>(reader);
                    }
                }

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
        public List<ReportForTotalPayment> SearchReportHSTotalHSForCompareMonth(int ToYear, int ToMonth, string reportTypeID)
        {
            OracleCommand command = null;
            try
            {
                var result = new List<ReportForTotalPayment>();
                using (command = DongADatabase.GetStoredProcCommandOracle("HS_TOTAL_TOTALHS.REPORT_COMPARE_MONTH"))
                {
                    command.CommandType = CommandType.StoredProcedure;
                    DongADatabase.AddInOracleParameter(command, "pToMonth", OracleDbType.Int32, ParameterDirection.Input, ToMonth);
                    DongADatabase.AddInOracleParameter(command, "pToYear", OracleDbType.Int32, ParameterDirection.Input, ToYear);
                    DongADatabase.AddInOracleParameter(command, "pReportType", OracleDbType.Int32, ParameterDirection.Input, reportTypeID);

                    DongADatabase.AddInOracleParameterCursor(command, "p_cur", OracleDbType.RefCursor, ParameterDirection.Output);

                    using (var reader = DongADatabase.ExecuteReader(command, this))
                    {
                        result = DongADatabase.ToList<ReportForTotalPayment>(reader);
                    }
                }
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
        /// List Report theo ngày
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtSTMarket> SearchDataReportDetailtDay(DateTime fromDate, DateTime toDate, string reportTypeID, string marketID)
        {
            OracleCommand command = null;
            try
            {
                var result = new List<ReportDetailtSTMarket>();

                using (command = DongADatabase.GetStoredProcCommandOracle("HS_DETAILT_TYPE.SEARCH_REPORT_FOR_MARKET"))
                {
                    command.CommandType = CommandType.StoredProcedure;
                    DongADatabase.AddInOracleParameter(command, "pFromDate", OracleDbType.Date, ParameterDirection.Input, fromDate);
                    DongADatabase.AddInOracleParameter(command, "pToDate", OracleDbType.Date, ParameterDirection.Input, toDate);
                    DongADatabase.AddInOracleParameter(command, "pReportType", OracleDbType.Int32, ParameterDirection.Input, reportTypeID);
                    DongADatabase.AddInOracleParameter(command, "pParentCode", OracleDbType.Int32, ParameterDirection.Input, marketID);

                    DongADatabase.AddInOracleParameterCursor(command, "p_cur", OracleDbType.RefCursor, ParameterDirection.Output);

                    using (var reader = DongADatabase.ExecuteReader(command, this))
                    {
                        result = DongADatabase.ToList<ReportDetailtSTMarket>(reader);
                    }
                }

                return result;
            }
            catch (Exception ex)
            {
                throw DongAException.FromCommand(command, ex);
            }
        }

        /// <summary>
        /// List Report theo tháng
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtSTMarket> SearchDataReportDetailtMonth(DateTime fromDate, DateTime toDate, string reportTypeID, string marketID)
        {
            OracleCommand command = null;
            try
            {
                var result = new List<ReportDetailtSTMarket>();

                using (command = DongADatabase.GetStoredProcCommandOracle("HS_DETAILT_TYPE.SEARCH_REPORT_FOR_MARKET"))
                {
                    command.CommandType = CommandType.StoredProcedure;
                    DongADatabase.AddInOracleParameter(command, "pFromDate", OracleDbType.Date, ParameterDirection.Input, fromDate);
                    DongADatabase.AddInOracleParameter(command, "pToDate", OracleDbType.Date, ParameterDirection.Input, toDate);
                    DongADatabase.AddInOracleParameter(command, "pReportType", OracleDbType.Int32, ParameterDirection.Input, reportTypeID);
                    DongADatabase.AddInOracleParameter(command, "pParentCode", OracleDbType.Int32, ParameterDirection.Input, marketID);

                    DongADatabase.AddInOracleParameterCursor(command, "p_cur", OracleDbType.RefCursor, ParameterDirection.Output);

                    using (var reader = DongADatabase.ExecuteReader(command, this))
                    {
                        result = DongADatabase.ToList<ReportDetailtSTMarket>(reader);
                    }
                }

                return result;
            }
            catch (Exception ex)
            {
                throw DongAException.FromCommand(command, ex);
            }
        }

        /// <summary>
        /// List Report theo tháng
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtSTMarket> SearchDataReportDetailtYear(DateTime fromDate, DateTime toDate, string reportTypeID, string marketID)
        {
            OracleCommand command = null;
            try
            {
                var result = new List<ReportDetailtSTMarket>();

                using (command = DongADatabase.GetStoredProcCommandOracle("HS_DETAILT_TYPE.SEARCH_REPORT_FOR_MARKET"))
                {
                    command.CommandType = CommandType.StoredProcedure;
                    DongADatabase.AddInOracleParameter(command, "pFromDate", OracleDbType.Date, ParameterDirection.Input, fromDate);
                    DongADatabase.AddInOracleParameter(command, "pToDate", OracleDbType.Date, ParameterDirection.Input, toDate);
                    DongADatabase.AddInOracleParameter(command, "pReportType", OracleDbType.Int32, ParameterDirection.Input, reportTypeID);
                    DongADatabase.AddInOracleParameter(command, "pParentCode", OracleDbType.Int32, ParameterDirection.Input, marketID);

                    DongADatabase.AddInOracleParameterCursor(command, "p_cur", OracleDbType.RefCursor, ParameterDirection.Output);

                    using (var reader = DongADatabase.ExecuteReader(command, this))
                    {
                        result = DongADatabase.ToList<ReportDetailtSTMarket>(reader);
                    }
                }

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
        public List<ReportDetailtServiceType> SearchDataReportDetailtForOneMarketForDay(DateTime fromDate, DateTime toDate, string reportTypeID, string marketID)
        {
            OracleCommand command = null;
            try
            {
                var result = new List<ReportDetailtServiceType>();

                using (command = DongADatabase.GetStoredProcCommandOracle("HS_DETAILT_TYPE.SEARCH_REPORT_FOR_ONE_MARKET"))
                {
                    command.CommandType = CommandType.StoredProcedure;
                    DongADatabase.AddInOracleParameter(command, "pFromDate", OracleDbType.Date, ParameterDirection.Input, fromDate);
                    DongADatabase.AddInOracleParameter(command, "pToDate", OracleDbType.Date, ParameterDirection.Input, toDate);
                    DongADatabase.AddInOracleParameter(command, "pReportType", OracleDbType.Int32, ParameterDirection.Input, reportTypeID);
                    DongADatabase.AddInOracleParameter(command, "pPartnerCode", OracleDbType.Varchar2, ParameterDirection.Input, marketID);

                    DongADatabase.AddInOracleParameterCursor(command, "p_cur", OracleDbType.RefCursor, ParameterDirection.Output);

                    using (var reader = DongADatabase.ExecuteReader(command, this))
                    {
                        result = DongADatabase.ToList<ReportDetailtServiceType>(reader);
                    }
                }

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
        public List<ReportDetailtServiceType> SearchDataReportDetailtForOneMarketForMonth(DateTime fromDate, DateTime toDate, string reportTypeID, string marketID)
        {
            OracleCommand command = null;
            try
            {
                var result = new List<ReportDetailtServiceType>();

                using (command = DongADatabase.GetStoredProcCommandOracle("HS_DETAILT_TYPE.SEARCH_REPORT_FOR_ONE_MARKET"))
                {
                    command.CommandType = CommandType.StoredProcedure;
                    DongADatabase.AddInOracleParameter(command, "pFromDate", OracleDbType.Date, ParameterDirection.Input, fromDate);
                    DongADatabase.AddInOracleParameter(command, "pToDate", OracleDbType.Date, ParameterDirection.Input, toDate);
                    DongADatabase.AddInOracleParameter(command, "pReportType", OracleDbType.Int32, ParameterDirection.Input, reportTypeID);
                    DongADatabase.AddInOracleParameter(command, "pPartnerCode", OracleDbType.Varchar2, ParameterDirection.Input, marketID);

                    DongADatabase.AddInOracleParameterCursor(command, "p_cur", OracleDbType.RefCursor, ParameterDirection.Output);

                    using (var reader = DongADatabase.ExecuteReader(command, this))
                    {
                        result = DongADatabase.ToList<ReportDetailtServiceType>(reader);
                    }
                }

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
        public List<ReportDetailtServiceType> SearchDataReportDetailtForOneMarketForYear(DateTime fromDate, DateTime toDate, string reportTypeID, string marketID)
        {
            OracleCommand command = null;
            try
            {
                var result = new List<ReportDetailtServiceType>();

                using (command = DongADatabase.GetStoredProcCommandOracle("HS_DETAILT_TYPE.SEARCH_REPORT_FOR_ONE_MARKET"))
                {
                    command.CommandType = CommandType.StoredProcedure;
                    DongADatabase.AddInOracleParameter(command, "pFromDate", OracleDbType.Date, ParameterDirection.Input, fromDate);
                    DongADatabase.AddInOracleParameter(command, "pToDate", OracleDbType.Date, ParameterDirection.Input, toDate);
                    DongADatabase.AddInOracleParameter(command, "pReportType", OracleDbType.Int32, ParameterDirection.Input, reportTypeID);
                    DongADatabase.AddInOracleParameter(command, "pPartnerCode", OracleDbType.Varchar2, ParameterDirection.Input, marketID);

                    DongADatabase.AddInOracleParameterCursor(command, "p_cur", OracleDbType.RefCursor, ParameterDirection.Output);

                    using (var reader = DongADatabase.ExecuteReader(command, this))
                    {
                        result = DongADatabase.ToList<ReportDetailtServiceType>(reader);
                    }
                }

                return result;
            }
            catch (Exception ex)
            {
                throw DongAException.FromCommand(command, ex);
            }
        }


        /// <summary>
        /// List Report cho so sánh giai đoạn theo chi tiết
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtServiceType> SearchDataReportDetailtGradationForAll(int ToYear, int typeID, string reportTypeID, string marketID)
        {
            OracleCommand command = null;
            try
            {
                var result = new List<ReportDetailtServiceType>();

                using (command = DongADatabase.GetStoredProcCommandOracle("HS_DETAILT_TYPE.REPORT_GRADATION_FOR_ALL"))
                {
                    command.CommandType = CommandType.StoredProcedure;
                    DongADatabase.AddInOracleParameter(command, "pYear", OracleDbType.Int32, ParameterDirection.Input, ToYear);
                    DongADatabase.AddInOracleParameter(command, "pGradation", OracleDbType.Int32, ParameterDirection.Input, typeID);
                    DongADatabase.AddInOracleParameter(command, "pReportType", OracleDbType.Int32, ParameterDirection.Input, reportTypeID);
                    DongADatabase.AddInOracleParameter(command, "pParentCode", OracleDbType.Varchar2, ParameterDirection.Input, marketID);
                    // Cursor
                    DongADatabase.AddInOracleParameterCursor(command, "p_cur", OracleDbType.RefCursor, ParameterDirection.Output);

                    using (var reader = DongADatabase.ExecuteReader(command, this))
                    {
                        result = DongADatabase.ToList<ReportDetailtServiceType>(reader);
                    }
                }

                return result;
            }
            catch (Exception ex)
            {
                throw DongAException.FromCommand(command, ex);
            }
        }

        /// <summary>
        /// List Report cho so sánh giai đoạn theo chi tiết
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtServiceType> SearchDataReportDetailtGradationForOne(int ToYear, int typeID, string reportTypeID, string marketID)
        {
            OracleCommand command = null;
            try
            {
                var result = new List<ReportDetailtServiceType>();

                using (command = DongADatabase.GetStoredProcCommandOracle("HS_DETAILT_TYPE.REPORT_GRADATION_FOR_ONE"))
                {
                    command.CommandType = CommandType.StoredProcedure;
                    DongADatabase.AddInOracleParameter(command, "pYear", OracleDbType.Int32, ParameterDirection.Input, ToYear);
                    DongADatabase.AddInOracleParameter(command, "pGradation", OracleDbType.Int32, ParameterDirection.Input, typeID);
                    DongADatabase.AddInOracleParameter(command, "pReportType", OracleDbType.Int32, ParameterDirection.Input, reportTypeID);
                    DongADatabase.AddInOracleParameter(command, "pPartnerCode", OracleDbType.Varchar2, ParameterDirection.Input, marketID);
                    // Cursor
                    DongADatabase.AddInOracleParameterCursor(command, "p_cur", OracleDbType.RefCursor, ParameterDirection.Output);

                    using (var reader = DongADatabase.ExecuteReader(command, this))
                    {
                        result = DongADatabase.ToList<ReportDetailtServiceType>(reader);
                    }
                }

                return result;
            }
            catch (Exception ex)
            {
                throw DongAException.FromCommand(command, ex);
            }
        }

        /// <summary>
        /// List Report detailt của tất cả thị trường theo tháng
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/08/2020]
        /// </history>
        public List<ReportDetailtSTMarket> SearchDataReportDetailtCompareMonthForAll(int ToYear, int ToMonth, string reportTypeID, string marketID)
        {
            OracleCommand command = null;
            try
            {
                var result = new List<ReportDetailtSTMarket>();
                using (command = DongADatabase.GetStoredProcCommandOracle("HS_DETAILT_TYPE.COMPARE_FOR_MONTH_ALL"))
                {
                    command.CommandType = CommandType.StoredProcedure;
                    DongADatabase.AddInOracleParameter(command, "pToMonth", OracleDbType.Int32, ParameterDirection.Input, ToMonth);
                    DongADatabase.AddInOracleParameter(command, "pToYear", OracleDbType.Int32, ParameterDirection.Input, ToYear);
                    DongADatabase.AddInOracleParameter(command, "pReportType", OracleDbType.Int32, ParameterDirection.Input, reportTypeID);
                    DongADatabase.AddInOracleParameter(command, "pParentCode", OracleDbType.Varchar2, ParameterDirection.Input, marketID);
                    // Cursor
                    DongADatabase.AddInOracleParameterCursor(command, "p_cur", OracleDbType.RefCursor, ParameterDirection.Output);

                    using (var reader = DongADatabase.ExecuteReader(command, this))
                    {
                        result = DongADatabase.ToList<ReportDetailtSTMarket>(reader);
                    }
                }
                return result;
            }
            catch (Exception ex)
            {
                throw DongAException.FromCommand(command, ex);
            }
        }

        /// <summary>
        /// List Report detailt của từng trường theo tháng
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/08/2020]
        /// </history>
        public List<ReportDetailtServiceType> SearchDataReportDetailtCompareMonthForOne(int toYear, int toMonth, string reportTypeID, string marketID)
        {
            OracleCommand command = null;
            try
            {
                var result = new List<ReportDetailtServiceType>();
                using (command = DongADatabase.GetStoredProcCommandOracle("HS_DETAILT_TYPE.COMPARE_FOR_MONTH_ONE"))
                {
                    command.CommandType = CommandType.StoredProcedure;
                    DongADatabase.AddInOracleParameter(command, "pToMonth", OracleDbType.Int32, ParameterDirection.Input, toMonth);
                    DongADatabase.AddInOracleParameter(command, "pToYear", OracleDbType.Int32, ParameterDirection.Input, toYear);
                    DongADatabase.AddInOracleParameter(command, "pReportType", OracleDbType.Int32, ParameterDirection.Input, reportTypeID);
                    DongADatabase.AddInOracleParameter(command, "pPartnerCode", OracleDbType.Varchar2, ParameterDirection.Input, marketID);

                    DongADatabase.AddInOracleParameterCursor(command, "p_cur", OracleDbType.RefCursor, ParameterDirection.Output);

                    using (var reader = DongADatabase.ExecuteReader(command, this))
                    {
                        result = DongADatabase.ToList<ReportDetailtServiceType>(reader);
                    }
                }
                return result;
            }
            catch (Exception ex)
            {
                throw DongAException.FromCommand(command, ex);
            }
        }

        #endregion

        #region Get dữ liệu hồ sơ theo loại tiền


        /// <summary>
        /// List Report theo ngày
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtForTotalMoneyType> SearchDataReportDetailtHSLTForDay(DateTime fromDate, DateTime toDate, string reportTypeID, string marketID)
        {
            OracleCommand command = null;
            try
            {
                var result = new List<ReportDetailtForTotalMoneyType>();

                using (command = DongADatabase.GetStoredProcCommandOracle("HS_DETAILT_MONEYTYPE.SEARCH_REPORT_FOR_MARKET"))
                {
                    command.CommandType = CommandType.StoredProcedure;
                    DongADatabase.AddInOracleParameter(command, "pFromDate", OracleDbType.Date, ParameterDirection.Input, fromDate);
                    DongADatabase.AddInOracleParameter(command, "pToDate", OracleDbType.Date, ParameterDirection.Input, toDate);
                    DongADatabase.AddInOracleParameter(command, "pReportType", OracleDbType.Int32, ParameterDirection.Input, reportTypeID);
                    DongADatabase.AddInOracleParameter(command, "pParentCode", OracleDbType.Int32, ParameterDirection.Input, marketID);

                    DongADatabase.AddInOracleParameterCursor(command, "p_cur", OracleDbType.RefCursor, ParameterDirection.Output);

                    using (var reader = DongADatabase.ExecuteReader(command, this))
                    {
                        result = DongADatabase.ToList<ReportDetailtForTotalMoneyType>(reader);
                    }
                }

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
        public List<ReportDetailtForTotalMoneyType> SearchDataReportDetailtHSLTForOneMarketForDay(DateTime fromDate, DateTime toDate, string reportTypeID, string marketID)
        {
            OracleCommand command = null;
            try
            {
                var result = new List<ReportDetailtForTotalMoneyType>();

                using (command = DongADatabase.GetStoredProcCommandOracle("HS_DETAILT_MONEYTYPE.SEARCH_REPORT_FOR_ONE_MARKET"))
                {
                    command.CommandType = CommandType.StoredProcedure;
                    DongADatabase.AddInOracleParameter(command, "pFromDate", OracleDbType.Date, ParameterDirection.Input, fromDate);
                    DongADatabase.AddInOracleParameter(command, "pToDate", OracleDbType.Date, ParameterDirection.Input, toDate);
                    DongADatabase.AddInOracleParameter(command, "pReportType", OracleDbType.Int32, ParameterDirection.Input, reportTypeID);
                    DongADatabase.AddInOracleParameter(command, "pPartnerCode", OracleDbType.Varchar2, ParameterDirection.Input, marketID);

                    DongADatabase.AddInOracleParameterCursor(command, "p_cur", OracleDbType.RefCursor, ParameterDirection.Output);

                    using (var reader = DongADatabase.ExecuteReader(command, this))
                    {
                        result = DongADatabase.ToList<ReportDetailtForTotalMoneyType>(reader);
                    }
                }

                return result;
            }
            catch (Exception ex)
            {
                throw DongAException.FromCommand(command, ex);
            }
        }


        /// <summary>
        /// List Report cho so sánh giai đoạn theo chi tiết
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtForTotalMoneyType> SearchDataReportDetailtMTGradationForAll(int ToYear, int typeID, string reportTypeID, string marketID)
        {
            OracleCommand command = null;
            try
            {
                var result = new List<ReportDetailtForTotalMoneyType>();

                using (command = DongADatabase.GetStoredProcCommandOracle("HS_DETAILT_MONEYTYPE.REPORT_GRADATION_FOR_ALL"))
                {
                    command.CommandType = CommandType.StoredProcedure;
                    DongADatabase.AddInOracleParameter(command, "pYear", OracleDbType.Int32, ParameterDirection.Input, ToYear);
                    DongADatabase.AddInOracleParameter(command, "pGradation", OracleDbType.Int32, ParameterDirection.Input, typeID);
                    DongADatabase.AddInOracleParameter(command, "pReportType", OracleDbType.Int32, ParameterDirection.Input, reportTypeID);
                    DongADatabase.AddInOracleParameter(command, "pReportType", OracleDbType.Varchar2, ParameterDirection.Input, marketID);
                    // Cursor
                    DongADatabase.AddInOracleParameterCursor(command, "p_cur", OracleDbType.RefCursor, ParameterDirection.Output);

                    using (var reader = DongADatabase.ExecuteReader(command, this))
                    {
                        result = DongADatabase.ToList<ReportDetailtForTotalMoneyType>(reader);
                    }
                }

                return result;
            }
            catch (Exception ex)
            {
                throw DongAException.FromCommand(command, ex);
            }
        }


        /// <summary>
        /// List Report cho so sánh giai đoạn theo chi tiết
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtForTotalMoneyType> SearchDataReportDetailtMTGradationForOne(int ToYear, int typeID, string reportTypeID, string marketID)
        {
            OracleCommand command = null;
            try
            {
                var result = new List<ReportDetailtForTotalMoneyType>();

                using (command = DongADatabase.GetStoredProcCommandOracle("HS_DETAILT_MONEYTYPE.REPORT_GRADATION_FOR_ONE"))
                {
                    command.CommandType = CommandType.StoredProcedure;
                    DongADatabase.AddInOracleParameter(command, "pYear", OracleDbType.Int32, ParameterDirection.Input, ToYear);
                    DongADatabase.AddInOracleParameter(command, "pGradation", OracleDbType.Int32, ParameterDirection.Input, typeID);
                    DongADatabase.AddInOracleParameter(command, "pReportType", OracleDbType.Int32, ParameterDirection.Input, reportTypeID);
                    DongADatabase.AddInOracleParameter(command, "pPartnerCode", OracleDbType.Varchar2, ParameterDirection.Input, marketID);
                    // Cursor
                    DongADatabase.AddInOracleParameterCursor(command, "p_cur", OracleDbType.RefCursor, ParameterDirection.Output);

                    using (var reader = DongADatabase.ExecuteReader(command, this))
                    {
                        result = DongADatabase.ToList<ReportDetailtForTotalMoneyType>(reader);
                    }
                }

                return result;
            }
            catch (Exception ex)
            {
                throw DongAException.FromCommand(command, ex);
            }
        }


        /// <summary>
        /// List Report detailt của tất cả thị trường theo tháng
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/08/2020]
        /// </history>
        public List<ReportDetailtForTotalMoneyType> SearchDataReportDetailtMTCompareMonthForAll(int ToYear, int ToMonth, string reportTypeID, string marketID)
        {
            OracleCommand command = null;
            try
            {
                var result = new List<ReportDetailtForTotalMoneyType>();
                using (command = DongADatabase.GetStoredProcCommandOracle("HS_DETAILT_MONEYTYPE.COMPARE_FOR_MONTH_ALL"))
                {
                    command.CommandType = CommandType.StoredProcedure;
                    DongADatabase.AddInOracleParameter(command, "pToMonth", OracleDbType.Int32, ParameterDirection.Input, ToMonth);
                    DongADatabase.AddInOracleParameter(command, "pToYear", OracleDbType.Int32, ParameterDirection.Input, ToYear);
                    DongADatabase.AddInOracleParameter(command, "pReportType", OracleDbType.Int32, ParameterDirection.Input, reportTypeID);
                    DongADatabase.AddInOracleParameter(command, "pPartnerCode", OracleDbType.Varchar2, ParameterDirection.Input, marketID);
                    // Cursor
                    DongADatabase.AddInOracleParameterCursor(command, "p_cur", OracleDbType.RefCursor, ParameterDirection.Output);

                    using (var reader = DongADatabase.ExecuteReader(command, this))
                    {
                        result = DongADatabase.ToList<ReportDetailtForTotalMoneyType>(reader);
                    }
                }
                return result;
            }
            catch (Exception ex)
            {
                throw DongAException.FromCommand(command, ex);
            }
        }


        /// <summary>
        /// List Report detailt của từng trường theo tháng
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/08/2020]
        /// </history>
        public List<ReportDetailtForTotalMoneyType> SearchDataReportDetailtMTCompareMonthForOne(int toYear, int toMonth, string reportTypeID, string marketID)
        {
            OracleCommand command = null;
            try
            {
                var result = new List<ReportDetailtForTotalMoneyType>();
                using (command = DongADatabase.GetStoredProcCommandOracle("HS_DETAILT_MONEYTYPE.COMPARE_FOR_MONTH_ONE"))
                {
                    command.CommandType = CommandType.StoredProcedure;
                    DongADatabase.AddInOracleParameter(command, "pToMonth", OracleDbType.Int32, ParameterDirection.Input, toMonth);
                    DongADatabase.AddInOracleParameter(command, "pToYear", OracleDbType.Int32, ParameterDirection.Input, toYear);
                    DongADatabase.AddInOracleParameter(command, "pReportType", OracleDbType.Int32, ParameterDirection.Input, reportTypeID);
                    DongADatabase.AddInOracleParameter(command, "pPartnerCode", OracleDbType.Varchar2, ParameterDirection.Input, marketID);

                    DongADatabase.AddInOracleParameterCursor(command, "p_cur", OracleDbType.RefCursor, ParameterDirection.Output);

                    using (var reader = DongADatabase.ExecuteReader(command, this))
                    {
                        result = DongADatabase.ToList<ReportDetailtForTotalMoneyType>(reader);
                    }
                }
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

// #################################################################
// # Copyright (C) 2010-2011, ASoft JSC.  All Rights Reserved.
// #
// # History：
// #	Date Time		Updated			Content
// #	10/06/2020		Truong Lam		Create New
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
    public class ReportDAL : DongABaseDAL
    {
        //        private const string SQL_GetListAccount = @"
        //select * from Account
        //";
        private const string SQL_GetListAccount = @"
select PAYMENT_TYPE_CODE AS EmployeeID, CURRENCY_CODE AS Password  from  RPT_TURNOVER 
where created_date ='05-MAR-20' AND PARTNER_CODE = 20100 and VND_CNV_AMOUNT = 20263060
";
        /// <summary>
        /// List Report theo ngày
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<Account> GetListAccount()
        {
            DbCommand command = null;
            try
            {
                var result = new List<Account>();
                using (command = DongADatabase.GetSqlStringCommand(SQL_GetListAccount))
                {
                    using (var reader = DongADatabase.ExecuteReader(command, this))
                    {
                        result = DongADatabase.ToList<Account>(reader);
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
        public List<Report> DataReportDay(DateTime now, int reportTypeID)
        {
            OracleCommand command = null;
            try
            {
                var result = new List<Report>();

                using (command = DongADatabase.GetStoredProcCommandOracle("REPORT_GENERAL.REPORT_DAY"))
                {
                    command.CommandType = CommandType.StoredProcedure;
                    DongADatabase.AddInOracleParameter(command, "pDate", OracleDbType.Date, ParameterDirection.Input, now);
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
        public List<ReportDetailtServiceType> InsertTableMarket()
        {
            OracleCommand command = null;
            try
            {
                var result = new List<ReportDetailtServiceType>();

                using (command = DongADatabase.GetStoredProcCommandOracle("REPORT_ST_MARKET.INSERT_TABLE"))
                {
                    command.CommandType = CommandType.StoredProcedure;

                    // Bind the parameters
                    // p1 is the RETURN REF CURSOR bound to SELECT * FROM EMP;
                    OracleParameter output = command.Parameters.Add("p_cur", OracleDbType.RefCursor);
                    output.Direction = ParameterDirection.Output;

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
        public List<Report> SearchReportDay(DateTime fromDate, DateTime toDate, int reportTypeID)
        {
            OracleCommand command = null;
            try
            {
                var result = new List<Report>();

                using (command = DongADatabase.GetStoredProcCommandOracle("REPORT_GENERAL.SEARCH_REPORT_DAY"))
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
        /// List Report theo tháng
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<Report> GetListReportMonth(DateTime now, string reportTypeID)
        {
            OracleCommand command = null;
            try
            {
                var result = new List<Report>();

                using (command = DongADatabase.GetStoredProcCommandOracle("REPORT_GENERAL.REPORT_MONTH"))
                {
                    command.CommandType = CommandType.StoredProcedure;
                    DongADatabase.AddInOracleParameter(command, "Now", OracleDbType.Date, ParameterDirection.Input, now);
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
        public List<Report> SearchReportMonth(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            OracleCommand command = null;
            try
            {
                var result = new List<Report>();

                using (command = DongADatabase.GetStoredProcCommandOracle("REPORT_GENERAL.SEARCH_REPORT_MONTH"))
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
        /// List Report theo tháng
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<Report> GetListReportYear(DateTime now, string reportTypeID)
        {
            OracleCommand command = null;
            try
            {
                var result = new List<Report>();

                using (command = DongADatabase.GetStoredProcCommandOracle("REPORT_GENERAL.REPORT_YEAR"))
                {
                    command.CommandType = CommandType.StoredProcedure;
                    DongADatabase.AddInOracleParameter(command, "Now", OracleDbType.Date, ParameterDirection.Input, now);
                    DongADatabase.AddInOracleParameter(command, "pReportType", OracleDbType.Int32, ParameterDirection.Input, reportTypeID);
                    // Cursor
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
        /// List Report theo ngày
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<Report> SearchReportYear(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            OracleCommand command = null;
            try
            {
                var result = new List<Report>();

                using (command = DongADatabase.GetStoredProcCommandOracle("REPORT_GENERAL.SEARCH_REPORT_YEAR"))
                {
                    command.CommandType = CommandType.StoredProcedure;
                    DongADatabase.AddInOracleParameter(command, "pFromDate", OracleDbType.Date, ParameterDirection.Input, fromDate);
                    DongADatabase.AddInOracleParameter(command, "pToDate", OracleDbType.Date, ParameterDirection.Input, toDate);
                    DongADatabase.AddInOracleParameter(command, "pReportType", OracleDbType.Int32, ParameterDirection.Input, reportTypeID);
                    // Cursor
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
        /// List Report cho so sánh giai đoạn
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<Report> ListDataGradationCompare(int ToYear, int typeID, string reportTypeID)
        {
            OracleCommand command = null;
            try
            {
                var result = new List<Report>();

                using (command = DongADatabase.GetStoredProcCommandOracle("REPORT_GENERAL.REPORT_GRADATION"))
                {
                    command.CommandType = CommandType.StoredProcedure;
                    DongADatabase.AddInOracleParameter(command, "pYear", OracleDbType.Int32, ParameterDirection.Input, ToYear);
                    DongADatabase.AddInOracleParameter(command, "pGradation", OracleDbType.Int32, ParameterDirection.Input, typeID);
                    DongADatabase.AddInOracleParameter(command, "pReportType", OracleDbType.Int32, ParameterDirection.Input, reportTypeID);
                    // Cursor
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
        /// List Report cho so sánh theo tháng
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<Report> ListDataCompareMonth(int ToYear, int ToMonth, string reportTypeID)
        {
            OracleCommand command = null;
            try
            {
                var result = new List<Report>();
                using (command = DongADatabase.GetStoredProcCommandOracle("REPORT_GENERAL.REPORT_COMPARE_MONTH"))
                {
                    command.CommandType = CommandType.StoredProcedure;
                    DongADatabase.AddInOracleParameter(command, "pToMonth", OracleDbType.Int32, ParameterDirection.Input, ToMonth);
                    DongADatabase.AddInOracleParameter(command, "pToYear", OracleDbType.Int32, ParameterDirection.Input, ToYear);
                    DongADatabase.AddInOracleParameter(command, "pReportType", OracleDbType.Int32, ParameterDirection.Input, reportTypeID);
                    // Cursor
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
        /// List Report theo ngày - Nguyên tệ
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportForTotalMoneyType> DataReportForDay(DateTime now, string reportTypeID)
        {
            OracleCommand command = null;
            try
            {
                var result = new List<ReportForTotalMoneyType>();
                using (command = DongADatabase.GetStoredProcCommandOracle("REPORT_FOR_TM_TYPE.REPORT_DAY"))
                {
                    command.CommandType = CommandType.StoredProcedure;
                    DongADatabase.AddInOracleParameter(command, "pDate", OracleDbType.Date, ParameterDirection.Input, now);
                    DongADatabase.AddInOracleParameter(command, "pDate", OracleDbType.Int32, ParameterDirection.Input, reportTypeID);
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
        public List<ReportForTotalMoneyType> SearchReportForDay(DateTime fromDay, DateTime toDay, string reportTypeID)
        {
            OracleCommand command = null;
            try
            {
                var result = new List<ReportForTotalMoneyType>();
                using (command = DongADatabase.GetStoredProcCommandOracle("REPORT_FOR_TM_TYPE.SEARCH_REPORT_DAY"))
                {
                    command.CommandType = CommandType.StoredProcedure;
                    DongADatabase.AddInOracleParameter(command, "pFromDate", OracleDbType.Date, ParameterDirection.Input, fromDay);
                    DongADatabase.AddInOracleParameter(command, "pToDate", OracleDbType.Date, ParameterDirection.Input, toDay);
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
        /// List Report theo ngày - USD
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportForTotalMoneyType> DataReportForDayConvert(DateTime now, string reportTypeID)
        {
            OracleCommand command = null;
            try
            {
                var result = new List<ReportForTotalMoneyType>();
                using (command = DongADatabase.GetStoredProcCommandOracle("REPORT_FOR_TM_TYPE.REPORT_DAY_CONVERT"))
                {
                    command.CommandType = CommandType.StoredProcedure;
                    DongADatabase.AddInOracleParameter(command, "pDate", OracleDbType.Date, ParameterDirection.Input, now);
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
        /// Danh sách report theo ngày - USD
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportForTotalMoneyType> SearchReportForDayConvert(DateTime fromDay, DateTime toDay, string reportTypeID)
        {
            OracleCommand command = null;
            try
            {
                var result = new List<ReportForTotalMoneyType>();
                using (command = DongADatabase.GetStoredProcCommandOracle("REPORT_FOR_TM_TYPE.SEARCH_REPORT_DAY_CONVERT"))
                {
                    command.CommandType = CommandType.StoredProcedure;
                    DongADatabase.AddInOracleParameter(command, "pFromDate", OracleDbType.Date, ParameterDirection.Input, fromDay);
                    DongADatabase.AddInOracleParameter(command, "pToDate", OracleDbType.Date, ParameterDirection.Input, toDay);
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
        /// List Report cho tháng - Nguyên tệ
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportForTotalMoneyType> DataReportypeForMonth(DateTime now, string reportTypeID)
        {
            OracleCommand command = null;
            try
            {
                var result = new List<ReportForTotalMoneyType>();
                using (command = DongADatabase.GetStoredProcCommandOracle("REPORT_FOR_TM_TYPE.REPORT_MONTH"))
                {
                    command.CommandType = CommandType.StoredProcedure;
                    DongADatabase.AddInOracleParameter(command, "pDate", OracleDbType.Date, ParameterDirection.Input, now);
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
        /// List Report theo ngày
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportForTotalMoneyType> DataReportypeForMonthConvert(DateTime now, string reportTypeID)
        {
            OracleCommand command = null;
            try
            {
                var result = new List<ReportForTotalMoneyType>();
                using (command = DongADatabase.GetStoredProcCommandOracle("REPORT_FOR_TM_TYPE.REPORT_MONTH_CONVERT"))
                {
                    command.CommandType = CommandType.StoredProcedure;
                    DongADatabase.AddInOracleParameter(command, "pDate", OracleDbType.Date, ParameterDirection.Input, now);
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
                using (command = DongADatabase.GetStoredProcCommandOracle("REPORT_FOR_TM_TYPE.SEARCH_REPORT_MONTH"))
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
        public List<ReportForTotalMoneyType> SearchReportForMonthConvert(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            OracleCommand command = null;
            try
            {
                var result = new List<ReportForTotalMoneyType>();
                using (command = DongADatabase.GetStoredProcCommandOracle("REPORT_FOR_TM_TYPE.SEARCH_REPORT_MONTH_CONVERT"))
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
        /// List Report theo ngày
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportForTotalMoneyType> DataReportypeForYear(DateTime now, string reportTypeID)
        {
            OracleCommand command = null;
            try
            {
                var result = new List<ReportForTotalMoneyType>();
                using (command = DongADatabase.GetStoredProcCommandOracle("REPORT_FOR_TM_TYPE.REPORT_YEAR"))
                {
                    command.CommandType = CommandType.StoredProcedure;
                    DongADatabase.AddInOracleParameter(command, "pDate", OracleDbType.Date, ParameterDirection.Input, now);
                    DongADatabase.AddInOracleParameter(command, "reportTypeID", OracleDbType.Int32, ParameterDirection.Input, reportTypeID);
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
        /// List Report theo ngày
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportForTotalMoneyType> DataReportypeForYearConvert(DateTime now, string reportTypeID)
        {
            OracleCommand command = null;
            try
            {
                var result = new List<ReportForTotalMoneyType>();
                using (command = DongADatabase.GetStoredProcCommandOracle("REPORT_FOR_TM_TYPE.REPORT_YEAR_CONVERT"))
                {
                    command.CommandType = CommandType.StoredProcedure;
                    DongADatabase.AddInOracleParameter(command, "pDate", OracleDbType.Date, ParameterDirection.Input, now);
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
                using (command = DongADatabase.GetStoredProcCommandOracle("REPORT_FOR_TM_TYPE.SEARCH_REPORT_YEAR"))
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
        public List<ReportForTotalMoneyType> SearchReportForYearConvert(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            OracleCommand command = null;
            try
            {
                var result = new List<ReportForTotalMoneyType>();
                using (command = DongADatabase.GetStoredProcCommandOracle("REPORT_FOR_TM_TYPE.SEARCH_REPORT_YEAR_CONVERT"))
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
        public List<ReportForTotalMoneyType> DataReportypeForGradationCompare(int year, int typeID, string reportTypeID)
        {
            OracleCommand command = null;
            try
            {
                var result = new List<ReportForTotalMoneyType>();
                using (command = DongADatabase.GetStoredProcCommandOracle("REPORT_FOR_TM_TYPE.REPORT_GRADATION"))
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
        public List<ReportForTotalMoneyType> DataReportypeForGradationCompareConvert(int year, int typeID, string reportTypeID)
        {
            OracleCommand command = null;
            try
            {
                var result = new List<ReportForTotalMoneyType>();
                using (command = DongADatabase.GetStoredProcCommandOracle("REPORT_FOR_TM_TYPE.REPORT_GRADATION_CONVERT"))
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
        public List<ReportForTotalMoneyType> DataReportypeCompareForMonth(int ToYear, int ToMonth, string reportTypeID)
        {
            OracleCommand command = null;
            try
            {
                var result = new List<ReportForTotalMoneyType>();
                using (command = DongADatabase.GetStoredProcCommandOracle("REPORT_FOR_TM_TYPE.REPORT_COMPARE_MONTH"))
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

        /// <summary>
        /// List Report
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportForTotalMoneyType> DataReportypeCompareForMonthConvert(int ToYear, int ToMonth, string reportTypeID)
        {
            OracleCommand command = null;
            try
            {
                var result = new List<ReportForTotalMoneyType>();
                using (command = DongADatabase.GetStoredProcCommandOracle("REPORT_FOR_TM_TYPE.REPORT_COMPARE_MONTH_CONVERT"))
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

        /// <summary>
        /// List Report theo ngày
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportForTotalPayment> ReportTPForDay(DateTime now, string reportTypeID)
        {
            OracleCommand command = null;
            try
            {
                var result = new List<ReportForTotalPayment>();

                using (command = DongADatabase.GetStoredProcCommandOracle("REPORT_FOR_TOTAL_PAYMENT.REPORT_DAY"))
                {
                    command.CommandType = CommandType.StoredProcedure;
                    DongADatabase.AddInOracleParameter(command, "Now", OracleDbType.Date, ParameterDirection.Input, now);
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
        public List<ReportForTotalPayment> SearchReportTPForDay(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            OracleCommand command = null;
            try
            {
                var result = new List<ReportForTotalPayment>();

                using (command = DongADatabase.GetStoredProcCommandOracle("REPORT_FOR_TOTAL_PAYMENT.SEARCH_REPORT_DAY"))
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
        /// List Report theo tháng
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportForTotalPayment> GetListReportTPForMonth(DateTime now, string reportTypeID)
        {
            OracleCommand command = null;
            try
            {
                var result = new List<ReportForTotalPayment>();

                using (command = DongADatabase.GetStoredProcCommandOracle("REPORT_FOR_TOTAL_PAYMENT.REPORT_MONTH"))
                {
                    command.CommandType = CommandType.StoredProcedure;
                    DongADatabase.AddInOracleParameter(command, "Now", OracleDbType.Date, ParameterDirection.Input, now);
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
        public List<ReportForTotalPayment> SearchReportTPForMonth(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            OracleCommand command = null;
            try
            {
                var result = new List<ReportForTotalPayment>();

                using (command = DongADatabase.GetStoredProcCommandOracle("REPORT_FOR_TOTAL_PAYMENT.SEARCH_REPORT_MONTH"))
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
        /// List Report theo tháng
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportForTotalPayment> GetListReportTPForYear(DateTime now, string reportTypeID)
        {
            OracleCommand command = null;
            try
            {
                var result = new List<ReportForTotalPayment>();

                using (command = DongADatabase.GetStoredProcCommandOracle("REPORT_FOR_TOTAL_PAYMENT.REPORT_YEAR"))
                {
                    command.CommandType = CommandType.StoredProcedure;
                    DongADatabase.AddInOracleParameter(command, "Now", OracleDbType.Date, ParameterDirection.Input, now);
                    DongADatabase.AddInOracleParameter(command, "pReportType", OracleDbType.Int32, ParameterDirection.Input, reportTypeID);
                    // Cursor
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
        public List<ReportForTotalPayment> SearchReportTPForYear(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            OracleCommand command = null;
            try
            {
                var result = new List<ReportForTotalPayment>();

                using (command = DongADatabase.GetStoredProcCommandOracle("REPORT_FOR_TOTAL_PAYMENT.SEARCH_REPORT_YEAR"))
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
        public List<ReportForTotalPayment> ListDataTPForGradationCompare(int ToYear, int typeID, string reportTypeID)
        {
            OracleCommand command = null;
            try
            {
                var result = new List<ReportForTotalPayment>();

                using (command = DongADatabase.GetStoredProcCommandOracle("REPORT_FOR_TOTAL_PAYMENT.REPORT_GRADATION"))
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
        public List<ReportForTotalPayment> ListDataTPForCompareMonth(int ToYear, int ToMonth, string reportTypeID)
        {
            OracleCommand command = null;
            try
            {
                var result = new List<ReportForTotalPayment>();
                using (command = DongADatabase.GetStoredProcCommandOracle("REPORT_FOR_TOTAL_PAYMENT.REPORT_COMPARE_MONTH"))
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

        /// <summary>
        /// List Report theo ngày
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportForMaket> ReportMaketForDay(DateTime now, string reportTypeID)
        {
            OracleCommand command = null;
            try
            {
                var result = new List<ReportForMaket>();

                using (command = DongADatabase.GetStoredProcCommandOracle("REPORT_FOR_MARKET.REPORT_DAY"))
                {
                    command.CommandType = CommandType.StoredProcedure;
                    DongADatabase.AddInOracleParameter(command, "Now", OracleDbType.Date, ParameterDirection.Input, now);
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
        public List<ReportForMaket> SearchReportMaketForDay(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            OracleCommand command = null;
            try
            {
                var result = new List<ReportForMaket>();

                using (command = DongADatabase.GetStoredProcCommandOracle("REPORT_FOR_MARKET.SEARCH_REPORT_DAY"))
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
        /// List Report theo tháng
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportForMaket> GetListReportMarketForMonth(DateTime now, string reportTypeID)
        {
            OracleCommand command = null;
            try
            {
                var result = new List<ReportForMaket>();

                using (command = DongADatabase.GetStoredProcCommandOracle("REPORT_FOR_MARKET.REPORT_MONTH"))
                {
                    command.CommandType = CommandType.StoredProcedure;
                    DongADatabase.AddInOracleParameter(command, "Now", OracleDbType.Date, ParameterDirection.Input, now);
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

                using (command = DongADatabase.GetStoredProcCommandOracle("REPORT_FOR_MARKET.SEARCH_REPORT_MONTH"))
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
        /// List Report theo tháng
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportForMaket> GetListReportMaketForYear(DateTime now, string reportTypeID)
        {
            OracleCommand command = null;
            try
            {
                var result = new List<ReportForMaket>();

                using (command = DongADatabase.GetStoredProcCommandOracle("REPORT_FOR_MARKET.REPORT_YEAR"))
                {
                    command.CommandType = CommandType.StoredProcedure;
                    DongADatabase.AddInOracleParameter(command, "Now", OracleDbType.Date, ParameterDirection.Input, now);
                    DongADatabase.AddInOracleParameter(command, "pReportType", OracleDbType.Int32, ParameterDirection.Input, reportTypeID);
                    // Cursor
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

                using (command = DongADatabase.GetStoredProcCommandOracle("REPORT_FOR_MARKET.SEARCH_REPORT_YEAR"))
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
        public List<ReportForMaket> ListDataMaketForGradationCompare(int ToYear, int typeID, string reportTypeID)
        {
            OracleCommand command = null;
            try
            {
                var result = new List<ReportForMaket>();

                using (command = DongADatabase.GetStoredProcCommandOracle("REPORT_FOR_MARKET.REPORT_GRADATION"))
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
        public List<ReportForMaket> ListDataMaketForCompareMonth(int ToYear, int ToMonth, string reportTypeID)
        {
            OracleCommand command = null;
            try
            {
                var result = new List<ReportForMaket>();
                using (command = DongADatabase.GetStoredProcCommandOracle("REPORT_FOR_MARKET.REPORT_COMPARE_MONTH"))
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

        /// <summary>
        /// List Report theo ngày
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportForMaket> CreateDataMarket()
        {
            OracleCommand command = null;
            try
            {
                List<ReportForMaket> result = new List<ReportForMaket>();

                using (command = DongADatabase.GetStoredProcCommandOracle("REPORT_FOR_MARKET.CREATE_TABLE_TEMP"))
                {
                    command.CommandType = CommandType.StoredProcedure;
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
        /// List Report deatailt cho ngày
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/08/2020]
        /// </history>
        public List<ReportDetailtServiceType> DataReportDetailtDay(DateTime now, int reportTypeID)
        {
            OracleCommand command = null;
            try
            {
                var result = new List<ReportDetailtServiceType>();

                using (command = DongADatabase.GetStoredProcCommandOracle("REPORT_SERVICE_TYPE.REPORT_DAY"))
                {
                    command.CommandType = CommandType.StoredProcedure;
                    DongADatabase.AddInOracleParameter(command, "pDate", OracleDbType.Date, ParameterDirection.Input, now);
                    DongADatabase.AddInOracleParameter(command, "pReportType", OracleDbType.Int32, ParameterDirection.Input, reportTypeID);

                    // Bind the parameters
                    // p1 is the RETURN REF CURSOR bound to SELECT * FROM EMP;
                    OracleParameter output = command.Parameters.Add("p_cur", OracleDbType.RefCursor);
                    output.Direction = ParameterDirection.Output;

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
        public List<ReportDetailtServiceType> SearchReportDetailtDay(DateTime fromDate, DateTime toDate, int reportTypeID)
        {
            OracleCommand command = null;
            try
            {
                var result = new List<ReportDetailtServiceType>();

                using (command = DongADatabase.GetStoredProcCommandOracle("REPORT_SERVICE_TYPE.SEARCH_REPORT_DAY"))
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
        /// List Report theo tháng
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtServiceType> GetListReportDetailtMonth(DateTime now, string reportTypeID)
        {
            OracleCommand command = null;
            try
            {
                var result = new List<ReportDetailtServiceType>();

                using (command = DongADatabase.GetStoredProcCommandOracle("REPORT_SERVICE_TYPE.MARKET_FOR_PARTNER"))
                {
                    command.CommandType = CommandType.StoredProcedure;
                    DongADatabase.AddInOracleParameter(command, "Now", OracleDbType.Date, ParameterDirection.Input, now);
                    DongADatabase.AddInOracleParameter(command, "pReportType", OracleDbType.Int32, ParameterDirection.Input, reportTypeID);

                    // Bind the parameters
                    // p1 is the RETURN REF CURSOR bound to SELECT * FROM EMP;
                    OracleParameter output = command.Parameters.Add("p_cur", OracleDbType.RefCursor);
                    output.Direction = ParameterDirection.Output;

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
        public List<ReportDetailtServiceType> SearchReportDetailtMonth(DateTime fromDate, DateTime toDate, string partnerCode, string reportTypeID)
        {
            OracleCommand command = null;
            try
            {
                var result = new List<ReportDetailtServiceType>();

                using (command = DongADatabase.GetStoredProcCommandOracle("REPORT_SERVICE_TYPE.SEARCH_REPORT_MONTH"))
                {
                    command.CommandType = CommandType.StoredProcedure;
                    DongADatabase.AddInOracleParameter(command, "pFromDate", OracleDbType.Date, ParameterDirection.Input, fromDate);
                    DongADatabase.AddInOracleParameter(command, "pToDate", OracleDbType.Date, ParameterDirection.Input, toDate);
                    DongADatabase.AddInOracleParameter(command, "pReportType", OracleDbType.Int32, ParameterDirection.Input, reportTypeID);
                    DongADatabase.AddInOracleParameter(command, "pPartnerCode", OracleDbType.Varchar2, ParameterDirection.Input, partnerCode);

                    // Bind the parameters
                    // p1 is the RETURN REF CURSOR bound to SELECT * FROM EMP;
                    OracleParameter output = command.Parameters.Add("p_cur", OracleDbType.RefCursor);
                    output.Direction = ParameterDirection.Output;

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
        /// List Report theo tháng
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtServiceType> GetListReportDetailtYear(DateTime now, string reportTypeID)
        {
            OracleCommand command = null;
            try
            {
                var result = new List<ReportDetailtServiceType>();

                using (command = DongADatabase.GetStoredProcCommandOracle("REPORT_SERVICE_TYPE.REPORT_YEAR"))
                {
                    command.CommandType = CommandType.StoredProcedure;
                    DongADatabase.AddInOracleParameter(command, "Now", OracleDbType.Date, ParameterDirection.Input, now);
                    DongADatabase.AddInOracleParameter(command, "pReportType", OracleDbType.Int32, ParameterDirection.Input, reportTypeID);
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
        /// List Report theo ngày
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtServiceType> SearchReportDetailtYear(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            OracleCommand command = null;
            try
            {
                var result = new List<ReportDetailtServiceType>();

                using (command = DongADatabase.GetStoredProcCommandOracle("REPORT_SERVICE_TYPE.SEARCH_REPORT_YEAR"))
                {
                    command.CommandType = CommandType.StoredProcedure;
                    DongADatabase.AddInOracleParameter(command, "pFromDate", OracleDbType.Date, ParameterDirection.Input, fromDate);
                    DongADatabase.AddInOracleParameter(command, "pToDate", OracleDbType.Date, ParameterDirection.Input, toDate);
                    DongADatabase.AddInOracleParameter(command, "pReportType", OracleDbType.Int32, ParameterDirection.Input, reportTypeID);
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
        /// List Report cho so sánh giai đoạn
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtServiceType> ListDataDetailtGradationCompare(int ToYear, int typeID, string reportTypeID)
        {
            OracleCommand command = null;
            try
            {
                var result = new List<ReportDetailtServiceType>();

                using (command = DongADatabase.GetStoredProcCommandOracle("REPORT_SERVICE_TYPE.REPORT_GRADATION"))
                {
                    command.CommandType = CommandType.StoredProcedure;
                    DongADatabase.AddInOracleParameter(command, "pYear", OracleDbType.Int32, ParameterDirection.Input, ToYear);
                    DongADatabase.AddInOracleParameter(command, "pGradation", OracleDbType.Int32, ParameterDirection.Input, typeID);
                    DongADatabase.AddInOracleParameter(command, "pReportType", OracleDbType.Int32, ParameterDirection.Input, reportTypeID);
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
        /// List Report cho so sánh theo tháng
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtServiceType> ListDataDetailtCompareMonth(int ToYear, int ToMonth, string reportTypeID)
        {
            OracleCommand command = null;
            try
            {
                var result = new List<ReportDetailtServiceType>();
                using (command = DongADatabase.GetStoredProcCommandOracle("REPORT_SERVICE_TYPE.REPORT_COMPARE_MONTH"))
                {
                    command.CommandType = CommandType.StoredProcedure;
                    DongADatabase.AddInOracleParameter(command, "pToMonth", OracleDbType.Int32, ParameterDirection.Input, ToMonth);
                    DongADatabase.AddInOracleParameter(command, "pToYear", OracleDbType.Int32, ParameterDirection.Input, ToYear);
                    DongADatabase.AddInOracleParameter(command, "pReportType", OracleDbType.Int32, ParameterDirection.Input, reportTypeID);
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
        /// List Report deatailt cho ngày
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/08/2020]
        /// </history>
        public List<Partner> ListPartNer()
        {
            OracleCommand command = null;
            try
            {
                var result = new List<Partner>();

                using (command = DongADatabase.GetStoredProcCommandOracle("REPORT_SERVICE_TYPE.LIST_PARTNER"))
                {
                    command.CommandType = CommandType.StoredProcedure;
                    // Bind the parameters
                    // p1 is the RETURN REF CURSOR bound to SELECT * FROM EMP;
                    OracleParameter output = command.Parameters.Add("p_cur", OracleDbType.RefCursor);
                    output.Direction = ParameterDirection.Output;

                    using (var reader = DongADatabase.ExecuteReader(command, this))
                    {
                        result = DongADatabase.ToList<Partner>(reader);
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
        /// List Report deatailt cho ngày
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/08/2020]
        /// </history>
        public List<Market> ListMarket()
        {
            OracleCommand command = null;
            try
            {
                var result = new List<Market>();

                using (command = DongADatabase.GetStoredProcCommandOracle("REPORT_ST_MARKET.LIST_MARKET"))
                {
                    command.CommandType = CommandType.StoredProcedure;
                    // Bind the parameters
                    // p1 is the RETURN REF CURSOR bound to SELECT * FROM EMP;
                    OracleParameter output = command.Parameters.Add("p_cur", OracleDbType.RefCursor);
                    output.Direction = ParameterDirection.Output;

                    using (var reader = DongADatabase.ExecuteReader(command, this))
                    {
                        result = DongADatabase.ToList<Market>(reader);
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
        /// List Report deatailt cho ngày
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/08/2020]
        /// </history>
        public List<ReportDetailtSTMarket> DataReportDetailtMarKetForDay(DateTime now, int reportTypeID)
        {
            OracleCommand command = null;
            try
            {
                var result = new List<ReportDetailtSTMarket>();

                using (command = DongADatabase.GetStoredProcCommandOracle("REPORT_ST_MARKET.REPORT_FOR_MARKET"))
                {
                    command.CommandType = CommandType.StoredProcedure;
                    DongADatabase.AddInOracleParameter(command, "pDate", OracleDbType.Date, ParameterDirection.Input, now);
                    DongADatabase.AddInOracleParameter(command, "pReportType", OracleDbType.Int32, ParameterDirection.Input, reportTypeID);

                    // Bind the parameters
                    // p1 is the RETURN REF CURSOR bound to SELECT * FROM EMP;
                    OracleParameter output = command.Parameters.Add("p_cur", OracleDbType.RefCursor);
                    output.Direction = ParameterDirection.Output;

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
        public List<ReportDetailtSTMarket> SearchDataReportDetailtDay(DateTime fromDate, DateTime toDate, string reportTypeID, string marketID)
        {
            OracleCommand command = null;
            try
            {
                var result = new List<ReportDetailtSTMarket>();

                using (command = DongADatabase.GetStoredProcCommandOracle("REPORT_ST_MARKET.SEARCH_REPORT_FOR_MARKET"))
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
        /// List Report deatailt cho ngày
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/08/2020]
        /// </history>
        public List<ReportDetailtServiceType> DataReportDetailtForOneMarket(DateTime now, int reportTypeID)
        {
            OracleCommand command = null;
            try
            {
                var result = new List<ReportDetailtServiceType>();

                using (command = DongADatabase.GetStoredProcCommandOracle("REPORT_ST_MARKET.REPORT_FOR_ONE_MARKET"))
                {
                    command.CommandType = CommandType.StoredProcedure;
                    DongADatabase.AddInOracleParameter(command, "pDate", OracleDbType.Date, ParameterDirection.Input, now);
                    DongADatabase.AddInOracleParameter(command, "pReportType", OracleDbType.Int32, ParameterDirection.Input, reportTypeID);

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
        public List<ReportDetailtServiceType> SearchDataReportDetailtForOneMarket(DateTime fromDate, DateTime toDate, string reportTypeID, string marketID)
        {
            OracleCommand command = null;
            try
            {
                var result = new List<ReportDetailtServiceType>();

                using (command = DongADatabase.GetStoredProcCommandOracle("REPORT_ST_MARKET.SEARCH_REPORT_FOR_ONE_MARKET"))
                {
                    command.CommandType = CommandType.StoredProcedure;
                    DongADatabase.AddInOracleParameter(command, "pFromDate", OracleDbType.Date, ParameterDirection.Input, fromDate);
                    DongADatabase.AddInOracleParameter(command, "pToDate", OracleDbType.Date, ParameterDirection.Input, toDate);
                    DongADatabase.AddInOracleParameter(command, "pReportType", OracleDbType.Int32, ParameterDirection.Input, reportTypeID);
                    DongADatabase.AddInOracleParameter(command, "pPartnerCode", OracleDbType.Int32, ParameterDirection.Input, marketID);

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
        public List<ReportDetailtServiceType> SearchDataReportDetailtGradationForAll(int ToYear, int typeID, string reportTypeID)
        {
            OracleCommand command = null;
            try
            {
                var result = new List<ReportDetailtServiceType>();

                using (command = DongADatabase.GetStoredProcCommandOracle("REPORT_ST_MARKET.REPORT_GRADATION_FOR_ALL"))
                {
                    command.CommandType = CommandType.StoredProcedure;
                    DongADatabase.AddInOracleParameter(command, "pYear", OracleDbType.Int32, ParameterDirection.Input, ToYear);
                    DongADatabase.AddInOracleParameter(command, "pGradation", OracleDbType.Int32, ParameterDirection.Input, typeID);
                    DongADatabase.AddInOracleParameter(command, "pReportType", OracleDbType.Int32, ParameterDirection.Input, reportTypeID);
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

                using (command = DongADatabase.GetStoredProcCommandOracle("REPORT_ST_MARKET.REPORT_GRADATION_FOR_ONE"))
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
        public List<ReportDetailtSTMarket> SearchDataReportDetailtCompareMonthForAll(int ToYear, int ToMonth, string reportTypeID)
        {
            OracleCommand command = null;
            try
            {
                var result = new List<ReportDetailtSTMarket>();
                using (command = DongADatabase.GetStoredProcCommandOracle("REPORT_ST_MARKET.COMPARE_FOR_MONTH_ALL"))
                {
                    command.CommandType = CommandType.StoredProcedure;
                    DongADatabase.AddInOracleParameter(command, "pToMonth", OracleDbType.Int32, ParameterDirection.Input, ToMonth);
                    DongADatabase.AddInOracleParameter(command, "pToYear", OracleDbType.Int32, ParameterDirection.Input, ToYear);
                    DongADatabase.AddInOracleParameter(command, "pReportType", OracleDbType.Int32, ParameterDirection.Input, reportTypeID);
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
                using (command = DongADatabase.GetStoredProcCommandOracle("REPORT_ST_MARKET.COMPARE_FOR_MONTH_ONE"))
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
    }
}

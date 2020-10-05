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
    }
}

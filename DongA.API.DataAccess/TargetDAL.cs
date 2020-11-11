// #################################################################
// # Copyright (C) 2010-2011, ASoft JSC.  All Rights Reserved.
// #
// # History：
// #	Date Time		Updated			Content
// #	10/11/2020		Truong Lam		Create New
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
    public class TargetDAL : DongABaseDAL
    {
        /// <summary>
        /// List Report theo ngày
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public bool InsertdataTarget(FormTarget datataget)
        {
            OracleCommand command = null;
            try
            {
                //var result = new List<FormTarget>();

                using (command = DongADatabase.GetStoredProcCommandOracle("INSERT_TARGET.INSERT_TABLE"))
                {
                    command.CommandType = CommandType.StoredProcedure;
                    DongADatabase.AddInOracleParameter(command, "ObjectID", OracleDbType.NVarchar2, ParameterDirection.Input, datataget.ObjectID);
                    DongADatabase.AddInOracleParameter(command, "ObjectName", OracleDbType.NVarchar2, ParameterDirection.Input, datataget.ObjectName);
                    DongADatabase.AddInOracleParameter(command, "TargetValue", OracleDbType.Double, ParameterDirection.Input, datataget.TargetValue);
                    DongADatabase.AddInOracleParameter(command, "COL1", OracleDbType.Double, ParameterDirection.Input, datataget.COL1);
                    DongADatabase.AddInOracleParameter(command, "COL2", OracleDbType.Double, ParameterDirection.Input, datataget.COL2);
                    DongADatabase.AddInOracleParameter(command, "COL3", OracleDbType.Double, ParameterDirection.Input, datataget.COL3);
                    DongADatabase.AddInOracleParameter(command, "COL4", OracleDbType.Double, ParameterDirection.Input, datataget.COL4);
                    DongADatabase.AddInOracleParameter(command, "COL5", OracleDbType.Double, ParameterDirection.Input, datataget.COL5);
                    DongADatabase.AddInOracleParameter(command, "COL6", OracleDbType.Double, ParameterDirection.Input, datataget.COL6);
                    DongADatabase.AddInOracleParameter(command, "COL7", OracleDbType.Double, ParameterDirection.Input, datataget.COL7);
                    DongADatabase.AddInOracleParameter(command, "COL8", OracleDbType.Double, ParameterDirection.Input, datataget.COL8);
                    DongADatabase.AddInOracleParameter(command, "COL9", OracleDbType.Double, ParameterDirection.Input, datataget.COL9);
                    DongADatabase.AddInOracleParameter(command, "COL10", OracleDbType.Double, ParameterDirection.Input, datataget.COL10);
                    DongADatabase.AddInOracleParameter(command, "COL11", OracleDbType.Double, ParameterDirection.Input, datataget.COL11);
                    DongADatabase.AddInOracleParameter(command, "COL12", OracleDbType.Double, ParameterDirection.Input, datataget.COL12);
                    DongADatabase.AddInOracleParameter(command, "CreatedDate", OracleDbType.Date, ParameterDirection.Input, datataget.CreatedDate);
                    DongADatabase.AddInOracleParameter(command, "CustomDate", OracleDbType.Date, ParameterDirection.Input, datataget.CustomDate);
                    
                    if (DongADatabase.ExecuteNonQuery(command, this) <= 0)
                    {
                        return true;
                    }
                }
                return false;

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
        public bool UpdatedataTarget(FormTarget datataget)
        {
            OracleCommand command = null;
            try
            {
                //var result = new List<FormTarget>();

                using (command = DongADatabase.GetStoredProcCommandOracle("INSERT_TARGET.UPDATE_TABLE"))
                {
                    command.CommandType = CommandType.StoredProcedure;
                    DongADatabase.AddInOracleParameter(command, "ObjectID", OracleDbType.NVarchar2, ParameterDirection.Input, datataget.ObjectID);
                    DongADatabase.AddInOracleParameter(command, "ObjectName", OracleDbType.NVarchar2, ParameterDirection.Input, datataget.ObjectName);
                    DongADatabase.AddInOracleParameter(command, "TargetValue", OracleDbType.Double, ParameterDirection.Input, datataget.TargetValue);
                    DongADatabase.AddInOracleParameter(command, "COLOUMN1", OracleDbType.Double, ParameterDirection.Input, datataget.COL1);
                    DongADatabase.AddInOracleParameter(command, "COLOUMN2", OracleDbType.Double, ParameterDirection.Input, datataget.COL2);
                    DongADatabase.AddInOracleParameter(command, "COLOUMN3", OracleDbType.Double, ParameterDirection.Input, datataget.COL3);
                    DongADatabase.AddInOracleParameter(command, "COLOUMN4", OracleDbType.Double, ParameterDirection.Input, datataget.COL4);
                    DongADatabase.AddInOracleParameter(command, "COLOUMN5", OracleDbType.Double, ParameterDirection.Input, datataget.COL5);
                    DongADatabase.AddInOracleParameter(command, "COLOUMN6", OracleDbType.Double, ParameterDirection.Input, datataget.COL6);
                    DongADatabase.AddInOracleParameter(command, "COLOUMN7", OracleDbType.Double, ParameterDirection.Input, datataget.COL7);
                    DongADatabase.AddInOracleParameter(command, "COLOUMN8", OracleDbType.Double, ParameterDirection.Input, datataget.COL8);
                    DongADatabase.AddInOracleParameter(command, "COLOUMN9", OracleDbType.Double, ParameterDirection.Input, datataget.COL9);
                    DongADatabase.AddInOracleParameter(command, "COLOUMN10", OracleDbType.Double, ParameterDirection.Input, datataget.COL10);
                    DongADatabase.AddInOracleParameter(command, "COLOUMN11", OracleDbType.Double, ParameterDirection.Input, datataget.COL11);
                    DongADatabase.AddInOracleParameter(command, "COLOUMN12", OracleDbType.Double, ParameterDirection.Input, datataget.COL12);
                    DongADatabase.AddInOracleParameter(command, "CreatedDate", OracleDbType.Date, ParameterDirection.Input, datataget.CreatedDate);
                    DongADatabase.AddInOracleParameter(command, "CustomDate", OracleDbType.Date, ParameterDirection.Input, datataget.CustomDate);

                    if (DongADatabase.ExecuteNonQuery(command, this) <= 0)
                    {
                        return true;
                    }
                }
                return false;

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
        public bool DeleteDataTarget(FormTarget datataget)
        {
            OracleCommand command = null;
            try
            {
                //var result = new List<FormTarget>();

                using (command = DongADatabase.GetStoredProcCommandOracle("INSERT_TARGET.DELETE_TABLE"))
                {
                    command.CommandType = CommandType.StoredProcedure;
                    DongADatabase.AddInOracleParameter(command, "ObjectID", OracleDbType.NVarchar2, ParameterDirection.Input, datataget.ObjectID);
                    DongADatabase.AddInOracleParameter(command, "ObjectName", OracleDbType.NVarchar2, ParameterDirection.Input, datataget.ObjectName);
                    DongADatabase.AddInOracleParameter(command, "TargetValue", OracleDbType.Double, ParameterDirection.Input, datataget.TargetValue);
                    DongADatabase.AddInOracleParameter(command, "COLOUMN1", OracleDbType.Double, ParameterDirection.Input, datataget.COL1);
                    DongADatabase.AddInOracleParameter(command, "COLOUMN2", OracleDbType.Double, ParameterDirection.Input, datataget.COL2);
                    DongADatabase.AddInOracleParameter(command, "COLOUMN3", OracleDbType.Double, ParameterDirection.Input, datataget.COL3);
                    DongADatabase.AddInOracleParameter(command, "COLOUMN4", OracleDbType.Double, ParameterDirection.Input, datataget.COL4);
                    DongADatabase.AddInOracleParameter(command, "COLOUMN5", OracleDbType.Double, ParameterDirection.Input, datataget.COL5);
                    DongADatabase.AddInOracleParameter(command, "COLOUMN6", OracleDbType.Double, ParameterDirection.Input, datataget.COL6);
                    DongADatabase.AddInOracleParameter(command, "COLOUMN7", OracleDbType.Double, ParameterDirection.Input, datataget.COL7);
                    DongADatabase.AddInOracleParameter(command, "COLOUMN8", OracleDbType.Double, ParameterDirection.Input, datataget.COL8);
                    DongADatabase.AddInOracleParameter(command, "COLOUMN9", OracleDbType.Double, ParameterDirection.Input, datataget.COL9);
                    DongADatabase.AddInOracleParameter(command, "COLOUMN10", OracleDbType.Double, ParameterDirection.Input, datataget.COL10);
                    DongADatabase.AddInOracleParameter(command, "COLOUMN11", OracleDbType.Double, ParameterDirection.Input, datataget.COL11);
                    DongADatabase.AddInOracleParameter(command, "COLOUMN12", OracleDbType.Double, ParameterDirection.Input, datataget.COL12);
                    DongADatabase.AddInOracleParameter(command, "CreatedDate", OracleDbType.Date, ParameterDirection.Input, datataget.CreatedDate);
                    DongADatabase.AddInOracleParameter(command, "CustomDate", OracleDbType.Date, ParameterDirection.Input, datataget.CustomDate);

                    if (DongADatabase.ExecuteNonQuery(command, this) <= 0)
                    {
                        return true;
                    }
                }
                return false;

            }
            catch (Exception ex)
            {
                throw DongAException.FromCommand(command, ex);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/11/2020]
        /// </history>
        public List<FormTarget> ListTarget()
        {
            OracleCommand command = null;
            try
            {
                var result = new List<FormTarget>();

                using (command = DongADatabase.GetStoredProcCommandOracle("INSERT_TARGET.LIST_TARGET"))
                {
                    command.CommandType = CommandType.StoredProcedure;
                    // Cursor
                    DongADatabase.AddInOracleParameterCursor(command, "p_cur", OracleDbType.RefCursor, ParameterDirection.Output);

                    using (var reader = DongADatabase.ExecuteReader(command, this))
                    {
                        result = DongADatabase.ToList<FormTarget>(reader);
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

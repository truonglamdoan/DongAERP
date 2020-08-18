// #################################################################
// # Copyright (C) 2010-2011, ASoft JSC.  All Rights Reserved.
// #
// # History：
// #	Date Time		Updated			Content
// #	10/06/2020		Truong Lam		Create New
// ##################################################################

using DongA.Core;
using Models.Framework;
using System;
using System.Collections.Generic;
using System.Data.Common;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DongA.DataAccess
{
    public class MarKetDetailDAL : DongABaseDAL
    {
        public MarKetDetailDAL()
        {
        }

        public MarKetDetailDAL(DbTransaction tran) : base(tran)
        {
        }

        private const string SQL_All = @"
select * from MarketDetail
";
        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<MarKetDetail> GetAll()
        {
            DbCommand command = null;
            try
            {
                var result = new List<MarKetDetail>();
                using (command = DongADatabase.GetSqlStringCommand(SQL_All))
                {

                    using (var reader = DongADatabase.ExecuteReader(command, this))
                    {
                        result = DongADatabase.ToList<MarKetDetail>(reader);
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

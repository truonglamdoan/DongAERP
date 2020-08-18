// #################################################################
// # Copyright (C) 2010-2011, ASoft JSC.  All Rights Reserved.
// #
// # History：
// #	Date Time		Updated			Content
// #	10/06/2020		Truong Lam		Create New
// ##################################################################

using DongA.Core;
using DongA.Entities;
using Models;
using Models.Framework;
using System;
using System.Collections.Generic;
using System.Data.Common;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DongA.DataAccess
{
    public class PartnerDAL : DongABaseDAL
    {
        public PartnerDAL()
        {
        }

        public PartnerDAL(DbTransaction tran) : base(tran)
        {
        }

        private const string SQL_All = @"
select * from Partner
";
        /// <summary>
        /// List Partner
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<Partner> GetListPartner()
        {
            DbCommand command = null;
            try
            {
                var result = new List<Partner>();
                using (command = DongADatabase.GetSqlStringCommand(SQL_All))
                {

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

    }
}

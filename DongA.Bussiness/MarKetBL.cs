﻿// #################################################################
// # Copyright (C) 2010-2011, ASoft JSC.  All Rights Reserved.
// #
// # History：
// #	Date Time		Updated			Content
// #	10/06/2020		Truong Lam		Create New
// ##################################################################

using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DongA.Core;
using Models.Framework;
using DongA.DataAccess;
using static DongA.Core.DongAException;

namespace DongA.Bussiness
{
    public class MarKetBL : DongABaseDAL
    {
        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<MarKetDetail> GetAll()
        {
            try
            {
                MarKetDetailDAL dal = new MarKetDetailDAL();
                List<MarKetDetail> result = dal.GetAll();
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }

        }
    }
}

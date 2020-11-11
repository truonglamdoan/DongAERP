// #################################################################
// # Copyright (C) 2010-2011, ASoft JSC.  All Rights Reserved.
// #
// # History：
// #	Date Time		Updated			Content
// #	09/11/2020		Truong Lam		Create New
// ##################################################################

using DongA.API.Bussiness;
using DongA.Core;
using DongA.Entities;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using static DongA.Core.DongAException;



namespace DongA.API.Controllers
{
    public class TargetController : ApiController
    {
        /// <summary>
        ///
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/11/2020]
        /// </history>
        [HttpGet]
        public List<FormTarget> ListTarget()
        {
            try
            {
                List<FormTarget> listReport = new TargetBL().ListTarget();
                return listReport;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        ///
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/11/2020]
        /// </history>
        [HttpPost]
        public bool InsertdataTarget(FormTarget data)
        {
            try
            {
                //FormTarget dataItem = JsonConvert.DeserializeObject<FormTarget>(data);
                return new TargetBL().InsertdataTarget(data);
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        ///
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/11/2020]
        /// </history>
        [HttpPost]
        public bool UpdatedataTarget(FormTarget data)
        {
            try
            {
                return new TargetBL().UpdatedataTarget(data);
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        ///
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/11/2020]
        /// </history>
        [HttpPost]
        public bool DeleteDataTarget(FormTarget data)
        {
            try
            {
                return new TargetBL().DeleteDataTarget(data);
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }
    }
}

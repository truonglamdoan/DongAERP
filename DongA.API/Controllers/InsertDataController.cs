using DongA.API.Bussiness;
using DongA.Core;
using DongA.Entities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using static DongA.Core.DongAException;

namespace DongA.API.Controllers
{
    public class InsertDataController : ApiController
    {

        /// <summary>
        /// Insert dữ liệu vào bảng
        /// </summary>
        /// <param name="reportTypeID"></param>
        /// <returns></returns>
        [HttpGet]
        public List<ReportDetailtServiceType> InsertTableMarket()
        {
            try
            {
                DateTime now = DateTime.Now;
                List<ReportDetailtServiceType> listReport = new ReportBL().InsertTableMarket();
                return listReport;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }
    }
}

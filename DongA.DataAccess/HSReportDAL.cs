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
    }
}

// #################################################################
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
using Models;
using DongA.Entities;

namespace DongA.Bussiness
{
    public class HSReportBL : DongABaseDAL
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
            try
            {
                HSReportDAL dal = new HSReportDAL();
                List<Report> result = dal.SearchDay(fromDate, toDate, reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
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
            try
            {
                HSReportDAL dal = new HSReportDAL();
                List<Report> result = dal.SearchMonth(fromDate, toDate, reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
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
            try
            {
                HSReportDAL dal = new HSReportDAL();
                List<Report> result = dal.SearchYear(fromDate, toDate, reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }

        }

        /// <summary>
        /// List Report cho so sánh giai đoạn
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<Report> SearchGradationCompare(int ToYear, int typeID, string reportTypeID)
        {
            try
            {
                HSReportDAL dal = new HSReportDAL();
                List<Report> result = dal.SearchGradationCompare(ToYear, typeID, reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// List data hiển thị phần trăm cho so sánh giai đoạn
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<Report> ListDataGradationComparePercent(int typeID, int ToYear, string reportTypeID)
        {
            try
            {
                HSReportDAL dal = new HSReportDAL();
                List<Report> result = dal.SearchGradationCompare(ToYear, typeID, reportTypeID);
                List<Report> DataConvertPercent = new List<Report>();
                // Xử lý việc chuyển đổi thành %
                foreach (Report item in result)
                {
                    item.TongDS = item.DSChiQuay + item.DSChiNha + item.DSCK;
                    Report itemConvert = new Report()
                    {
                        DSChiQuay = Math.Round(item.DSChiQuay / item.TongDS * 100, 2, MidpointRounding.ToEven),
                        DSChiNha = Math.Round(item.DSChiNha / item.TongDS * 100, 2, MidpointRounding.ToEven),
                        DSCK = Math.Round(item.DSCK / item.TongDS * 100, 2, MidpointRounding.ToEven),
                        TongDS = Math.Round(item.TongDS / item.TongDS * 100, 2, MidpointRounding.ToEven),
                        Year = item.Year
                    };
                    DataConvertPercent.Add(itemConvert);
                }

                return DataConvertPercent;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        /// <summary>
        /// List Report cho so sánh last month
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<Report> SearchMonthCompareGrid(int ToYear, int ToMonth, string reportTypeID)
        {
            try
            {
                HSReportDAL dal = new HSReportDAL();
                List<Report> result = dal.SearchMonthCompareGrid(ToYear, ToMonth, reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }

        }

        /// <summary>
        /// List Report cho so sánh last month
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<Report> ListDataLastMonthCompareProportionPercent(int toYear, int toMonth, string reportTypeID)
        {
            try
            {
                HSReportDAL dal = new HSReportDAL();
                List<Report> DataConvertPercent = new List<Report>();
                List<Report> result = dal.SearchMonthCompareGrid(toYear, toMonth, reportTypeID);
                // Tháng hiện tại, tháng trước
                if (result.Count.Equals(3))
                {
                    foreach (Report item in result)
                    {
                        item.TongDS = item.DSChiQuay + item.DSChiNha + item.DSCK;
                        Report reportPercent = new Report()
                        {
                            ReportID = item.ReportID,
                            Day = item.Day,
                            Month = item.Month,
                            Year = item.Year,
                            GradationID = item.GradationID,
                            DSChiQuay = Math.Round(item.DSChiQuay / item.TongDS * 100, 2, MidpointRounding.ToEven),
                            DSChiNha = Math.Round(item.DSChiNha / item.TongDS * 100, 2, MidpointRounding.ToEven),
                            DSCK = Math.Round(item.DSCK / item.TongDS * 100, 2, MidpointRounding.ToEven)
                        };
                        DataConvertPercent.Add(reportPercent);
                    }
                }
                return DataConvertPercent;
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }
    }
}

﻿using DongA.Core;
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
    public class ReportDAL: DongABaseDAL
    {

        /// <summary>
        /// List Report theo ngày
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtServiceType> InsertTableMarket()
        {

            var result = new List<ReportDetailtServiceType>();

            string Baseurl = WebConfigurationManager.AppSettings["urlAPI"];

            using (var client = new HttpClient())
            {
                //Passing service base url  
                client.BaseAddress = new Uri(Baseurl);

                client.DefaultRequestHeaders.Clear();

                //Sending request to find web api REST service resource GetAllEmployees using HttpClient
                HttpResponseMessage Res = client.GetAsync("/api/InsertData/InsertTableMarket").Result;
                
                //Checking the response is successful or not which is sent using HttpClient  
                if (Res.IsSuccessStatusCode)
                {
                    //Storing the response details recieved from web api   
                    var AccResponse = Res.Content.ReadAsStringAsync().Result;

                    //Deserializing the response recieved from web api and storing into the Employee list  
                    result = JsonConvert.DeserializeObject<List<ReportDetailtServiceType>>(AccResponse);
                }
            }

            return result;
        }

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
        public List<Report> GetListReport(string reportType)
        {
            DbCommand command = null;
            try
            {
                var result = new List<Report>();
                result = DongADatabase.ToDataAPIObject<Report>("ReportData", "GetListReport", reportType);

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
        public List<Report> SearchDay(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<Report>();
                result = DongADatabase.ToDataAPIObject<Report>("ReportData", "SearchReportDay", "fromDate", fromDate, "toDate", toDate, "reportTypeID", reportTypeID);
                
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
            DbCommand command = null;
            try
            {
                var result = new List<Report>();
                // Get data cho theo API
                result = DongADatabase.ToDataAPIObject<Report>("ReportData", "ReportMonth", reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw DongAException.FromCommand(command, ex);
            }
        }
       
        /// <summary>
        /// List Report theo month
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
                result = DongADatabase.ToDataAPIObject<Report>("ReportData", "SearchReportMonth", "fromDate", fromDate, "toDate", toDate, "reportTypeID", reportTypeID);
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
        public List<Report> GetListReportYear(string reportTypeID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<Report>();
                result = DongADatabase.ToDataAPIObject<Report>("ReportData", "ReportYear", reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw DongAException.FromCommand(command, ex);
            }
        }

        /// <summary>
        /// List Report theo month
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<Report> SearchYear(DateTime fromYear, DateTime toYear, string reportTypeID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<Report>();
                result = DongADatabase.ToDataAPIObject<Report>("ReportData", "SearchReportYear", "fromDate", fromYear, "toDate", toYear, "reportTypeID", reportTypeID);
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
        public List<Report> ListDataGradationCompare(int toYear, int typeID, string reportTypeID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<Report>();
                result = DongADatabase.ToDataAPIObject<Report>("ReportData", "ListDataGradationCompare", "toYear", toYear, "typeID", typeID, "reportTypeID", reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw DongAException.FromCommand(command, ex);
            }
        }

        private const string SQL_All_GradationComparePercent = @"
select * from Report
where TYPEID = @typeID AND (YEAR(CreateDate) = @ToYear OR Year(DATEADD(YEAR, 1 , CreateDate)) = @ToYear)
";
        /// <summary>
        /// List data hiển thị phần trăm cho so sánh giai đoạn
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<Report> ListDataGradationComparePercent(int typeID, int ToYear)
        {
            DbCommand command = null;
            try
            {
                var result = new List<Report>();
                using (command = DongADatabase.GetSqlStringCommand(SQL_All_GradationComparePercent))
                {
                    DongADatabase.AddInParameter(command, "typeID", DbType.Int32, typeID);
                    DongADatabase.AddInParameter(command, "ToYear", DbType.Int32, ToYear);
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

        private const string SQL_All_GradationComparePie = @"
select * from Report
where TYPEID = @typeID AND YEAR(CreateDate) = @ToYear
";
        /// <summary>
        /// List data hiển thị phần trăm cho so sánh giai đoạn
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<Report> ListDataGradationComparePie(int typeID, int ToYear)
        {
            DbCommand command = null;
            try
            {
                var result = new List<Report>();
                using (command = DongADatabase.GetSqlStringCommand(SQL_All_GradationComparePie))
                {
                    DongADatabase.AddInParameter(command, "typeID", DbType.Int32, typeID);
                    DongADatabase.AddInParameter(command, "ToYear", DbType.Int32, ToYear);
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
        public List<Report> ListDataMonthCompareGrid(int toYear, int toMonth, string reportTypeID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<Report>();
                result = DongADatabase.ToDataAPIObject<Report>("ReportData", "ListDataCompareMonth", "toYear", toYear, "toMonth", toMonth, "reportTypeID", reportTypeID);
                
                return result;
            }
            catch (Exception ex)
            {
                throw DongAException.FromCommand(command, ex);
            }
        }

        private const string SQL_All_DataGirdLastMonthCompareProportion = @"
select * from Report
where TYPEID = @typeID AND (YEAR(CreateDate) = @ToYear OR Year(DATEADD(YEAR, 1 , CreateDate)) = @ToYear)
AND (MONTH(CreateDate) = @ToMonth OR MONTH(DATEADD(MONTH, 1 , CreateDate)) = @ToMonth)
";
        /// <summary>
        /// List Report cho so sánh giai đoạn
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<Report> ListDataLastMonthComparProportion(int typeID, int ToYear, int ToMonth)
        {
            DbCommand command = null;
            try
            {
                var result = new List<Report>();
                using (command = DongADatabase.GetSqlStringCommand(SQL_All_DataGirdLastMonthCompareProportion))
                {
                    DongADatabase.AddInParameter(command, "typeID", DbType.Int32, typeID);
                    DongADatabase.AddInParameter(command, "ToMonth", DbType.Int32, ToMonth);
                    DongADatabase.AddInParameter(command, "ToYear", DbType.Int32, ToYear);
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
        public List<ReportForMaket> DataReportMaketForDay(string reportTypeID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportForMaket>();
                result = DongADatabase.ToDataAPIObject<ReportForMaket>("ReportMaketData", "ReportDay", reportTypeID);
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
        public List<ReportForMaket> SearchReportMaketForDay(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportForMaket>();
                result = DongADatabase.ToDataAPIObject<ReportForMaket>("ReportMaketData", "SearchReportDay", "fromDate", fromDate, "toDate", toDate, "reportTypeID", reportTypeID);

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
        public List<ReportForMaket> DataReportMaketForMonth(string reportTypeID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportForMaket>();

                // Get data cho theo API
                result = DongADatabase.ToDataAPIObject<ReportForMaket>("ReportMaketData", "ReportMonth", reportTypeID);
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
        public List<ReportForMaket> SearchReportMaketForMonth(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportForMaket>();
                result = DongADatabase.ToDataAPIObject<ReportForMaket>("ReportMaketData", "SearchReportMonth", "fromDate", fromDate, "toDate", toDate, "reportTypeID", reportTypeID);

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
        public List<ReportForMaket> DataReportMaketForYear(string reportTypeID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportForMaket>();
                result = DongADatabase.ToDataAPIObject<ReportForMaket>("ReportMaketData", "ReportYear", reportTypeID);
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
        public List<ReportForMaket> SearchReportMaketForYear(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportForMaket>();
                result = DongADatabase.ToDataAPIObject<ReportForMaket>("ReportMaketData", "SearchReportYear", "fromDate", fromDate, "toDate", toDate, "reportTypeID", reportTypeID);

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
        public List<ReportForMaket> DataReportMaketForGradationCompare(int toYear, int typeID, string reportTypeID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportForMaket>();
                result = DongADatabase.ToDataAPIObject<ReportForMaket>("ReportMaketData", "ListDataGradationCompare", "toYear", toYear, "typeID", typeID, "reportTypeID", reportTypeID);

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
        public List<ReportForMaket> DataReportMaketForGradationCompareCompare(int toYear, int typeID, string reportTypeID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportForMaket>();
                result = DongADatabase.ToDataAPIObject<ReportForMaket>("ReportMaketData", "ListDataGradationCompare", "toYear", toYear, "typeID", typeID, "reportTypeID", reportTypeID);

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
        public List<ReportForMaket> DataReportCompareForMonth(int toYear, int toMonth, string reportTypeID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportForMaket>();
                result = DongADatabase.ToDataAPIObject<ReportForMaket>("ReportMaketData", "ListDataCompareMonth", "toYear", toYear, "toMonth", toMonth, "reportTypeID", reportTypeID);

                return result;
            }
            catch (Exception ex)
            {
                throw DongAException.FromCommand(command, ex);
            }
        }

        /// <summary>
        /// List Report theo ngày cho total payment
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportForTotalPayment> DataReportTPForDay(string reportTypeID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportForTotalPayment>();
                result = DongADatabase.ToDataAPIObject<ReportForTotalPayment>("ReportTPData", "ReportDay", reportTypeID);
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
        public List<ReportForTotalPayment> SearchReportTPForDay(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportForTotalPayment>();
                result = DongADatabase.ToDataAPIObject<ReportForTotalPayment>("ReportTPData", "SearchReportDay", "fromDate", fromDate, "toDate", toDate, "reportTypeID", reportTypeID);
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
        public List<ReportForTotalPayment> DataReportTPForMonth(string reportTypeID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportForTotalPayment>();
                result = DongADatabase.ToDataAPIObject<ReportForTotalPayment>("ReportTPData", "ReportMonth", reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw DongAException.FromCommand(command, ex);
            }
        }

        private const string SQL_SearchReportTotalPaymentForMonth = @"
select * from ReportTotalPayment where CreateDate between @FromDate and @ToDate AND TypeID = 2
";
        /// <summary>
        /// List Report
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportForTotalPayment> SearchReportTPForMonth(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportForTotalPayment>();
                result = DongADatabase.ToDataAPIObject<ReportForTotalPayment>("ReportTPData", "SearchReportMonth", "fromDate", fromDate, "toDate", toDate, "reportTypeID", reportTypeID);
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
        public List<ReportForTotalPayment> DataReportTPForYear(string reportTypeID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportForTotalPayment>();
                result = DongADatabase.ToDataAPIObject<ReportForTotalPayment>("ReportTPData", "ReportYear", reportTypeID);
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
        public List<ReportForTotalPayment> SearchReportTPForYear(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportForTotalPayment>();
                result = DongADatabase.ToDataAPIObject<ReportForTotalPayment>("ReportTPData", "SearchReportYear", "fromDate", fromDate, "toDate", toDate, "reportTypeID", reportTypeID);
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
        public List<ReportForTotalPayment> DataReportTPForGradationCompare(int toYear, int typeID, string reportTypeID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportForTotalPayment>();
                result = DongADatabase.ToDataAPIObject<ReportForTotalPayment>("ReportTPData", "ListDataGradationCompare", "toYear", toYear, "typeID", typeID, "reportTypeID", reportTypeID);
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
        public List<ReportForTotalPayment> DataReportTPCompareForMonth(int toYear, int toMonth, string reportTypeID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportForTotalPayment>();
                result = DongADatabase.ToDataAPIObject<ReportForTotalPayment>("ReportTPData", "ListDataCompareMonth", "toYear", toYear, "toMonth", toMonth, "reportTypeID", reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw DongAException.FromCommand(command, ex);
            }
        }

//        private const string SQL_SearchReportTotalPaymentCompareForMonth = @"
//select * from ReportTotalPayment
//where TYPEID = @typeID AND (YEAR(CreateDate) = @Year OR Year(DATEADD(YEAR, 1 , CreateDate)) = @Year) 
//    AND (MONTH(CreateDate) = @Month OR MONTH(DATEADD(MONTH, 1 , CreateDate)) = @Month)
//";
//        /// <summary>
//        /// List Report cho so sánh giai đoạn
//        /// </summary>
//        /// <returns></returns>
//        /// <history>
//        ///     [Truong Lam]   Created [10/06/2020]
//        /// </history>
//        public List<ReportForTotalPayment> SearchReportTotalPaymentCompareForMonth(int typeID, int year, int month)
//        {
//            DbCommand command = null;
//            try
//            {
//                var result = new List<ReportForTotalPayment>();
//                using (command = DongADatabase.GetSqlStringCommand(SQL_SearchReportTotalPaymentCompareForMonth))
//                {
//                    DongADatabase.AddInParameter(command, "typeID", DbType.Int32, typeID);
//                    DongADatabase.AddInParameter(command, "Month", DbType.Int32, month);
//                    DongADatabase.AddInParameter(command, "Year", DbType.Int32, year);
//                    using (var reader = DongADatabase.ExecuteReader(command, this))
//                    {
//                        result = DongADatabase.ToList<ReportForTotalPayment>(reader);
//                    }
//                }
//                return result;
//            }
//            catch (Exception ex)
//            {
//                throw DongAException.FromCommand(command, ex);
//            }
//        }

        /// <summary>
        /// List Report theo ngày
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportForTotalMoneyType> DataReportTotalMoneyTypeForDay(DateTime now, string reportTypeID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportForTotalMoneyType>();
                result = DongADatabase.ToDataAPIObject<ReportForTotalMoneyType>("ReportTMTData", "DataReportForDay", reportTypeID);
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
        public List<ReportForTotalMoneyType> DataReportTMTForDayConvert(DateTime now, string reportTypeID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportForTotalMoneyType>();
                result = DongADatabase.ToDataAPIObject<ReportForTotalMoneyType>("ReportTMTData", "DataReportForDayConvert", reportTypeID);
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
        public List<ReportForTotalMoneyType> SearchReportTMTForDay(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportForTotalMoneyType>();
                result = DongADatabase.ToDataAPIObject<ReportForTotalMoneyType>("ReportTMTData", "SearchReportForDay", "fromDate", fromDate, "toDate", toDate, "reportTypeID", reportTypeID);
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
        public List<ReportForTotalMoneyType> SearchReportTMTForDayConvert(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportForTotalMoneyType>();
                result = DongADatabase.ToDataAPIObject<ReportForTotalMoneyType>("ReportTMTData", "SearchReportForDayConvert", "fromDate", fromDate, "toDate", toDate, "reportTypeID", reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw DongAException.FromCommand(command, ex);
            }
        }

        /// <summary>
        /// List Report theo tháng - Nguyên tệ
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportForTotalMoneyType> DataReportTMTForMonth(string reportTypeID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportForTotalMoneyType>();
                result = DongADatabase.ToDataAPIObject<ReportForTotalMoneyType>("ReportTMTData", "DataReportypeForMonth", reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw DongAException.FromCommand(command, ex);
            }
        }

        /// <summary>
        /// List Report theo tháng - Nguyên tệ
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportForTotalMoneyType> DataReportTMTForMonthConvert(string reportTypeID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportForTotalMoneyType>();
                result = DongADatabase.ToDataAPIObject<ReportForTotalMoneyType>("ReportTMTData", "DataReportypeForMonthConvert", reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw DongAException.FromCommand(command, ex);
            }
        }

        /// <summary>
        /// Danh sách report tháng - Nguyên tệ
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportForTotalMoneyType> SearchReportTMTForMonth(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportForTotalMoneyType>();
                result = DongADatabase.ToDataAPIObject<ReportForTotalMoneyType>("ReportTMTData", "SearchReportForMonth", "fromDate", fromDate, "toDate", toDate, "reportTypeID", reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw DongAException.FromCommand(command, ex);
            }
        }

        /// <summary>
        /// Danh sách report tháng - Nguyên tệ
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportForTotalMoneyType> SearchReportTMTForMonthConvert(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportForTotalMoneyType>();
                result = DongADatabase.ToDataAPIObject<ReportForTotalMoneyType>("ReportTMTData", "SearchReportForMonthConvert", "fromDate", fromDate, "toDate", toDate, "reportTypeID", reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw DongAException.FromCommand(command, ex);
            }
        }

        /// <summary>
        /// List Report theo năm - Nguyên tệ
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportForTotalMoneyType> DataReportTMTForYear(string reportTypeID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportForTotalMoneyType>();
                result = DongADatabase.ToDataAPIObject<ReportForTotalMoneyType>("ReportTMTData", "DataReportypeForYear", reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw DongAException.FromCommand(command, ex);
            }
        }

        /// <summary>
        /// List Report theo năm - USD
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportForTotalMoneyType> DataReportTMTForYearConvert(string reportTypeID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportForTotalMoneyType>();
                result = DongADatabase.ToDataAPIObject<ReportForTotalMoneyType>("ReportTMTData", "DataReportypeForYearConvert", reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw DongAException.FromCommand(command, ex);
            }
        }

        /// <summary>
        /// List Report theo năm - Nguyên tệ
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportForTotalMoneyType> SearchReportTMTForYear(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportForTotalMoneyType>();
                result = DongADatabase.ToDataAPIObject<ReportForTotalMoneyType>("ReportTMTData", "SearchReportForYear", "fromDate", fromDate, "toDate", toDate, "reportTypeID", reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw DongAException.FromCommand(command, ex);
            }
        }

        /// <summary>
        /// List Report theo năm - USD
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportForTotalMoneyType> SearchReportTMTForYearConvert(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportForTotalMoneyType>();
                result = DongADatabase.ToDataAPIObject<ReportForTotalMoneyType>("ReportTMTData", "SearchReportForYearConvert", "fromDate", fromDate, "toDate", toDate, "reportTypeID", reportTypeID);
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
        public List<ReportForTotalMoneyType> DataReportTMTForGradationCompare(int year, int typeID, string reportTypeID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportForTotalMoneyType>();
                result = DongADatabase.ToDataAPIObject<ReportForTotalMoneyType>("ReportTMTData", "DataReportypeForGradationCompare", "year", year, "typeID", typeID, "reportTypeID", reportTypeID);
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
        public List<ReportForTotalMoneyType> DataReportTMTForGradationCompareConvert(int year, int typeID, string reportTypeID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportForTotalMoneyType>();
                result = DongADatabase.ToDataAPIObject<ReportForTotalMoneyType>("ReportTMTData", "DataReportypeForGradationCompareConvert", "year", year, "typeID", typeID, "reportTypeID", reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw DongAException.FromCommand(command, ex);
            }
        }

        /// <summary>
        /// List Report cho so sánh theo tháng - Nguyên tệ
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportForTotalMoneyType> DataReportTMTCompareForMonth(int toYear, int toMonth, string reportTypeID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportForTotalMoneyType>();
                result = DongADatabase.ToDataAPIObject<ReportForTotalMoneyType>("ReportTMTData", "DataReportypeCompareForMonth", "toYear", toYear, "toMonth", toMonth, "reportTypeID", reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw DongAException.FromCommand(command, ex);
            }
        }

        /// <summary>
        /// List Report cho so sánh theo tháng - USD
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportForTotalMoneyType> DataReportTMTCompareForMonthConvert(int toYear, int toMonth, string reportTypeID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportForTotalMoneyType>();
                result = DongADatabase.ToDataAPIObject<ReportForTotalMoneyType>("ReportTMTData", "DataReportypeCompareForMonthConvert", "toYear", toYear, "toMonth", toMonth, "reportTypeID", reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw DongAException.FromCommand(command, ex);
            }
        }

        private const string SQL_DataReportGradationComparePieForMonth = @"
select * from ReportTotalMoneyType
where TYPEID = @typeID AND (YEAR(CreateDate) = @Year OR Year(DATEADD(YEAR, 1 , CreateDate)) = @Year) 
    AND (MONTH(CreateDate) = @Month OR MONTH(DATEADD(MONTH, 1 , CreateDate)) = @Month)
";
        /// <summary>
        /// List Report cho so sánh giai đoạn
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportForTotalMoneyType> DataReportGradationComparePieForMonth(int typeID, int year, int month)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportForTotalMoneyType>();
                using (command = DongADatabase.GetSqlStringCommand(SQL_DataReportGradationComparePieForMonth))
                {
                    DongADatabase.AddInParameter(command, "typeID", DbType.Int32, typeID);
                    DongADatabase.AddInParameter(command, "Month", DbType.Int32, month);
                    DongADatabase.AddInParameter(command, "Year", DbType.Int32, year);
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
        /// List Report cho so sánh giai đoạn
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportForMaket> CreateDataMarket()
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportForMaket>();

                string Baseurl = WebConfigurationManager.AppSettings["urlAPI"];

                using (var client = new HttpClient())
                {
                    //Passing service base url  
                    client.BaseAddress = new Uri(Baseurl);

                    client.DefaultRequestHeaders.Clear();

                    //Sending request to find web api REST service resource GetAllEmployees using HttpClient  
                    HttpResponseMessage Res = client.GetAsync("/api/ReportMaketData/CreateDataMarket").Result;

                    //Checking the response is successful or not which is sent using HttpClient  
                    if (Res.IsSuccessStatusCode)
                    {
                        //Storing the response details recieved from web api   
                        var AccResponse = Res.Content.ReadAsStringAsync().Result;

                        //Deserializing the response recieved from web api and storing into the Employee list  
                        result = JsonConvert.DeserializeObject<List<ReportForMaket>>(AccResponse);
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
        /// List Report detailt theo ngày
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/08/2020]
        /// </history>
        public List<ReportDetailtServiceType> GetListReportDetailtForDay(string reportType)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportDetailtServiceType>();
                result = DongADatabase.ToDataAPIObject<ReportDetailtServiceType>("ReportDetailtDataForMarket", "DataReportDetailtDay", reportType);

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
        public List<ReportDetailtSTMarket> SearchMarketForTotalForDay(DateTime fromDate, DateTime toDate, string reportTypeID, string marketID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportDetailtSTMarket>();
                result = DongADatabase.ToDataAPIObject<ReportDetailtSTMarket>("ReportDetailtDataForMarket", "SearchDataReportDetailtDay", "fromDate", fromDate, "toDate", toDate, "reportTypeID", reportTypeID, "marketID", marketID);

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
        public List<ReportDetailtSTMarket> SearchMarketForTotalForMonth(DateTime fromDate, DateTime toDate, string reportTypeID, string marketID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportDetailtSTMarket>();
                result = DongADatabase.ToDataAPIObject<ReportDetailtSTMarket>("ReportDetailtDataForMarket", "SearchDataReportDetailtMonth", "fromDate", fromDate, "toDate", toDate, "reportTypeID", reportTypeID, "marketID", marketID);

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
        public List<ReportDetailtSTMarket> SearchMarketForTotalForYear(DateTime fromDate, DateTime toDate, string reportTypeID, string marketID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportDetailtSTMarket>();
                result = DongADatabase.ToDataAPIObject<ReportDetailtSTMarket>("ReportDetailtDataForMarket", "SearchDataReportDetailtYear", "fromDate", fromDate, "toDate", toDate, "reportTypeID", reportTypeID, "marketID", marketID);

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
        public List<ReportDetailtServiceType> MarketForPartner(string reportTypeID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportDetailtServiceType>();
                // Get data cho theo API
                //result = DongADatabase.ToDataAPIMarketObject<ReportDetailtServiceType>("ReportDetailtData", "ReportMonth", "partnerCode", partnerCode, reportTypeID);
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
        public List<Partner> ListPartner()
        {
            DbCommand command = null;
            try
            {
                var result = new List<Partner>();
                // Get data cho theo API
                result = DongADatabase.ToDataAPIObject<Partner>("ReportDetailtData", "ListPartner");
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
        public List<Market> ListMarket()
        {
            DbCommand command = null;
            try
            {
                var result = new List<Market>();
                // Get data cho theo API
                result = DongADatabase.ToDataAPIObject<Market>("ReportDetailtDataForMarket", "ListMarket");
                return result;
            }
            catch (Exception ex)
            {
                throw DongAException.FromCommand(command, ex);
            }
        }

        /// <summary>
        /// List Report detailt theo ngày của từng thị trường
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/08/2020]
        /// </history>
        public List<ReportDetailtServiceType> GetListReportDetailtForOneMarket(string reportType)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportDetailtServiceType>();
                result = DongADatabase.ToDataAPIObject<ReportDetailtServiceType>("ReportDetailtDataForMarket", "DataReportDetailtForOneMarket", reportType);

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
        public List<ReportDetailtServiceType> SearchMarketForOne(DateTime fromDate, DateTime toDate, string reportTypeID, string marketID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportDetailtServiceType>();
                result = DongADatabase.ToDataAPIObject<ReportDetailtServiceType>("ReportDetailtDataForMarket", "SearchDataReportDetailtForOneMarket", "fromDate", fromDate, "toDate", toDate, "reportTypeID", reportTypeID, "marketID", marketID);

                return result;
            }
            catch (Exception ex)
            {
                throw DongAException.FromCommand(command, ex);
            }
        }

        /// <summary>
        /// List Report detailt cho tháng của báo cáo chi tiết
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtServiceType> SearchMarketForOneForMonth(DateTime fromDate, DateTime toDate, string reportTypeID, string marketID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportDetailtServiceType>();
                result = DongADatabase.ToDataAPIObject<ReportDetailtServiceType>("ReportDetailtDataForMarket", "SearchDataReportDetailtForOneMarketForMonth", "fromDate", fromDate, "toDate", toDate, "reportTypeID", reportTypeID, "marketID", marketID);

                return result;
            }
            catch (Exception ex)
            {
                throw DongAException.FromCommand(command, ex);
            }
        }

        /// <summary>
        /// List Report detailt cho tháng của báo cáo chi tiết
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtServiceType> SearchMarketForOneForYear(DateTime fromDate, DateTime toDate, string reportTypeID, string marketID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportDetailtServiceType>();
                result = DongADatabase.ToDataAPIObject<ReportDetailtServiceType>("ReportDetailtDataForMarket", "SearchDataReportDetailtForOneMarketForYear", "fromDate", fromDate, "toDate", toDate, "reportTypeID", reportTypeID, "marketID", marketID);

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
        public List<ReportDetailtServiceType> ReportDetailtGradationCompareForAll(int toYear, int typeID, string reportTypeID, string marketID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportDetailtServiceType>();
                result = DongADatabase.ToDataAPIObject<ReportDetailtServiceType>("ReportDetailtDataForMarket", "SearchDataReportDetailtGradationForAll", "toYear", toYear, "typeID", typeID, "reportTypeID", reportTypeID, "marketID", marketID);
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
        public List<ReportDetailtServiceType> ReportDetailtGradationCompareForOne(int toYear, int typeID, string reportTypeID, string marketID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportDetailtServiceType>();
                result = DongADatabase.ToDataAPIObject<ReportDetailtServiceType>("ReportDetailtDataForMarket", "SearchDataReportDetailtGradationForOne", "toYear", toYear, "typeID", typeID, "reportTypeID", reportTypeID, "marketID", marketID);
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
        public List<ReportDetailtSTMarket> ReportDetailtCompareMonthForAll(int toYear, int toMonth, string reportTypeID, string marketID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportDetailtSTMarket>();
                result = DongADatabase.ToDataAPIObject<ReportDetailtSTMarket>("ReportDetailtDataForMarket", "SearchDataReportDetailtCompareMonthForAll", "toYear", toYear, "toMonth", toMonth, "reportTypeID", reportTypeID, "marketID", marketID);
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
        public List<ReportDetailtServiceType> ReportDetailtCompareMonthForOne(int toYear, int toMonth, string reportTypeID, string marketID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportDetailtServiceType>();
                result = DongADatabase.ToDataAPIObject<ReportDetailtServiceType>("ReportDetailtDataForMarket", "SearchDataReportDetailtCompareMonthForOne", "toYear", toYear, "toMonth", toMonth, "reportTypeID", reportTypeID, "marketID", marketID);
                return result;
            }
            catch (Exception ex)
            {
                throw DongAException.FromCommand(command, ex);
            }
        }

        #region Báo cáo chi tiết thị trường theo loại tiền chi trả
        
        /// <summary>
        /// List Report detailt theo ngày
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/08/2020]
        /// </history>
        public List<ReportDetailtForTotalMoneyType> ListReportDetailtMTForAll(string reportType)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportDetailtForTotalMoneyType>();
                result = DongADatabase.ToDataAPIObject<ReportDetailtForTotalMoneyType>("ReportDetailtDataForMarket", "DataReportDetailtMTForAll", reportType);

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
        public List<ReportDetailtForTotalMoneyType> SearchReportDetailtMTForAll(DateTime fromDate, DateTime toDate, string reportTypeID, string marketID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportDetailtForTotalMoneyType>();
                result = DongADatabase.ToDataAPIObject<ReportDetailtForTotalMoneyType>("ReportDetailtDataForMarket", "SearchDataReportDetailtMTForAll", "fromDate", fromDate, "toDate", toDate, "reportTypeID", reportTypeID, "marketID", marketID);

                return result;
            }
            catch (Exception ex)
            {
                throw DongAException.FromCommand(command, ex);
            }
        }

        /// <summary>
        /// List Report detailt theo ngày
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/08/2020]
        /// </history>
        public List<ReportDetailtForTotalMoneyType> ListReportDetailtMTForAllConvert(string reportType)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportDetailtForTotalMoneyType>();
                result = DongADatabase.ToDataAPIObject<ReportDetailtForTotalMoneyType>("ReportDetailtDataForMarket", "DataReportDetailtMTForAllConvert", reportType);

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
        public List<ReportDetailtForTotalMoneyType> SearchReportDetailtMTForAllConvert(DateTime fromDate, DateTime toDate, string reportTypeID, string marketID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportDetailtForTotalMoneyType>();
                result = DongADatabase.ToDataAPIObject<ReportDetailtForTotalMoneyType>("ReportDetailtDataForMarket", "SearchDataReportDetailtMTForAllConvert", "fromDate", fromDate, "toDate", toDate, "reportTypeID", reportTypeID, "marketID", marketID);

                return result;
            }
            catch (Exception ex)
            {
                throw DongAException.FromCommand(command, ex);
            }
        }

        /// <summary>
        /// List Report detailt theo ngày của từng thị trường
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/08/2020]
        /// </history>
        public List<ReportDetailtForTotalMoneyType> ListReportDetailtMTForOne(string reportType)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportDetailtForTotalMoneyType>();
                result = DongADatabase.ToDataAPIObject<ReportDetailtForTotalMoneyType>("ReportDetailtDataForMarket", "DataReportDetailtMTForOne", reportType);

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
        public List<ReportDetailtForTotalMoneyType> SearchReportDetailtMTForOne(DateTime fromDate, DateTime toDate, string reportTypeID, string marketID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportDetailtForTotalMoneyType>();
                result = DongADatabase.ToDataAPIObject<ReportDetailtForTotalMoneyType>("ReportDetailtDataForMarket", "SearchDataReportDetailtMTForOne", "fromDate", fromDate, "toDate", toDate, "reportTypeID", reportTypeID, "marketID", marketID);

                return result;
            }
            catch (Exception ex)
            {
                throw DongAException.FromCommand(command, ex);
            }
        }

        /// <summary>
        /// List Report detailt theo ngày của từng thị trường
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/08/2020]
        /// </history>
        public List<ReportDetailtForTotalMoneyType> ListReportDetailtMTForOneConvert(string reportType)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportDetailtForTotalMoneyType>();
                result = DongADatabase.ToDataAPIObject<ReportDetailtForTotalMoneyType>("ReportDetailtDataForMarket", "DataReportDetailtMTForOneConvert", reportType);

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
        public List<ReportDetailtForTotalMoneyType> SearchReportDetailtMTForOneConvert(DateTime fromDate, DateTime toDate, string reportTypeID, string marketID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportDetailtForTotalMoneyType>();
                result = DongADatabase.ToDataAPIObject<ReportDetailtForTotalMoneyType>("ReportDetailtDataForMarket", "SearchDataReportDetailtMTForOneConvert", "fromDate", fromDate, "toDate", toDate, "reportTypeID", reportTypeID, "marketID", marketID);

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
        public List<ReportDetailtForTotalMoneyType> ReportDetailtMTGradationCompareForAllConvert(int toYear, int typeID, string reportTypeID, string marketID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportDetailtForTotalMoneyType>();
                result = DongADatabase.ToDataAPIObject<ReportDetailtForTotalMoneyType>("ReportDetailtDataForMarket", "SearchDataReportDetailtMTGradationForAllConvert", "toYear", toYear, "typeID", typeID, "reportTypeID", reportTypeID, "marketID", marketID);
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
        public List<ReportDetailtForTotalMoneyType> ReportDetailtMTGradationCompareForOneConvert(int toYear, int typeID, string reportTypeID, string marketID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportDetailtForTotalMoneyType>();
                result = DongADatabase.ToDataAPIObject<ReportDetailtForTotalMoneyType>("ReportDetailtDataForMarket", "SearchDataReportDetailtMTGradationForOneConvert", "toYear", toYear, "typeID", typeID, "reportTypeID", reportTypeID, "marketID", marketID);
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
        public List<ReportDetailtForTotalMoneyType> ReportDetailtMTCompareMonthForAllConvert(int toYear, int toMonth, string reportTypeID, string marketID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportDetailtForTotalMoneyType>();
                result = DongADatabase.ToDataAPIObject<ReportDetailtForTotalMoneyType>("ReportDetailtDataForMarket", "SearchDataReportDetailtMTCompareMonthForAllConvert", "toYear", toYear, "toMonth", toMonth, "reportTypeID", reportTypeID, "marketID", marketID);
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
        public List<ReportDetailtForTotalMoneyType> ReportDetailtMTCompareMonthForOneConvert(int toYear, int toMonth, string reportTypeID, string marketID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportDetailtForTotalMoneyType>();
                result = DongADatabase.ToDataAPIObject<ReportDetailtForTotalMoneyType>("ReportDetailtDataForMarket", "SearchDataReportDetailtMTCompareMonthForOneConvert", "toYear", toYear, "toMonth", toMonth, "reportTypeID", reportTypeID, "marketID", marketID);
                return result;
            }
            catch (Exception ex)
            {
                throw DongAException.FromCommand(command, ex);
            }
        }

        #endregion

        #region Báo cáo chi tiết cho đối tác
        /// <summary>
        /// List Report detailt cho đối tác
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtForPartner> SearchReportDetailtPartnerForDay(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportDetailtForPartner>();
                result = DongADatabase.ToDataAPIObject<ReportDetailtForPartner>("ReportDetailtDataForPartner", "SearchDataReportDetailtForPartnerForAll", "fromDate", fromDate, "toDate", toDate, "reportTypeID", reportTypeID);

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
        public List<ReportDetailtForPartner> SearchPartnerForOne(DateTime fromDate, DateTime toDate, string reportTypeID, string partnerID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportDetailtForPartner>();
                result = DongADatabase.ToDataAPIObject<ReportDetailtForPartner>("ReportDetailtDataForPartner", "SearchDataReportDetailtForOneForDay", "fromDate", fromDate, "toDate", toDate, "reportTypeID", reportTypeID, "partnerID", partnerID);

                return result;
            }
            catch (Exception ex)
            {
                throw DongAException.FromCommand(command, ex);
            }
        }

        /// <summary>
        /// List Report detailt cho từng đối tác theo tháng
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtForPartner> SearchPartnerForOneForMonth(DateTime fromDate, DateTime toDate, string reportTypeID, string partnerID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportDetailtForPartner>();
                result = DongADatabase.ToDataAPIObject<ReportDetailtForPartner>("ReportDetailtDataForPartner", "SearchDataReportDetailtForOneForMonth", "fromDate", fromDate, "toDate", toDate, "reportTypeID", reportTypeID, "partnerID", partnerID);

                return result;
            }
            catch (Exception ex)
            {
                throw DongAException.FromCommand(command, ex);
            }
        }


        /// <summary>
        /// List Report detailt cho từng đối tác theo tháng
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtForPartner> SearchPartnerForOneForYear(DateTime fromDate, DateTime toDate, string reportTypeID, string partnerID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportDetailtForPartner>();
                result = DongADatabase.ToDataAPIObject<ReportDetailtForPartner>("ReportDetailtDataForPartner", "SearchDataReportDetailtForOneForYear", "fromDate", fromDate, "toDate", toDate, "reportTypeID", reportTypeID, "partnerID", partnerID);

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
        public List<ReportDetailtForPartner> ReportDetailtPartnerGradationCompareForAll(int toYear, int typeID, string reportTypeID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportDetailtForPartner>();
                result = DongADatabase.ToDataAPIObject<ReportDetailtForPartner>("ReportDetailtDataForPartner", "SearchDataReportDetailtGradationForAll", "toYear", toYear, "typeID", typeID, "reportTypeID", reportTypeID);
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
        public List<ReportDetailtForPartner> ReportDetailtPartnerGradationCompareForOne(int toYear, int typeID, string reportTypeID, string partnerID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportDetailtForPartner>();
                result = DongADatabase.ToDataAPIObject<ReportDetailtForPartner>("ReportDetailtDataForPartner", "SearchDataReportDetailtGradationForOne", "toYear", toYear, "typeID", typeID, "reportTypeID", reportTypeID, "partnerID", partnerID);
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
        public List<ReportDetailtForPartner> ReportDetailtPartnerCompareMonthForAll(int toYear, int toMonth, string reportTypeID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportDetailtForPartner>();
                result = DongADatabase.ToDataAPIObject<ReportDetailtForPartner>("ReportDetailtDataForPartner", "SearchDataReportDetailtCompareMonthForAll", "toYear", toYear, "toMonth", toMonth, "reportTypeID", reportTypeID);
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
        public List<ReportDetailtForPartner> ReportDetailtPartnerCompareMonthForOne(int toYear, int toMonth, string reportTypeID, string partnerID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportDetailtForPartner>();
                result = DongADatabase.ToDataAPIObject<ReportDetailtForPartner>("ReportDetailtDataForPartner", "SearchDataReportDetailtCompareMonthForOne", "toYear", toYear, "toMonth", toMonth, "reportTypeID", reportTypeID, "partnerID", partnerID);
                return result;
            }
            catch (Exception ex)
            {
                throw DongAException.FromCommand(command, ex);
            }
        }

        #endregion

        #region Báo cáo chi tiết - đối tác - loại tiền chi trả
        /// <summary>
        /// List Report detailt cho đối tác Nguyên tệ
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtForTotalMoneyType> SearchReportDetailtPartnerLTForDay(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportDetailtForTotalMoneyType>();
                result = DongADatabase.ToDataAPIObject<ReportDetailtForTotalMoneyType>("ReportDetailtDataForPartnerLT", "SearchDataReportDetailtForPartnerForAll", "fromDate", fromDate, "toDate", toDate, "reportTypeID", reportTypeID);

                return result;
            }
            catch (Exception ex)
            {
                throw DongAException.FromCommand(command, ex);
            }
        }

        /// <summary>
        /// List Report detailt cho đối tác Nguyên tệ
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtForTotalMoneyType> SearchReportDetailtPartnerLTForDayConvert(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportDetailtForTotalMoneyType>();
                result = DongADatabase.ToDataAPIObject<ReportDetailtForTotalMoneyType>("ReportDetailtDataForPartnerLT", "SearchDataReportDetailtForPartnerForAllConvert", "fromDate", fromDate, "toDate", toDate, "reportTypeID", reportTypeID);

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
        public List<ReportDetailtForTotalMoneyType> SearchPartnerLTForOne(DateTime fromDate, DateTime toDate, string reportTypeID, string partnerID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportDetailtForTotalMoneyType>();
                result = DongADatabase.ToDataAPIObject<ReportDetailtForTotalMoneyType>("ReportDetailtDataForPartnerLT", "SearchDataReportDetailtForOneForDay", "fromDate", fromDate, "toDate", toDate, "reportTypeID", reportTypeID, "partnerID", partnerID);

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
        public List<ReportDetailtForTotalMoneyType> SearchPartnerLTForOneConvert(DateTime fromDate, DateTime toDate, string reportTypeID, string partnerID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportDetailtForTotalMoneyType>();
                result = DongADatabase.ToDataAPIObject<ReportDetailtForTotalMoneyType>("ReportDetailtDataForPartnerLT", "SearchDataReportDetailtForOneForDayConvert", "fromDate", fromDate, "toDate", toDate, "reportTypeID", reportTypeID, "partnerID", partnerID);

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
        public List<ReportDetailtForTotalMoneyType> SearchPartnerLTForOneForMonth(DateTime fromDate, DateTime toDate, string reportTypeID, string partnerID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportDetailtForTotalMoneyType>();
                result = DongADatabase.ToDataAPIObject<ReportDetailtForTotalMoneyType>("ReportDetailtDataForPartnerLT", "SearchDataReportDetailtForOneForMonth", "fromDate", fromDate, "toDate", toDate, "reportTypeID", reportTypeID, "partnerID", partnerID);

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
        public List<ReportDetailtForTotalMoneyType> SearchPartnerLTForOneForMonthConvert(DateTime fromDate, DateTime toDate, string reportTypeID, string partnerID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportDetailtForTotalMoneyType>();
                result = DongADatabase.ToDataAPIObject<ReportDetailtForTotalMoneyType>("ReportDetailtDataForPartnerLT", "SearchDataReportDetailtForOneForMonthConvert", "fromDate", fromDate, "toDate", toDate, "reportTypeID", reportTypeID, "partnerID", partnerID);

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
        public List<ReportDetailtForTotalMoneyType> SearchPartnerLTForOneForYear(DateTime fromDate, DateTime toDate, string reportTypeID, string partnerID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportDetailtForTotalMoneyType>();
                result = DongADatabase.ToDataAPIObject<ReportDetailtForTotalMoneyType>("ReportDetailtDataForPartnerLT", "SearchDataReportDetailtForOneForYear", "fromDate", fromDate, "toDate", toDate, "reportTypeID", reportTypeID, "partnerID", partnerID);

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
        public List<ReportDetailtForTotalMoneyType> SearchPartnerLTForOneForYearConvert(DateTime fromDate, DateTime toDate, string reportTypeID, string partnerID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportDetailtForTotalMoneyType>();
                result = DongADatabase.ToDataAPIObject<ReportDetailtForTotalMoneyType>("ReportDetailtDataForPartnerLT", "SearchDataReportDetailtForOneForYearConvert", "fromDate", fromDate, "toDate", toDate, "reportTypeID", reportTypeID, "partnerID", partnerID);

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
        public List<ReportDetailtForTotalMoneyType> ReportDetailtPartnerLTGradationCompareForAll(int toYear, int typeID, string reportTypeID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportDetailtForTotalMoneyType>();
                result = DongADatabase.ToDataAPIObject<ReportDetailtForTotalMoneyType>("ReportDetailtDataForPartnerLT", "SearchDataReportDetailtGradationForAll", "toYear", toYear, "typeID", typeID, "reportTypeID", reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw DongAException.FromCommand(command, ex);
            }
        }


        /// <summary>
        /// List Report cho so sánh giai đoạn quy USD
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtForTotalMoneyType> ReportDetailtPartnerLTGradationCompareForAllConvert(int toYear, int typeID, string reportTypeID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportDetailtForTotalMoneyType>();
                result = DongADatabase.ToDataAPIObject<ReportDetailtForTotalMoneyType>("ReportDetailtDataForPartnerLT", "SearchDataReportDetailtGradationForAllConvert", "toYear", toYear, "typeID", typeID, "reportTypeID", reportTypeID);
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
        public List<ReportDetailtForTotalMoneyType> ReportDetailtPartnerLTGradationCompareForOne(int toYear, int typeID, string reportTypeID, string partnerID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportDetailtForTotalMoneyType>();
                result = DongADatabase.ToDataAPIObject<ReportDetailtForTotalMoneyType>("ReportDetailtDataForPartnerLT", "SearchDataReportDetailtGradationForOne", "toYear", toYear, "typeID", typeID, "reportTypeID", reportTypeID, "partnerID", partnerID);
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
        public List<ReportDetailtForTotalMoneyType> ReportDetailtPartnerLTGradationCompareForOneConvert(int toYear, int typeID, string reportTypeID, string partnerID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportDetailtForTotalMoneyType>();
                result = DongADatabase.ToDataAPIObject<ReportDetailtForTotalMoneyType>("ReportDetailtDataForPartnerLT", "SearchDataReportDetailtGradationForOneConvert", "toYear", toYear, "typeID", typeID, "reportTypeID", reportTypeID, "partnerID", partnerID);
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
        public List<ReportDetailtForTotalMoneyType> ReportDetailtPartnerLTCompareMonthForAll(int toYear, int toMonth, string reportTypeID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportDetailtForTotalMoneyType>();
                result = DongADatabase.ToDataAPIObject<ReportDetailtForTotalMoneyType>("ReportDetailtDataForPartnerLT", "SearchDataReportDetailtCompareMonthForAll", "toYear", toYear, "toMonth", toMonth, "reportTypeID", reportTypeID);
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
        public List<ReportDetailtForTotalMoneyType> ReportDetailtPartnerLTCompareMonthForAllConvert(int toYear, int toMonth, string reportTypeID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportDetailtForTotalMoneyType>();
                result = DongADatabase.ToDataAPIObject<ReportDetailtForTotalMoneyType>("ReportDetailtDataForPartnerLT", "SearchDataReportDetailtCompareMonthForAllConvert", "toYear", toYear, "toMonth", toMonth, "reportTypeID", reportTypeID);
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
        public List<ReportDetailtForTotalMoneyType> ReportDetailtPartnerLTCompareMonthForOne(int toYear, int toMonth, string reportTypeID, string partnerID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportDetailtForTotalMoneyType>();
                result = DongADatabase.ToDataAPIObject<ReportDetailtForTotalMoneyType>("ReportDetailtDataForPartnerLT", "SearchDataReportDetailtCompareMonthForOne", "toYear", toYear, "toMonth", toMonth, "reportTypeID", reportTypeID, "partnerID", partnerID);
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
        public List<ReportDetailtForTotalMoneyType> ReportDetailtPartnerLTCompareMonthForOneConvert(int toYear, int toMonth, string reportTypeID, string partnerID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportDetailtForTotalMoneyType>();
                result = DongADatabase.ToDataAPIObject<ReportDetailtForTotalMoneyType>("ReportDetailtDataForPartnerLT", "SearchDataReportDetailtCompareMonthForOneConvert", "toYear", toYear, "toMonth", toMonth, "reportTypeID", reportTypeID, "partnerID", partnerID);
                return result;
            }
            catch (Exception ex)
            {
                throw DongAException.FromCommand(command, ex);
            }
        }
        #endregion

        #region Get dữ liệu của City
        /// <summary>
        /// List Report theo month
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<City> SearchCity(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<City>();
                result = DongADatabase.ToDataAPIObject<City>("ReportData", "SearchDataCity", "fromDate", fromDate, "toDate", toDate, "reportTypeID", reportTypeID);
                return result;
            }
            catch (Exception ex)
            {
                throw DongAException.FromCommand(command, ex);
            }
        }
        #endregion


        #region Get dữ liệu cho TCKT

        /// <summary>
        /// List Report
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public List<ReportDetailtSTMarket> SearchMarketTCKTForDay(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportDetailtSTMarket>();
                result = DongADatabase.ToDataAPIObject<ReportDetailtSTMarket>("ReportDataTCKT", "SearchReportDay", "fromDate", fromDate, "toDate", toDate, "reportTypeID", reportTypeID);

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
        public List<ReportDetailtSTMarket> SearchMarketTCKTForMonth(DateTime fromDate, DateTime toDate, string reportTypeID)
        {
            DbCommand command = null;
            try
            {
                var result = new List<ReportDetailtSTMarket>();
                result = DongADatabase.ToDataAPIObject<ReportDetailtSTMarket>("ReportDataTCKT", "SearchReportMonth", "fromDate", fromDate, "toDate", toDate, "reportTypeID", reportTypeID);

                return result;
            }
            catch (Exception ex)
            {
                throw DongAException.FromCommand(command, ex);
            }
        }

        #endregion
    }
}

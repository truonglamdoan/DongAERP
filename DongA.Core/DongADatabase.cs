using System;
using System.Linq;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Practices.EnterpriseLibrary.Common.Configuration;
using Microsoft.Practices.EnterpriseLibrary.Data;
using EnterpriseDatabase = Microsoft.Practices.EnterpriseLibrary.Data.Database;
using System.ComponentModel.DataAnnotations.Schema;
using Oracle.ManagedDataAccess.Client;
using Dapper;
using System.Web.Configuration;
using System.Net.Http;
using Newtonsoft.Json;

namespace DongA.Core
{
    public class DongADatabase 
    {
        private static EnterpriseDatabase _Instance = null;
        private const string PARA_CREATEUSERID = "CreateUserID";
        private const string PARA_LASTMODIFYUSERID = "LastModifyUserID";
        private static Dictionary<int, object> _RowMappers = new Dictionary<int, object>();
        public static DbCommand GetSqlStringCommand(string query, bool autoAddDivisionID = true, bool isUsingCondition = true)
        {
            //// Tạo đối tượng command
            DbCommand result = Instance.GetSqlStringCommand(query);

            // Bỏ timeout
            result.CommandTimeout = 0;

            return result;
        }

        public static EnterpriseDatabase Instance
        {
            get
            {
                if (_Instance == null)
                {

                    //_Instance = EnterpriseLibraryContainer.Current.GetInstance<EnterpriseDatabase>("DongADbContext");
                    _Instance = EnterpriseLibraryContainer.Current.GetInstance<EnterpriseDatabase>("OracleDbContext");
                }
                return _Instance;
            }
            set { _Instance = value; }
        }

        ///// <summary>
        ///// Thêm các tham số mặc định
        ///// </summary>
        //public static void ApplyDefaultParameters(DbCommand command, bool autoAddDivisionID, bool isUsingCondition = true)
        //{
        //    AddInParameter(command, PARA_CREATEUSERID, DbType.String, DongAEnvironment.UserID, true);
        //    AddInParameter(command, PARA_LASTMODIFYUSERID, DbType.String, DongAEnvironment.UserID, true);

        //    ASOFTCondition condition = DongAEnvironment.GetCondition();
        //    if (condition != null && isUsingCondition)
        //    {
        //        AddCondition(command,
        //                     PARA_CONDITIONVT, condition.ConditionVT,
        //                     PARA_ISUSEDCONDITIONVT, condition.IsUsedConditionVT);

        //        AddCondition(command,
        //                     PARA_CONDITIONOB, condition.ConditionOB,
        //                     PARA_ISUSEDCONDITIONOB, condition.IsUsedConditionOB);

        //        AddCondition(command,
        //                     PARA_CONDITIONWA, condition.ConditionWA,
        //                     PARA_ISUSEDCONDITIONWA, condition.IsUsedConditionWA);

        //        AddCondition(command,
        //                     PARA_CONDITIONIV, condition.ConditionIV,
        //                     PARA_ISUSEDCONDITIONIV, condition.IsUsedConditionIV);

        //        AddCondition(command,
        //                     PARA_CONDITIONAC, condition.ConditionAC,
        //                     PARA_ISUSEDCONDITIONAC, condition.IsUsedConditionAC);

        //        AddCondition(command,
        //                     PARA_CONDITIONDE, condition.ConditionDE,
        //                     PARA_ISUSEDCONDITIONDE, condition.IsUsedConditionDE);
        //    }
        //}

        public static void AddInParameter(DbCommand command, string name, DbType dbType, object value, bool overrideValue)
        {
            // Đảm bảo tên tham số đúng chuẩn
            if (!name.StartsWith("@")) name = "@" + name;

            if (!command.Parameters.Contains(name))
            {
                Instance.AddInParameter(command, name, dbType, value);
            }
            else if (overrideValue)
            {
                Instance.SetParameterValue(command, name, value);
            }
        }

        public static void AddInParameter(DbCommand command, string name, DbType dbType, object value)
        {
            AddInParameter(command, name, dbType, value, false);
        }

        public static IDataReader ExecuteReader(DbCommand command, DongABaseDAL dal)
        {
            // Kiểm tra, xử lý khi có transaction và không có transaction
            var result = dal.Transaction == null
                             ? Instance.ExecuteReader(command)
                             : Instance.ExecuteReader(command, dal.Transaction);

            return result;
        }

        /// <summary>
        /// Tạo danh sách từ IDataReader
        /// </summary>
        /// <typeparam name="T">Kiểu đối tượng trong danh sách.</typeparam>
        /// <param name="reader">Đối tượng IDataReader</param>
        /// <param name="sqlName">Tên đầy đủ của câu sql. VD: Sql.BT0000.GetAll</param>
        /// <returns>Nếu không có record nào thì trả về danh sách rỗng.</returns>
        public static List<T> ToList<T>(IDataReader reader, string sqlName = "") where T : new()
        {
            var mapper = GetRowMapper<T>(reader, sqlName);

            var result = new List<T>();
            while (reader.Read())
            {
                result.Add(mapper.MapRow(reader));
            }
            return result;
        }

        /// <summary>
        /// Tạo đối tượng từ IDataReader
        /// </summary>
        /// <typeparam name="T">Kiểu đối tượng</typeparam>
        /// <param name="reader">Đối tượng IDataReader</param>
        /// <param name="sqlName">Tên đầy đủ của câu sql. VD: Sql.BT0000.GetAll</param>
        /// <returns>Nếu không có record nào thì trả về null.</returns>
        public static T FirstOrDefault<T>(IDataReader reader, string sqlName = "") where T : new()
        {
            var mapper = GetRowMapper<T>(reader, sqlName);

            T result = default(T);
            if (reader.Read())
            {
                result = mapper.MapRow(reader);
            }
            return result;
        }

        /// <summary>
        /// Tạo mapper
        /// </summary>
        /// <typeparam name="T">Kiểu đối tượng</typeparam>
        /// <param name="reader">Đối tượng IDataReader</param>
        /// <param name="sqlName">Tên đầy đủ của câu sql. VD: Sql.BT0000.GetAll</param>
        /// <returns></returns>
        private static IRowMapper<T> GetRowMapper<T>(IDataReader reader, string sqlName = "") where T : new()
        {
            IRowMapper<T> result = null;

            Type type = typeof(T);
            bool cache = !string.IsNullOrEmpty(sqlName);
            int hashCode = (type.FullName + sqlName).GetHashCode();

            // Kiểm tra mapper đã được cache hay chưa
            if (cache && RowMappers.ContainsKey(hashCode))
            {
                result = (IRowMapper<T>)RowMappers[hashCode];
            }
            // Nếu không được cache thì tạo mới
            else
            {
                IMapBuilderContext<T> context = MapBuilder<T>.MapNoProperties();

                PropertyInfo property = null;
                string propertyName = string.Empty;
                for (int i = reader.FieldCount - 1; i >= 0; i--)
                {

                    propertyName = reader.GetName(i);
                    property = type.GetProperty(propertyName);
                    // Trường hợp không lấy đc property thì lấy theo column name
                    if (property == null)
                    {
                        // Lấy dữ liệu property theo column của Object
                        property = type.GetProperties().FirstOrDefault(prop =>
                                        prop.GetCustomAttributes(false)
                                            .OfType<ColumnAttribute>()
                                            .Any(attribute => attribute.Name == propertyName));
                    }

                    // Chỉ map những field vừa có trong câu sql vừa có trong đối tượng
                    if (property != null && property.CanRead && property.CanWrite)
                        context.MapByName(property);
                }

                result = context.Build();

                // Nếu có yêu cầu cache thì lưu vào cache
                if (reader.FieldCount > 0 && cache)
                {
                    RowMappers.Add(hashCode, result);
                }
            }

            return result;
        }

        private static Dictionary<int, object> RowMappers
        {
            get { return _RowMappers; }
            set { _RowMappers = value; }
        }

        public static DataSet ExecuteDataSet(DbCommand command, DongABaseDAL dal)
        {
            // Kiểm tra, xử lý khi có transaction và không có transaction
            var result = dal.Transaction == null
                             ? Instance.ExecuteDataSet(command)
                             : Instance.ExecuteDataSet(command, dal.Transaction);

            return result;
        }

        public static DbCommand GetStoredProcCommand(string storedProcedureName)
        {
            DbCommand result = Instance.GetStoredProcCommand(storedProcedureName);
            result.CommandTimeout = 0;
            return result;
        }

        public static OracleCommand GetStoredProcCommandOracle(string storedProcedureName)
        {
            OracleCommand result = (OracleCommand)Instance.GetStoredProcCommand(storedProcedureName);
            result.CommandTimeout = 0;
            return result;
        }

        public static void AddInOracleParameter(DbCommand command, string name, OracleDbType type, ParameterDirection input, object nameValue)
        {
            OracleParameter ivalue = new OracleParameter(name, type);
            ivalue.Direction = input;
            ivalue.Value = nameValue;
            command.Parameters.Add(ivalue);
        }

        public static void AddInOracleParameterCursor(OracleCommand command, string name, OracleDbType type, ParameterDirection input)
        {
            OracleParameter output = command.Parameters.Add("p_cur", type);
            output.Direction = input;
        }

        public static int ExecuteNonQuery(DbCommand command, DongABaseDAL dal)
        {
            //Kiểm tra, xử lý khi có transaction và không có transaction
            var result = dal.Transaction == null
                             ? Instance.ExecuteNonQuery(command)
                             : Instance.ExecuteNonQuery(command, dal.Transaction);

            return result;
        }
        
        /// <summary>
        /// Đọc API theo Now
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="controllerAPI"></param>
        /// <param name="function"></param>
        /// <returns></returns>

        public static List<T> ToDataAPIObject<T>(string controllerAPI, string function, string reportTypeID = null) where T : new()
        {
            var result = new List<T>();

            string Baseurl = WebConfigurationManager.AppSettings["urlAPI"];

            using (var client = new HttpClient())
            {
                //Passing service base url  
                client.BaseAddress = new Uri(Baseurl);

                client.DefaultRequestHeaders.Clear();

                //Sending request to find web api REST service resource GetAllEmployees using HttpClient
                HttpResponseMessage Res = client.GetAsync(string.Format("/api/{0}/{1}?reportTypeID={2}", controllerAPI, function, reportTypeID)).Result;

                //Checking the response is successful or not which is sent using HttpClient  
                if (Res.IsSuccessStatusCode)
                {
                    //Storing the response details recieved from web api   
                    var AccResponse = Res.Content.ReadAsStringAsync().Result;

                    //Deserializing the response recieved from web api and storing into the Employee list  
                    result = JsonConvert.DeserializeObject<List<T>>(AccResponse);
                }
            }

            return result;
        }

        /// <summary>
        /// Search data API theo from date và toDate
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="controllerAPI"></param>
        /// <param name="function"></param>
        /// <param name="param1"></param>
        /// <param name="fromDate"></param>
        /// <param name="param2"></param>
        /// <param name="toDate"></param>
        /// <returns></returns>
        public static List<T> ToDataAPIObject<T>(string controllerAPI, string function, string param1, DateTime fromDate, string param2, DateTime toDate, string param3 = null, string reportTypeID = null, string param4 = null, string marketID = null) where T : new()
        {
            var result = new List<T>();

            string Baseurl = WebConfigurationManager.AppSettings["urlAPI"];

            using (var client = new HttpClient())
            {
                //Passing service base url  
                client.BaseAddress = new Uri(Baseurl);

                client.DefaultRequestHeaders.Clear();

                HttpResponseMessage Res = new HttpResponseMessage();
                //Sending request to find web api REST service resource GetAllEmployees using HttpClient
                if (string.IsNullOrEmpty(marketID))
                {
                    Res = client.GetAsync(string.Format("/api/{0}/{1}?{2}={3}&{4}={5}&{6}={7}", controllerAPI, function, param1, fromDate.ToString("MM/dd/yyyy"), param2, toDate.ToString("MM/dd/yyyy"), param3, reportTypeID)).Result;
                }
                else
                {
                    Res = client.GetAsync(string.Format("/api/{0}/{1}?{2}={3}&{4}={5}&{6}={7}&{8}={9}", controllerAPI, function, param1, fromDate.ToString("MM/dd/yyyy"), param2, toDate.ToString("MM/dd/yyyy"), param3, reportTypeID, param4, marketID)).Result;
                }

                //Checking the response is successful or not which is sent using HttpClient  
                if (Res.IsSuccessStatusCode)
                {
                    //Storing the response details recieved from web api   
                    var AccResponse = Res.Content.ReadAsStringAsync().Result;

                    //Deserializing the response recieved from web api and storing into the Employee list  
                    result = JsonConvert.DeserializeObject<List<T>>(AccResponse);
                }
            }

            return result;
        }

        /// <summary>
        /// Search data API theo giai đoạn và so sánh
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="controllerAPI"></param>
        /// <param name="function"></param>
        /// <param name="param1"></param>
        /// <param name="typeID"></param>
        /// <param name="param2"></param>
        /// <param name="year"></param>
        /// <returns></returns>
        public static List<T> ToDataAPIObject<T>(string controllerAPI, string function, string param1, int typeID, string param2, int year, string param3 = null, string reportTypeID = null, string param4 = null, string marketID = null) where T : new()
        {
            var result = new List<T>();

            string Baseurl = WebConfigurationManager.AppSettings["urlAPI"];

            using (var client = new HttpClient())
            {
                //Passing service base url  
                client.BaseAddress = new Uri(Baseurl);

                client.DefaultRequestHeaders.Clear();

                //Sending request to find web api REST service resource GetAllEmployees using HttpClient
                HttpResponseMessage Res = new HttpResponseMessage();
                if (!string.IsNullOrEmpty(param4) && !string.IsNullOrEmpty(marketID))
                {
                    Res = client.GetAsync(string.Format("/api/{0}/{1}?{2}={3}&{4}={5}&{6}={7}&{8}={9}", controllerAPI, function, param1, typeID, param2, year, param3, reportTypeID, param4, marketID)).Result;
                }
                else
                {
                    Res = client.GetAsync(string.Format("/api/{0}/{1}?{2}={3}&{4}={5}&{6}={7}", controllerAPI, function, param1, typeID, param2, year, param3, reportTypeID)).Result;
                }

                //Checking the response is successful or not which is sent using HttpClient  
                if (Res.IsSuccessStatusCode)
                {
                    //Storing the response details recieved from web api   
                    var AccResponse = Res.Content.ReadAsStringAsync().Result;

                    //Deserializing the response recieved from web api and storing into the Employee list  
                    result = JsonConvert.DeserializeObject<List<T>>(AccResponse);
                }
            }

            return result;
        }

        public static List<T> ToDataAPIMarketObject<T>(string controllerAPI, string function, string param1, string value1, string reportTypeID = null) where T : new()
        {
            var result = new List<T>();

            string Baseurl = WebConfigurationManager.AppSettings["urlAPI"];

            using (var client = new HttpClient())
            {
                //Passing service base url  
                client.BaseAddress = new Uri(Baseurl);

                client.DefaultRequestHeaders.Clear();

                //Sending request to find web api REST service resource GetAllEmployees using HttpClient
                HttpResponseMessage Res = client.GetAsync(string.Format("/api/{0}/{1}?{2}={3}&reportTypeID={4}", controllerAPI, function, param1, value1, reportTypeID)).Result;

                //Checking the response is successful or not which is sent using HttpClient  
                if (Res.IsSuccessStatusCode)
                {
                    //Storing the response details recieved from web api   
                    var AccResponse = Res.Content.ReadAsStringAsync().Result;

                    //Deserializing the response recieved from web api and storing into the Employee list  
                    result = JsonConvert.DeserializeObject<List<T>>(AccResponse);
                }
            }

            return result;
        }
    }
}

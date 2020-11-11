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
    public class TargetDAL : DongABaseDAL
    {
        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/11/2020]
        /// </history>
        public List<FormTarget> GetListTarget()
        {
            DbCommand command = null;
            try
            {
                var result = new List<FormTarget>();
                result = DongADatabase.ToDataAPIObject<FormTarget>("Target", "ListTarget");

                return result;
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
        
        public bool InsertTarget(FormTarget data)
        {
            DbCommand command = null;
            try
            {
                string Baseurl = WebConfigurationManager.AppSettings["urlAPI"];

                using (var client = new HttpClient())
                {
                    //Passing service base url  
                    client.BaseAddress = new Uri(Baseurl);

                    client.DefaultRequestHeaders.Clear();

                    //Sending request to find web api REST service resource GetAllEmployees using HttpClient
                    HttpResponseMessage Res = client.PostAsJsonAsync<FormTarget>("/api/Target/InsertdataTarget", data).Result;

                    //Checking the response is successful or not which is sent using HttpClient  
                    if (Res.IsSuccessStatusCode)
                    {
                        //Storing the response details recieved from web api   
                        var AccResponse = Res.Content.ReadAsStringAsync().Result;
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

        public bool UpdateDataTarget(FormTarget data)
        {
            DbCommand command = null;
            try
            {
                if(!string.IsNullOrEmpty(data.ObjectID))
                {
                    data.CustomDate = DateTime.Now;
                }
                
                string Baseurl = WebConfigurationManager.AppSettings["urlAPI"];

                using (var client = new HttpClient())
                {
                    //Passing service base url  
                    client.BaseAddress = new Uri(Baseurl);

                    client.DefaultRequestHeaders.Clear();

                    //Sending request to find web api REST service resource GetAllEmployees using HttpClient
                    HttpResponseMessage Res = client.PostAsJsonAsync<FormTarget>("/api/Target/UpdatedataTarget", data).Result;

                    //Checking the response is successful or not which is sent using HttpClient  
                    if (Res.IsSuccessStatusCode)
                    {
                        //Storing the response details recieved from web api   
                        var AccResponse = Res.Content.ReadAsStringAsync().Result;
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

        public bool DeleteDataTarget(FormTarget data)
        {
            DbCommand command = null;
            try
            {
                if (!string.IsNullOrEmpty(data.ObjectID))
                {
                    data.CustomDate = DateTime.Now;
                }

                string Baseurl = WebConfigurationManager.AppSettings["urlAPI"];

                using (var client = new HttpClient())
                {
                    //Passing service base url  
                    client.BaseAddress = new Uri(Baseurl);

                    client.DefaultRequestHeaders.Clear();

                    //Sending request to find web api REST service resource GetAllEmployees using HttpClient
                    HttpResponseMessage Res = client.PostAsJsonAsync<FormTarget>("/api/Target/DeleteDataTarget", data).Result;

                    //Checking the response is successful or not which is sent using HttpClient  
                    if (Res.IsSuccessStatusCode)
                    {
                        //Storing the response details recieved from web api   
                        var AccResponse = Res.Content.ReadAsStringAsync().Result;
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
    }
}

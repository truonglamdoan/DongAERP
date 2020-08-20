using DongA.Bussiness;
using DongA.Core;
using DongA.Entities;
using DongAERP.Areas.Admin.Code;
using DongAERP.Areas.Admin.Models;
using Models;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Security.Cryptography;
using System.Text;
using System.Web;
using System.Web.Configuration;
using System.Web.Mvc;
using System.Web.Security;
using static DongA.Core.DongAException;

namespace DongAERP.Areas.Admin.Controllers
{
    public class LoginController : Controller
    {
        // Key mã hóa
        const string AES_KEY = "NdRgUkXp2r5u8x/A";
        const string AES_IV = "+KbPeShVkYp3s6v9";
        const string AppCode = "E-STATISTIC";

        // GET: Admin/Login
        [HttpGet]
        public ActionResult Index()
        {
            return View(new LoginModel());
        }

        /// <summary>
        /// Authenticate thực hiện xác thực người dùng
        /// </summary>
        /// <param name="username"></param>
        /// <param name="password"></param>
        /// <param name="appCode"></param>
        /// <returns>
        /// <para>Code = "Success": Thành công</para>
        /// <para>Code = "InfoNotFound": Đăng nhập thành công nhưng không có thông tin nhân viên</para>
        /// <para>Code = "InvalidUser": Tên đăng nhập không tồn tại</para>
        /// <para>Code = "WrongPassword": Sai mật khẩu</para>
        /// <para>Code = "AccessDenied": Không có quyền truy cập ứng dụng</para>
        /// <para>Code = "LdapError": Lỗi máy chủ LDAP</para>
        /// <para>Code = "DatabaseError": Lỗi database</para>
        /// <para>Code = "SystemError": Lỗi hệ thống</para>
        /// </returns>
        [HttpPost]
        [ValidateAntiForgeryToken]
        [AllowAnonymous]
        public ActionResult Index(LoginModel model)
        {
            try
            {
                string Baseurl = WebConfigurationManager.AppSettings["urlLoginAPI"];
                string userNameConvert = string.IsNullOrEmpty(model.UserName) ? string.Empty :  AESEncrypt(model.UserName);
                string passConvert = string.IsNullOrEmpty(model.Password) ? string.Empty : AESEncrypt(model.Password);
                string appCodeConvert = AESEncrypt(AppCode);

                string urlStr = string.Format("{0}damtc/auth/authenticate?Username={1}&Password={2}&AppCode={3}", Baseurl, HttpUtility.UrlEncodeUnicode(userNameConvert), HttpUtility.UrlEncodeUnicode(passConvert), HttpUtility.UrlEncodeUnicode(appCodeConvert));

                //string Baseurl = WebConfigurationManager.AppSettings["urlAPI"];
                if (!string.IsNullOrEmpty(Baseurl))
                {
                    using (var client = new HttpClient())
                    {
                        //Passing service base url  
                        client.BaseAddress = new Uri(Baseurl);

                        client.DefaultRequestHeaders.Clear();

                        //Sending request to find web api REST service resource GetAllEmployees using HttpClient  
                        HttpResponseMessage Res = client.GetAsync(urlStr).Result;

                        //Checking the response is successful or not which is sent using HttpClient  
                        if (Res.IsSuccessStatusCode)
                        {
                            //Storing the response details recieved from web api   
                            var AccResponse = Res.Content.ReadAsStringAsync().Result;

                            //Deserializing the response recieved from web api and storing into the Employee list  
                            Account result = JsonConvert.DeserializeObject<Account>(AccResponse);

                            if (result.Code.Contains("Success") && ModelState.IsValid)
                            {
                                // Set biến môi trường
                                DongAEnvironment.UserID = result.EmployeeCode;
                                DongAEnvironment.Fullname = result.FullName;
                                //SessionHelper.SetSession(new UserSession() { UserName = model.UserName });
                                FormsAuthentication.SetAuthCookie(model.UserName, model.RememberMe);

                                //// Chạy insert bảng dữ liệu cho doanh số chi tiết
                                //List<ReportDetailtServiceType> listData = new ReportBL().InsertTableMarket();

                                return RedirectToAction("Index", "Home");
                            }
                            else
                            {
                                ModelState.AddModelError("", result.Message);
                            }

                        }
                        else
                        {
                            ModelState.AddModelError(string.Empty, "Tên đăng nhập và mật khẩu không được để trống!");
                        }
                        //returning the employee list to view  
                        return View(model);
                    }
                }
                else
                {
                    ModelState.AddModelError("", "Có lỗi trong quá trình đăng nhập, liên hệ admin");
                    return View(model);
                }
                
            }
            catch (Exception ex)
            {
                throw new DongAException(DongALayer.Business, ex.Message, ex);
            }
        }

        public ActionResult Logout()
        {
            FormsAuthentication.SignOut();
            return RedirectToAction("Index", "Login");
        }

        private string AESEncrypt(string plainText)
        {
            var keyBytes = Encoding.UTF8.GetBytes(AES_KEY);
            var ivBytes = Encoding.UTF8.GetBytes(AES_IV);

            var result = AESEncrypt(plainText, keyBytes, ivBytes);

            return Convert.ToBase64String(result);
        }

        private byte[] AESEncrypt(string plainText, byte[] key, byte[] iv)
        {
            if (string.IsNullOrEmpty(plainText))
                throw new ArgumentNullException("plainText");

            if (key == null || key.Length == 0)
                throw new ArgumentNullException("key");

            if (iv == null || iv.Length == 0)
                throw new ArgumentNullException("iv");

            byte[] result;
            using (var rijAlg = new RijndaelManaged
            {
                Mode = CipherMode.CBC,
                Padding = PaddingMode.PKCS7,
                FeedbackSize = 128,
                Key = key,
                IV = iv,
            })
            {
                var encryptor = rijAlg.CreateEncryptor(rijAlg.Key, rijAlg.IV);

                using (var ms = new MemoryStream())
                {
                    using (var cs = new CryptoStream(ms, encryptor, CryptoStreamMode.Write))
                    {
                        using (var swEncrypt = new StreamWriter(cs))
                            swEncrypt.Write(plainText);
                        result = ms.ToArray();
                    }
                }
            }
            return result;
        }

    }
}
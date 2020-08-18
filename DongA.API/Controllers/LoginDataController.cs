using DongA.API.Bussiness;
using DongA.Entities;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Security.Cryptography;
using System.Text;
using System.Web.Http;

namespace DongA.API.Controllers
{
    public class LoginDataController : ApiController
    {

        [HttpGet]
        public List<Account> Login()
        {
            //bool result = new AccountModel().Login(model.UserName, model.Password);
            List<Account> listAccount = new ReportBL().GetListAccount();
            // Gán quyền Admin
            listAccount.Add(new Account() { EmployeeID = "Admin", Password = "admin" });
            return listAccount;
        }
    }
}

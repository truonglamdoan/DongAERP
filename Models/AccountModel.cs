// #################################################################
// # Copyright (C) 2010-2011, ASoft JSC.  All Rights Reserved.
// #
// # History：
// #	Date Time		Updated			Content
// #	10/06/2020		Đức Nhân		Tạo mới
// ##################################################################

using Models.Framework;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Models
{
    public class AccountModel
    {
        private DongADbContext context = null;

        public AccountModel()
        {
            context = new DongADbContext();
        }

        /// <summary>
        /// Thông tin đăng nhập
        /// </summary>
        /// <returns></returns>
        /// <history>
        ///     [Truong Lam]   Created [10/06/2020]
        /// </history>
        public bool Login(string userName, string password)
        {
            if (string.IsNullOrEmpty(password))
            {
                password=string.Empty;
            }
            object[] sqlParams =
            {
                new SqlParameter("@UserName", userName),
                new SqlParameter("@Password", password)
            };
            var res = context.Database.SqlQuery<bool>("SP_Account @UserName, @Password", sqlParams).SingleOrDefault();
            return res;
        }
    }
}

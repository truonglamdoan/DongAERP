using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace DongAERP.Areas.Admin.Code
{
    public class SessionHelper
    {
        /// <summary>
        /// Set session
        /// </summary>
        /// <param name="session"></param>
        public static void SetSession(UserSession session)
        {
            HttpContext.Current.Session["loginSession"] = session;
        }

        /// <summary>
        /// Get session
        /// </summary>
        /// <returns></returns>
        public static UserSession GetSession()
        {
            var session = HttpContext.Current.Session["loginSession"];
            if (session == null)
            {
                return null;
            }
            else
            {
                return session as UserSession;
            }
        }
    }
}
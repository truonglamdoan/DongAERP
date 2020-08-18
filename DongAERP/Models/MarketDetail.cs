// #################################################################
// # Copyright (C) 2010-2011, ASoft JSC.  All Rights Reserved.
// #
// # History：
// #	Date Time		Updated			Content
// #	10/06/2020		Đức Nhân		Tạo mới
// ##################################################################

using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace DongAERP.Models
{
    public class MarketDetail
    {
        public int APK { get; set; }
        public string MarKetDetailID { get; set; }
        public string MarKetDetailName { get; set; }
        public string CreateDate { get; set; }
        public string CreateUserID { get; set; }
        public string MarKetID { get; set; }
        public string MarKetAsianID { get; set; }
    }
}
// #################################################################
// # Copyright (C) 2010-2011, ASoft JSC.  All Rights Reserved.
// #
// # History：
// #	Date Time		Updated			Content
// #	10/08/2020		Truong Lam		Create New
// ##################################################################

using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DongA.Entities
{
    [Table("MD_MARKET")]
    public class Market
    {
        [Column("MARKETCODE")]
        public string MarketCode { set; get; }
        [Column("MARKETNAME")]
        public string MarketName { set; get; }
        [Column("PARENTCODE")]
        public string ParentCode { set; get; }
    }
}

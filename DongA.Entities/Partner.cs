// #################################################################
// # Copyright (C) 2010-2011, ASoft JSC.  All Rights Reserved.
// #
// # History：
// #	Date Time		Updated			Content
// #	10/06/2020		Truong Lam		Create New
// ##################################################################

using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DongA.Entities
{
    [Table("MD_PARTNER")]
    public class Partner
    {
        [Column("PARTNERCODE")]
        public string PartnerCode { set; get; }
        public string PartnerID { set; get; }

        [Column("PARTNERNAME")]
        public string PartnerName { set; get; }
        public string CreateUserID { set; get; }
        public DateTime CreateDate { set; get; }
    }
}

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
    [Table("RPT_TURNOVER")]
    public class Report
    {
        public string ReportID { set; get; }
        [Column("DSCHIQUAY")]
        public double DSChiQuay { set; get; }
        [Column("DSCHINHA")]
        public double DSChiNha { set; get; }
        [Column("DSCK")]
        public double DSCK { set; get; }
        public double TongDS { set; get; }
        public double Type { set; get; }
        [Column("CREATEDDATE")]
        public DateTime CreatedDate { set; get; }

        public string Day { get; set; }
        [Column("MONTH")]
        public string Month { get; set; }
        [Column("YEAR")]
        public string Year { get; set; }

        public string GradationID { get; set; }


    }
}

// #################################################################
// # Copyright (C) 2010-2011, ASoft JSC.  All Rights Reserved.
// #
// # History：
// #	Date Time		Updated			Content
// #	10/11/2020		Truong Lam		Create New
// ##################################################################

using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DongA.Entities
{
    public class FormTarget
    {
        [Column("OBJECTID")]
        public string ObjectID { set; get; }
        [Column("OBJECTNAME")]
        public string ObjectName { set; get; }
        [Column("TARGETVALUE")]
        public double TargetValue { set; get; }
        public double COL1 { set; get; }
        public double COL2 { set; get; }
        public double COL3 { set; get; }
        public double COL4 { set; get; }
        public double COL5 { set; get; }
        public double COL6 { set; get; }
        public double COL7 { set; get; }
        public double COL8 { set; get; }
        public double COL9 { set; get; }
        public double COL10 { set; get; }
        public double COL11 { set; get; }
        public double COL12 { set; get; }
        [Column("YEAR")]
        public int Year { set; get; }
        [Column("CREATEDDATE")]
        public DateTime CreatedDate { set; get; }
        [Column("CUSTOMDATE")]
        public DateTime CustomDate { set; get; }

        [Column("TYPEID")]
        public string TypeID { set; get; }
    }
}

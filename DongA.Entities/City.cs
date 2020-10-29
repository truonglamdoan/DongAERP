// #################################################################
// # Copyright (C) 2010-2011, ASoft JSC.  All Rights Reserved.
// #
// # History：
// #	Date Time		Updated			Content
// #	10/10/2020		Truong Lam		Create New
// ##################################################################

using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
namespace DongA.Entities
{
    public class City
    {
        [Column("CITYCODE")]
        public string CityCode { set; get; }

        [Column("CITYNAME")]
        public string CityName { set; get; }
        [Column("AMOUNT")]
        public double AMOUNT { set; get; }
        [Column("RECNUM")]
        public double RECNUM { set; get; }
        public double TongCity { set; get; }

        public string GradationID { get; set; }
    }
}

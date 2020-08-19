using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DongA.Entities
{
    [Table("RPT_TURNOVER")]
    public class ReportForMaket
    {
        public string ReportID { set; get; }
        [Column("AMERICAN")]
        public double American { set; get; }
        [Column("ASIA")]
        public double Asia { set; get; }
        [Column("GLOBAL")]
        public double Global { set; get; }
        [Column("EUROPE")]
        public double Europe { set; get; }
        [Column("CANADA")]
        public double Canada { set; get; }
        [Column("AUSTRALIA")]
        public double Australia { set; get; }
        public double TongDS { set; get; }
        public double Type { set; get; }
        [Column("CREATEDDATE")]
        public DateTime CreatedDate { set; get; }
        [Column("MONTH")]
        public string Month { get; set; }
        [Column("YEAR")]
        public string Year { get; set; }
        public string TypeID { set; get; }
    }
}

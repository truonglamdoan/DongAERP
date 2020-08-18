using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DongA.Entities
{
    [Table("RPT_TURNOVER")]
    public class ReportForTotalPayment
    {
        public string ReportID { set; get; }
        [Column("PAYED")]
        public double Payed { set; get; }
        public double TongDS { set; get; }
        [Column("CREATEDDATE")]
        public DateTime CreatedDate { set; get; }
        public string Day { get; set; }
        [Column("MONTH")]
        public string Month { get; set; }
        [Column("YEAR")]
        public string Year { get; set; }
        public string TypeID { set; get; }
        public double Type { set; get; }
    }
}

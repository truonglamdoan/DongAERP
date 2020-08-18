using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DongA.Entities
{
    [Table("RPT_TURNOVER")]
    public class ReportForTotalMoneyType
    {
        public string ReportID { set; get; }
        public double VND { set; get; }
        public double USD { set; get; }
        public double EUR { set; get; }
        public double CAD { set; get; }
        public double AUD { set; get; }
        public double GBP { set; get; }
        public double TongDS { set; get; }
        public double Type { set; get; }
        [Column("CREATEDDATE")]
        public DateTime CreatedDate { set; get; }

        public string Day { get; set; }
        [Column("MONTH")]
        public string Month { get; set; }
        [Column("YEAR")]
        public string Year { get; set; }
        public string TypeID { set; get; }
    }
}

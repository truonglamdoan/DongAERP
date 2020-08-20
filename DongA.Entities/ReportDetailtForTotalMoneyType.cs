using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace DongA.Entities
{
    public class ReportDetailtForTotalMoneyType
    {

        public string STT { set; get; }

        [Column("MARKETCODE")]
        public string MarketCode { get; set; }

        [Column("MARKETNAME")]
        public string MarketName { get; set; }

        [Column("PARTNERCODE")]
        public string PartnerCode { get; set; }

        [Column("PARTNERNAME")]
        public string PartnerName { get; set; }
        
        public double VND { set; get; }
        public double USD { set; get; }
        public double EUR { set; get; }
        public double CAD { set; get; }
        public double AUD { set; get; }
        public double GBP { set; get; }
        public double TongDS { set; get; }
        [Column("CREATEDDATE")]
        public DateTime CreatedDate { set; get; }
        
        [Column("MONTH")]
        public string Month { get; set; }
        [Column("YEAR")]
        public string Year { get; set; }
    }
}

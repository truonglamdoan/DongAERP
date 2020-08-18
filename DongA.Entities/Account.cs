using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DongA.Entities
{
    [Table("RPT_TURNOVER")]
    public class Account
    {
        [Column("EMPLOYEEID")]
        public string EmployeeID { set; get; }
        [Column("FULLNAME")]
        public string FullName { set; get; }
        [Column("PASSWORD")]
        public string Password { set; get; }

        public string Code { get; set; }

        public string Message { get; set; }

        public string EmployeeCode { get; set; }

        public bool IsSystem { get; set; }
    }
}

using DongA.Core;
using DongA.DataAccess;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static DongA.Core.DongAException;

namespace DongA.Bussiness
{
    public class CommonBL : DongABaseDAL
    {
        private CommonDAL dal = new CommonDAL();
        ///// <summary>
        /////   Lấy dữ liệu cho báo cáo
        ///// </summary>
        ///// <param name="SQL">Câu sql truy vấn</param>
        ///// <returns>Dataset chứa dữ liệu</returns>
        /////<history> 
        /////</history>
        //public DataSet GetReportData(string SQL)
        //{
        //    try
        //    {
        //        return dal.GetReportData(SQL);
        //    }
        //    catch (DongAException) { throw; }
        //    catch (Exception ex)
        //    {
        //        throw new DongAException(DongAException.DongALayer.Business, ex.Message, ex);
        //    }
        //}

        ///// <summary>
        ///// 
        ///// </summary>
        ///// <returns></returns>
        ///// <history>
        /////     [Truong Lam]   Created [10/06/2020]
        ///// </history>
        //public DbCommand GetListPartner()
        //{
        //    try
        //    {
        //        return new CommonDAL;
        //    }
        //    catch (Exception ex)
        //    {
        //        throw new DongAException(DongALayer.Business, ex.Message, ex);
        //    }

        //}
    }
}

using DongA.Core;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DongA.DataAccess
{
    public class CommonDAL : DongABaseDAL
    {
        ///// <summary>
        /////   Lấy dữ liệu cho báo cáo
        ///// </summary>
        ///// <param name="SQL">Câu sql truy vấn</param>
        ///// <returns>Dataset chứa dữ liệu</returns>
        /////<history>
        /////</history>
        //public DataSet GetReportData(string SQL, List<string> listParam)
        //{
        //    DbCommand command = null;
        //    DataSet result = new DataSet();
        //    try
        //    {

        //        using (command = DongADatabase.GetSqlStringCommand(SQL))
        //        {
        //            foreach(string item in listParam)
        //            {
        //                DongADatabase.AddInParameter(command, item, DbType., item);
        //            }

        //            var reader = DongADatabase.ExecuteDataSet(command, this);
        //            result = reader.Copy();
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        throw DongAException.FromCommand(command, ex);
        //    }
        //}
    }
}

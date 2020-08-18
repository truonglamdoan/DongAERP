using System;
using System.Data;
using System.Data.Common;

namespace DongA.Core
{
    public class DongABaseDAL
    {
        public DongABaseDAL()
        {

        }
         public DongABaseDAL(DbTransaction transaction)
        {
            Transaction = transaction;
        }

        internal DbTransaction Transaction { get; private set; }
    }
}

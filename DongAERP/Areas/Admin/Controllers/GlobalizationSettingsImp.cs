using Aspose.Cells;
using System;

namespace DongAERP.Areas.Admin.Controllers
{
    public class GlobalizationSettingsImp : GlobalizationSettings
    {
        // This function will return the sub total name
        public override String GetTotalName(ConsolidationFunction functionType)
        {
            return "Tổng";
        }

        // This function will return the grand total name
        public override String GetGrandTotalName(ConsolidationFunction functionType)
        {
            return "Tổng tất cả";
        }
    }
}
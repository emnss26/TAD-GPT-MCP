using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace mcp_app.Core
{
    internal class Units
    {
        public static double MetersToFt(double m) =>
            UnitUtils.ConvertToInternalUnits(m, UnitTypeId.Meters);
    }
}

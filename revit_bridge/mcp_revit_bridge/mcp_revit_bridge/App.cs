
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace mcp_revit_bridge
{
    public class App : IExternalApplication
    {
        private RevitBridge _bridge;

        public Result OnStartup(UIControlledApplication application)
        {
            _bridge = new RevitBridge();
            _bridge.Start("http://127.0.0.1:55244/"); // expone /mcp
            return Result.Succeeded;
        }

        public Result OnShutdown(UIControlledApplication application)
        {
            _bridge?.Dispose();
            return Result.Succeeded;
        }
    }
}

using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection.Emit;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using static System.Net.Mime.MediaTypeNames;

namespace mcp_app.Core
{
    internal class ViewHelpers
    {
        public static View ResolveView(UIApplication app, int? viewId)
        {
            var uidoc = app.ActiveUIDocument ?? throw new Exception("No active document.");
            if (viewId is null) return uidoc.ActiveView;
            var v = uidoc.Document.GetElement(new ElementId(viewId.Value)) as View;
            return v ?? throw new Exception($"View {viewId} not found.");
        }

        public static Level ResolveLevel(Document doc, string nameOrNull, View active)
        {
            if (!string.IsNullOrWhiteSpace(nameOrNull))
            {
                var lvl = new FilteredElementCollector(doc).OfClass(typeof(Level)).Cast<Level>()
                    .FirstOrDefault(l => l.Name.Equals(nameOrNull, StringComparison.OrdinalIgnoreCase));
                if (lvl == null) throw new Exception($"Level '{nameOrNull}' not found.");
                return lvl;
            }
            if (active is ViewPlan vp && vp.GenLevel != null) return vp.GenLevel;
            var first = new FilteredElementCollector(doc).OfClass(typeof(Level)).Cast<Level>()
                .OrderBy(l => l.Elevation).FirstOrDefault();
            if (first == null) throw new Exception("No levels available.");
            return first;
        }

        // ✅ Title block correcto: FAMILY SYMBOL
        public static FamilySymbol FindTitleBlock(Document doc, string name)
        {
            return new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>()
                .Where(fs => fs.Category?.Id.IntegerValue == (int)BuiltInCategory.OST_TitleBlocks)
                .FirstOrDefault(fs =>
                    fs.Name.Equals(name, StringComparison.OrdinalIgnoreCase) ||
                    $"{fs.FamilyName}: {fs.Name}".Equals(name, StringComparison.OrdinalIgnoreCase));
        }
    }
}

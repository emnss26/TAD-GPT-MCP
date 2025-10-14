using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using static System.Net.Mime.MediaTypeNames;

namespace mcp_app.Actions
{
    internal class QueryActions
    {
        public static Func<UIApplication, object> LevelsList(JObject _)
        {
            return (UIApplication app) =>
            {
                var doc = app.ActiveUIDocument?.Document ?? throw new Exception("No active document.");
                var items = new FilteredElementCollector(doc).OfClass(typeof(Level)).Cast<Level>()
                    .Select(l => new
                    {
                        id = l.Id.IntegerValue,
                        name = l.Name,
                        elevation_ft = l.Elevation,
                        elevation_m = UnitUtils.ConvertFromInternalUnits(l.Elevation, UnitTypeId.Meters)
                    })
                    .OrderBy(x => x.elevation_ft)
                    .ToList();
                return new { count = items.Count, items };
            };
        }

        public static Func<UIApplication, object> WallTypesList(JObject _)
        {
            return (UIApplication app) =>
            {
                var doc = app.ActiveUIDocument?.Document ?? throw new Exception("No active document.");
                var items = new FilteredElementCollector(doc).OfClass(typeof(WallType)).Cast<WallType>()
                    .Select(t => new
                    {
                        id = t.Id.IntegerValue,
                        family = t.FamilyName,
                        name = t.Name,
                        kind = t.Kind.ToString()
                    })
                    .OrderBy(x => x.family).ThenBy(x => x.name)
                    .ToList();
                return new { count = items.Count, items };
            };
        }

        public static Func<UIApplication, object> ActiveViewInfo(JObject _)
        {
            return (UIApplication app) =>
            {
                var v = app.ActiveUIDocument?.ActiveView ?? throw new Exception("No active view.");
                var vp = v as ViewPlan;
                string levelName = vp?.GenLevel?.Name;
                return new
                {
                    id = v.Id.IntegerValue,
                    name = v.Name,
                    viewType = v.ViewType.ToString(),
                    level = levelName
                };
            };
        }

        public static Func<UIApplication, object> ViewsList(JObject _)
        {
            return (UIApplication app) =>
            {
                var doc = app.ActiveUIDocument?.Document ?? throw new Exception("No active document.");
                var items = new FilteredElementCollector(doc).OfClass(typeof(View)).Cast<View>()
                    .Where(v => !v.IsTemplate)
                    .Select(v => new
                    {
                        id = v.Id.IntegerValue,
                        name = v.Name,
                        viewType = v.ViewType.ToString(),
                        template = (v.ViewTemplateId != ElementId.InvalidElementId)
                    })
                    .OrderBy(x => x.viewType).ThenBy(x => x.name)
                    .ToList();
                return new { count = items.Count, items };
            };
        }

        public static Func<UIApplication, object> SchedulesList(JObject _)
        {
            return (UIApplication app) =>
            {
                var doc = app.ActiveUIDocument?.Document ?? throw new Exception("No active document.");
                var items = new FilteredElementCollector(doc).OfClass(typeof(ViewSchedule)).Cast<ViewSchedule>()
                    .Select(vs => new { id = vs.Id.IntegerValue, name = vs.Name })
                    .OrderBy(x => x.name).ToList();
                return new { count = items.Count, items };
            };
        }

        public static Func<UIApplication, object> MaterialsList(JObject _)
        {
            return (UIApplication app) =>
            {
                var doc = app.ActiveUIDocument?.Document ?? throw new Exception("No active document.");
                var items = new FilteredElementCollector(doc).OfClass(typeof(Material)).Cast<Material>()
                    .Select(m => new { id = m.Id.IntegerValue, name = m.Name })
                    .OrderBy(x => x.name).ToList();
                return new { count = items.Count, items };
            };
        }

        public static Func<UIApplication, object> CategoriesList(JObject _)
        {
            return (UIApplication app) =>
            {
                var doc = app.ActiveUIDocument?.Document ?? throw new Exception("No active document.");

                var items = doc.Settings.Categories.Cast<Category>()
                    .Select(c =>
                    {
                        int bicId = c.Id.IntegerValue;
                        string bicName = Enum.IsDefined(typeof(BuiltInCategory), bicId)
                            ? ((BuiltInCategory)bicId).ToString()
                            : null;
                        return new
                        {
                            id = bicId,
                            name = c.Name,
                            bicId,
                            bicName
                        };
                    })
                    .OrderBy(x => x.name)
                    .ToList();

                return new { count = items.Count, items };
            };
        }

        public static Func<UIApplication, object> FamiliesTypesList(JObject _)
        {
            return (UIApplication app) =>
            {
                var doc = app.ActiveUIDocument?.Document ?? throw new Exception("No active document.");
                var items = new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>()
                    .Select(fs => new
                    {
                        id = fs.Id.IntegerValue,
                        family = fs.FamilyName,
                        type = fs.Name,
                        category = fs.Category?.Name
                    })
                    .OrderBy(x => x.category).ThenBy(x => x.family).ThenBy(x => x.type).ToList();
                return new { count = items.Count, items };
            };
        }

        public static Func<UIApplication, object> LinksList(JObject _)
        {
            return (UIApplication app) =>
            {
                var doc = app.ActiveUIDocument?.Document ?? throw new Exception("No active document.");
                var items = new FilteredElementCollector(doc).OfClass(typeof(RevitLinkInstance)).Cast<RevitLinkInstance>()
                    .Select(li => new { id = li.Id.IntegerValue, name = li.Name, pinned = li.Pinned })
                    .OrderBy(x => x.name).ToList();
                return new { count = items.Count, items };
            };
        }

        public static Func<UIApplication, object> ImportsList(JObject _)
        {
            return (UIApplication app) =>
            {
                var doc = app.ActiveUIDocument?.Document ?? throw new Exception("No active document.");
                var items = new FilteredElementCollector(doc).OfClass(typeof(ImportInstance)).Cast<ImportInstance>()
                    .Select(ii => new { id = ii.Id.IntegerValue, name = ii.Name })
                    .OrderBy(x => x.name).ToList();
                return new { count = items.Count, items };
            };
        }

        public static Func<UIApplication, object> WorksetsList(JObject _)
        {
            return (UIApplication app) =>
            {
                var doc = app.ActiveUIDocument?.Document ?? throw new Exception("No active document.");
                var all = new FilteredWorksetCollector(doc).ToWorksets();
                var items = all
                    .Select(ws => new
                    {
                        id = ws.Id.IntegerValue,
                        name = ws.Name,
                        kind = ws.Kind.ToString()
                    })
                    .OrderBy(x => x.kind).ThenBy(x => x.name)
                    .ToList();
                return new { count = items.Count, items };
            };
        }

        public static Func<UIApplication, object> TextNotesFind(JObject _)
        {
            return (UIApplication app) =>
            {
                var v = app.ActiveUIDocument?.ActiveView ?? throw new Exception("No active view.");
                var items = new FilteredElementCollector(v.Document, v.Id)
                    .OfClass(typeof(TextNote)).Cast<TextNote>()
                    .Select(tn => new { id = tn.Id.IntegerValue, text = tn.Text })
                    .ToList();
                return new { viewId = v.Id.IntegerValue, count = items.Count, items };
            };
        }

        public static Func<UIApplication, object> DuctTypesList(JObject _)
        {
            return (UIApplication app) =>
            {
                var doc = app.ActiveUIDocument?.Document ?? throw new Exception("No active document.");
                var items = new FilteredElementCollector(doc).OfClass(typeof(DuctType)).Cast<DuctType>()
                    .Select(t => new { id = t.Id.IntegerValue, family = t.FamilyName, name = t.Name })
                    .OrderBy(x => x.family).ThenBy(x => x.name).ToList();
                return new { count = items.Count, items };
            };
        }

        public static Func<UIApplication, object> PipeTypesList(JObject _)
        {
            return (UIApplication app) =>
            {
                var doc = app.ActiveUIDocument?.Document ?? throw new Exception("No active document.");
                var items = new FilteredElementCollector(doc).OfClass(typeof(PipeType)).Cast<PipeType>()
                    .Select(t => new { id = t.Id.IntegerValue, family = t.FamilyName, name = t.Name })
                    .OrderBy(x => x.family).ThenBy(x => x.name).ToList();
                return new { count = items.Count, items };
            };
        }

        public static Func<UIApplication, object> CableTrayTypesList(JObject _)
        {
            return (UIApplication app) =>
            {
                var doc = app.ActiveUIDocument?.Document ?? throw new Exception("No active document.");

                var items = new FilteredElementCollector(doc)
                    .OfClass(typeof(CableTrayType))
                    .Cast<CableTrayType>()
                    .Select(t => new
                    {
                        id = t.Id.IntegerValue,
                        name = t.Name
                        // family = t.FamilyName // si te interesa y tu versión lo expone
                    })
                    .OrderBy(x => x.name)
                    .ToList();

                return new { count = items.Count, items };
            };
        }

        private static object ParamToJson(Parameter p)
        {
            if (p == null) return null;
            string name = p.Definition?.Name ?? p.Id.IntegerValue.ToString();
            object val = null;
            try
            {
                switch (p.StorageType)
                {
                    case StorageType.Double: val = p.AsDouble(); break;
                    case StorageType.Integer: val = p.AsInteger(); break;
                    case StorageType.String: val = p.AsString(); break;
                    case StorageType.ElementId: val = p.AsElementId()?.IntegerValue; break;
                }
            }
            catch { }

            string valueString = null;
            try { valueString = p.AsValueString(); } catch { }

            return new
            {
                id = p.Id.IntegerValue,
                name,
                storageType = p.StorageType.ToString(),
                value = val,
                valueString,
                isReadOnly = p.IsReadOnly
            };
        }

        public static Func<UIApplication, object> SelectionInfo(JObject args)
        {
            bool includeParams = args?["includeParameters"]?.ToObject<bool?>() ?? true;
            int topN = args?["topNParams"]?.ToObject<int?>() ?? 50;

            return (UIApplication app) =>
            {
                var uidoc = app.ActiveUIDocument ?? throw new Exception("No active document.");
                var doc = uidoc.Document;
                var sel = uidoc.Selection.GetElementIds();
                var items = new List<object>();

                foreach (var id in sel)
                {
                    var el = doc.GetElement(id);
                    items.Add(BuildElementInfo(doc, el, includeParams, topN));
                }

                return new { count = items.Count, items };
            };
        }

        public static Func<UIApplication, object> ElementInfo(JObject args)
        {
            int elementId = args?["elementId"]?.ToObject<int>() ?? 0;
            bool includeParams = args?["includeParameters"]?.ToObject<bool?>() ?? true;
            int topN = args?["topNParams"]?.ToObject<int?>() ?? 50;

            if (elementId <= 0)
                throw new Exception("element.info requires a valid elementId.");

            return (UIApplication app) =>
            {
                var doc = app.ActiveUIDocument?.Document ?? throw new Exception("No active document.");
                var el = doc.GetElement(new ElementId(elementId)) ?? throw new Exception($"Element {elementId} not found.");
                return BuildElementInfo(doc, el, includeParams, topN);
            };
        }

        private static double? GetDoubleParam(Element e, BuiltInParameter bip)
        {
            var p = e?.get_Parameter(bip);
            if (p == null || p.StorageType != StorageType.Double) return null;
            try { return p.AsDouble(); } catch { return null; }
        }

        private static object BuildElementInfo(Document doc, Element el, bool includeParams, int topNParams)
        {
            if (el == null) return null;

            var cat = el.Category?.Name;

            var typeId = el.GetTypeId();
            var typeName = (typeId != ElementId.InvalidElementId) ? doc.GetElement(typeId)?.Name : null;

            // Mejor esfuerzo para conseguir nivel
            var levelId = (el as FamilyInstance)?.LevelId;
            if (levelId == null || levelId == ElementId.InvalidElementId)
                levelId = el.LevelId;
            var levelName = (levelId != null && levelId != ElementId.InvalidElementId)
                ? doc.GetElement(levelId)?.Name
                : null;

            // BBox
            var bb = el.get_BoundingBox(null);
            object bbox = (bb != null)
                ? new
                {
                    min = new { x = bb.Min.X, y = bb.Min.Y, z = bb.Min.Z },
                    max = new { x = bb.Max.X, y = bb.Max.Y, z = bb.Max.Z }
                }
                : null;

            // Métricas comunes (mejor esfuerzo)
            double? length_ft = null, area_ft2 = null, vol_ft3 = null;

            // length: si hay LocationCurve
            var lc = (el.Location as LocationCurve);
            if (lc?.Curve != null) length_ft = lc.Curve.Length;

            // área/volumen típicos
            area_ft2 = GetDoubleParam(el, BuiltInParameter.HOST_AREA_COMPUTED);
            vol_ft3 = GetDoubleParam(el, BuiltInParameter.HOST_VOLUME_COMPUTED);

            // Convertir a SI
            double? length_m = length_ft.HasValue ? UnitUtils.ConvertFromInternalUnits(length_ft.Value, UnitTypeId.Meters) : (double?)null;
            double? area_m2 = area_ft2.HasValue ? UnitUtils.ConvertFromInternalUnits(area_ft2.Value, UnitTypeId.SquareMeters) : (double?)null;
            double? volume_m3 = vol_ft3.HasValue ? UnitUtils.ConvertFromInternalUnits(vol_ft3.Value, UnitTypeId.CubicMeters) : (double?)null;

            object parameters = null;
            if (includeParams)
            {
                var ps = el.Parameters.Cast<Parameter>()
                    .Take(topNParams <= 0 ? int.MaxValue : topNParams)
                    .Select(ParamToJson)
                    .ToList();
                parameters = ps;
            }

            return new
            {
                elementId = el.Id.IntegerValue,
                category = cat,
                typeName,
                level = levelName,
                name = el.Name,
                bbox,
                metrics = new { length_m, area_m2, volume_m3 },
                parameters
            };
        }
    }
}

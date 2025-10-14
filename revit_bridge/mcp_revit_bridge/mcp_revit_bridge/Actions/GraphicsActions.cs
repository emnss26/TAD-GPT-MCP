using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using mcp_app.Actions;
using mcp_app.Contracts;
using mcp_app.Core;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static System.Net.Mime.MediaTypeNames;

namespace mcp_app.Actions
{
    internal class GraphicsActions
    {
        public static System.Func<UIApplication, object> SetVisibility(JObject args)
        {
            var categories = args["categories"]?.ToObject<List<string>>() ?? new List<string>();
            bool visible = args["visible"]?.ToObject<bool>() ?? true;
            bool forceDetach = args["forceDetachTemplate"]?.ToObject<bool>() ?? false;
            int? viewId = args["viewId"]?.ToObject<int?>();

            return (UIApplication app) =>
            {
                var v = ResolveView(app, viewId);
                EnsureTemplateWritable(v, forceDetach);

                using (var t = new Transaction(v.Document, "MCP: Set Category Visibility"))
                {
                    t.Start();
                    foreach (var tok in categories)
                    {
                        var cat = CategoryLookup.Require(v.Document, tok);
                        v.SetCategoryHidden(cat.Id, !visible);
                    }
                    t.Commit();
                }
                return new { changed = categories.Count };
            };
        }

        public static System.Func<UIApplication, object> ClearOverrides(JObject args)
        {
            var categories = args["categories"]?.ToObject<List<string>>() ?? new List<string>();
            bool forceDetach = args["forceDetachTemplate"]?.ToObject<bool>() ?? false;
            int? viewId = args["viewId"]?.ToObject<int?>();

            return (UIApplication app) =>
            {
                var v = ResolveView(app, viewId);
                EnsureTemplateWritable(v, forceDetach);

                using (var t = new Transaction(v.Document, "MCP: Clear Category Overrides"))
                {
                    t.Start();
                    foreach (var tok in categories)
                    {
                        var cat = CategoryLookup.Require(v.Document, tok);
                        v.SetCategoryOverrides(cat.Id, new OverrideGraphicSettings());
                    }
                    t.Commit();
                }
                return new { cleared = categories.Count };
            };
        }

        public static System.Func<UIApplication, object> OverrideColor(JObject args)
        {
            var categories = args["categories"]?.ToObject<List<string>>() ?? new List<string>();
            var colorAny = args["color"]?.ToObject<object>();
            int transparency = args["transparency"]?.ToObject<int?>() ?? 0;
            bool halftone = args["halftone"]?.ToObject<bool?>() ?? false;
            bool surfaceSolid = args["surfaceSolid"]?.ToObject<bool?>() ?? true;
            bool projectionLines = args["projectionLines"]?.ToObject<bool?>() ?? false;
            bool forceDetach = args["forceDetachTemplate"]?.ToObject<bool?>() ?? false;
            int? viewId = args["viewId"]?.ToObject<int?>();

            return (UIApplication app) =>
            {
                var v = ResolveView(app, viewId);
                EnsureTemplateWritable(v, forceDetach);
                var color = ColorParser.FromAny(colorAny);

                using (var t = new Transaction(v.Document, "MCP: Override Category Color"))
                {
                    t.Start();

                    var ogs = new OverrideGraphicSettings();

                    // Líneas
                    if (projectionLines)
                        ogs.SetProjectionLineColor(color);

                    // Superficie (solid fill)
                    if (surfaceSolid)
                    {
                        var solid = FindSolidFill(v.Document);
                        if (solid != null)
                        {
                            ogs.SetSurfaceForegroundPatternId(solid.Id);
                            ogs.SetSurfaceForegroundPatternColor(color);
                        }
                    }

                    if (transparency >= 0 && transparency <= 100)
                        ogs.SetSurfaceTransparency(transparency);

                    ogs.SetHalftone(halftone);

                    foreach (var tok in categories)
                    {
                        var cat = CategoryLookup.Require(v.Document, tok);
                        v.SetCategoryOverrides(cat.Id, ogs);
                    }
                    t.Commit();
                }
                return new { overridden = categories.Count, color = new { } };
            };
        }

        // ===== Helpers =====
        private static View ResolveView(UIApplication app, int? viewId)
        {
            var uidoc = app.ActiveUIDocument ?? throw new System.Exception("No active document.");
            if (viewId is null) return uidoc.ActiveView;
            var v = uidoc.Document.GetElement(new ElementId(viewId.Value)) as View;
            return v ?? throw new System.Exception($"View {viewId} not found.");
        }

        private static void EnsureTemplateWritable(View v, bool forceDetach)
        {
            if (v.ViewTemplateId != ElementId.InvalidElementId)
            {
                if (!forceDetach)
                    throw new System.Exception("View has a template; set forceDetachTemplate=true to detach.");
                v.ViewTemplateId = ElementId.InvalidElementId;
            }
        }

        private static FillPatternElement FindSolidFill(Document doc)
        {
            return new FilteredElementCollector(doc)
                .OfClass(typeof(FillPatternElement))
                .Cast<FillPatternElement>()
                .FirstOrDefault(f => f.GetFillPattern()?.IsSolidFill == true
                    || f.Name.Equals("Solid fill", System.StringComparison.OrdinalIgnoreCase)
                    || f.Name.Equals("Relleno sólido", System.StringComparison.OrdinalIgnoreCase));
        }

        public static Func<UIApplication, object> ApplyTemplate(JObject args)
        {
            var req = args.ToObject<ViewApplyTemplateRequest>() ?? new ViewApplyTemplateRequest();
            return (UIApplication app) =>
            {
                var v = Core.ViewHelpers.ResolveView(app, req.viewId);
                var doc = v.Document;

                View viewTemplate = null;
                if (req.templateId is int tid && tid > 0)
                    viewTemplate = doc.GetElement(new ElementId(tid)) as View;
                else if (!string.IsNullOrWhiteSpace(req.templateName))
                {
                    viewTemplate = new FilteredElementCollector(doc).OfClass(typeof(View)).Cast<View>()
                        .FirstOrDefault(t => t.IsTemplate && t.Name.Equals(req.templateName, StringComparison.OrdinalIgnoreCase));
                }
                if (viewTemplate == null) throw new Exception("View template not found.");

                using (var t = new Transaction(doc, "MCP: View.ApplyTemplate"))
                {
                    t.Start();
                    v.ViewTemplateId = viewTemplate.Id;
                    t.Commit();
                }
                return new { viewId = v.Id.IntegerValue, template = viewTemplate.Name };
            };
        }

        public static Func<UIApplication, object> SetScale(JObject args)
        {
            var req = args.ToObject<ViewSetScaleRequest>() ?? throw new Exception("Invalid args for view.set_scale.");
            return (UIApplication app) =>
            {
                var v = Core.ViewHelpers.ResolveView(app, req.viewId);
                using (var t = new Transaction(v.Document, "MCP: View.SetScale"))
                {
                    t.Start();
                    v.Scale = Math.Max(1, req.scale);
                    t.Commit();
                }
                return new { viewId = v.Id.IntegerValue, scale = v.Scale };
            };
        }

        public static Func<UIApplication, object> SetDetailLevel(JObject args)
        {
            var req = args.ToObject<ViewSetDetailRequest>() ?? throw new Exception("Invalid args for view.set_detail_level.");
            return (UIApplication app) =>
            {
                var v = Core.ViewHelpers.ResolveView(app, req.viewId);
                var map = new Dictionary<string, ViewDetailLevel>(StringComparer.OrdinalIgnoreCase)
                {
                    ["coarse"] = ViewDetailLevel.Coarse,
                    ["medium"] = ViewDetailLevel.Medium,
                    ["fine"] = ViewDetailLevel.Fine
                };
                if (!map.TryGetValue(req.detailLevel ?? "medium", out var lvl)) lvl = ViewDetailLevel.Medium;

                using (var t = new Transaction(v.Document, "MCP: View.SetDetailLevel"))
                {
                    t.Start();
                    v.DetailLevel = lvl;
                    t.Commit();
                }
                return new { viewId = v.Id.IntegerValue, detailLevel = v.DetailLevel.ToString() };
            };
        }

        public static Func<UIApplication, object> SetDiscipline(JObject args)
        {
            var req = args.ToObject<ViewSetDisciplineRequest>() ?? throw new Exception("Invalid args for view.set_discipline.");
            return (UIApplication app) =>
            {
                var v = Core.ViewHelpers.ResolveView(app, req.viewId);
                var map = new Dictionary<string, ViewDiscipline>(StringComparer.OrdinalIgnoreCase)
                {
                    ["architectural"] = ViewDiscipline.Architectural,
                    ["structural"] = ViewDiscipline.Structural,
                    ["mechanical"] = ViewDiscipline.Mechanical,
                    ["coordination"] = ViewDiscipline.Coordination
                };
                if (!map.TryGetValue(req.discipline ?? "coordination", out var disc)) disc = ViewDiscipline.Coordination;

                using (var t = new Transaction(v.Document, "MCP: View.SetDiscipline"))
                {
                    t.Start();
                    v.Discipline = disc;
                    t.Commit();
                }
                return new { viewId = v.Id.IntegerValue, discipline = v.Discipline.ToString() };
            };
        }

        public static Func<UIApplication, object> SetPhase(JObject args)
        {
            var req = args.ToObject<ViewSetPhaseRequest>() ?? throw new Exception("Invalid args for view.set_phase.");
            return (UIApplication app) =>
            {
                var v = mcp_app.Core.ViewHelpers.ResolveView(app, req.viewId);
                var doc = v.Document;

                // Usamos doc.Phases (PhaseArray) para obtener orden de proyecto
                Phase target = null;
                if (!string.IsNullOrWhiteSpace(req.phase))
                {
                    foreach (Phase ph in doc.Phases)
                    {
                        if (ph.Name.Equals(req.phase, StringComparison.OrdinalIgnoreCase))
                        { target = ph; break; }
                    }
                    if (target == null) throw new Exception($"Phase '{req.phase}' not found.");
                }
                else
                {
                    if (doc.Phases == null || doc.Phases.Size == 0) throw new Exception("No phases available.");
                    target = doc.Phases.get_Item(doc.Phases.Size - 1); // última fase
                }

                using (var t = new Transaction(doc, "MCP: View.SetPhase"))
                {
                    t.Start();
                    var p = v.get_Parameter(BuiltInParameter.VIEW_PHASE);
                    if (p == null || p.IsReadOnly) throw new Exception("VIEW_PHASE parameter not settable for this view.");
                    p.Set(target.Id);
                    t.Commit();
                }
                return new { viewId = v.Id.IntegerValue, phase = target.Name };
            };
        }

        public static Func<UIApplication, object> ViewsDuplicate(JObject args)
        {
            var req = args.ToObject<ViewsDuplicateRequest>() ?? throw new Exception("Invalid args for views.duplicate.");
            return (UIApplication app) =>
            {
                var doc = app.ActiveUIDocument?.Document ?? throw new Exception("No active document.");
                var mode = (req.mode ?? "duplicate").ToLowerInvariant();
                var opt = ViewDuplicateOption.Duplicate;
                if (mode == "with_detailing") opt = ViewDuplicateOption.WithDetailing;
                else if (mode == "as_dependent") opt = ViewDuplicateOption.AsDependent;

                var created = new List<int>();
                using (var t = new Transaction(doc, "MCP: Views.Duplicate"))
                {
                    t.Start();
                    foreach (var id in req.viewIds ?? Array.Empty<int>())
                    {
                        var v = doc.GetElement(new ElementId(id)) as View;
                        if (v == null) continue;
                        var nid = v.Duplicate(opt);
                        created.Add(nid.IntegerValue);
                    }
                    t.Commit();
                }
                return new { created };
            };
        }

        public static Func<UIApplication, object> HideImports(JObject args)
        {
            var req = args.ToObject<ViewIdArg>() ?? new ViewIdArg();
            return (UIApplication app) =>
            {
                var v = Core.ViewHelpers.ResolveView(app, req.viewId);
                var doc = v.Document;
                var imports = new FilteredElementCollector(doc, v.Id)
                    .OfClass(typeof(ImportInstance)).ToElementIds();

                using (var t = new Transaction(doc, "MCP: Imports.Hide"))
                {
                    t.Start();
                    if (imports.Count > 0) v.HideElements(imports);
                    t.Commit();
                }
                return new { hidden = imports.Count };
            };
        }

        public static Func<UIApplication, object> SetDisplayStyle(JObject args)
        {
            var styleStrRaw = args["displayStyle"]?.ToObject<string>() ?? "HiddenLine";
            bool forceDetach = args["forceDetachTemplate"]?.ToObject<bool?>() ?? false;
            int? viewId = args["viewId"]?.ToObject<int?>();

            return (UIApplication app) =>
            {
                var v = ResolveView(app, viewId);
                EnsureTemplateWritable(v, forceDetach);

                // Obtener la propiedad real del API (evita conflictos con otros DisplayStyle)
                var dsProp = typeof(View).GetProperty("DisplayStyle");
                if (dsProp == null || !dsProp.CanWrite)
                    throw new Exception("Esta versión del API no permite setear View.DisplayStyle.");

                var enumType = dsProp.PropertyType;               // enum real
                var available = Enum.GetNames(enumType);          // nombres disponibles en esta build

                // Normaliza la entrada
                var key = (styleStrRaw ?? "").Replace(" ", "").ToLowerInvariant();

                // Aliases -> nombre oficial del enum (sin usar switch expression)
                var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "wireframe",              "Wireframe" },
            { "hiddenline",             "HiddenLine" },
            { "hidden_line",            "HiddenLine" },
            { "hiddenlines",            "HiddenLine" },
            { "shaded",                 "Shaded" },
            { "shadedwithedges",        "Shaded" },           // bordes se controlan aparte
            { "consistentcolors",       "ConsistentColors" },
            { "realistic",              "Realistic" },
            { "realisticwithedges",     "Realistic" }         // bordes se controlan aparte
        };

                string wanted;
                if (!map.TryGetValue(key, out wanted))
                    wanted = "HiddenLine";

                // Si el valor deseado no existe en esta build, usar un fallback razonable
                string finalName = null;

                // 1) ¿Existe exactamente el pedido?
                foreach (var name in available)
                {
                    if (string.Equals(name, wanted, StringComparison.OrdinalIgnoreCase))
                    {
                        finalName = name;
                        break;
                    }
                }

                // 2) Fallback por preferencia
                if (finalName == null)
                {
                    var fallbackPref = new[] { "HiddenLine", "Wireframe", "Shaded", "ConsistentColors", "Realistic" };
                    foreach (var pref in fallbackPref)
                    {
                        foreach (var name in available)
                        {
                            if (string.Equals(name, pref, StringComparison.OrdinalIgnoreCase))
                            {
                                finalName = name;
                                break;
                            }
                        }
                        if (finalName != null) break;
                    }
                }

                // 3) Último recurso
                if (finalName == null && available.Length > 0)
                    finalName = available[0];

                var enumValue = Enum.Parse(enumType, finalName, true);

                using (var t = new Transaction(v.Document, "MCP: View.SetDisplayStyle"))
                {
                    t.Start();
                    dsProp.SetValue(v, enumValue);
                    t.Commit();
                }

                return new { viewId = v.Id.IntegerValue, displayStyle = finalName };
            };
        }

        public static Func<UIApplication, object> LinksHideInView(JObject args)
        {
            int? viewId = args["viewId"]?.ToObject<int?>();
            return (UIApplication app) =>
            {
                var v = ResolveView(app, viewId);
                var doc = v.Document;

                var links = new FilteredElementCollector(doc, v.Id)
                    .OfClass(typeof(RevitLinkInstance)).ToElementIds();

                using (var t = new Transaction(doc, "MCP: Links.HideInView"))
                {
                    t.Start();
                    if (links.Count > 0) v.HideElements(links);
                    t.Commit();
                }
                return new { hidden = links.Count, viewId = v.Id.IntegerValue };
            };
        }

        public static Func<UIApplication, object> LinksHalftoneInView(JObject args)
        {
            bool halftone = args["halftone"]?.ToObject<bool?>() ?? true;
            int? viewId = args["viewId"]?.ToObject<int?>();

            return (UIApplication app) =>
            {
                var v = ResolveView(app, viewId);
                var doc = v.Document;

                var links = new FilteredElementCollector(doc, v.Id)
                    .OfClass(typeof(RevitLinkInstance)).Cast<Element>().ToList();

                int changed = 0;
                using (var t = new Transaction(doc, "MCP: Links.HalftoneInView"))
                {
                    t.Start();
                    foreach (var e in links)
                    {
                        var ogs = v.GetElementOverrides(e.Id) ?? new OverrideGraphicSettings();
                        ogs.SetHalftone(halftone);
                        v.SetElementOverrides(e.Id, ogs);
                        changed++;
                    }
                    t.Commit();
                }
                return new { changed, halftone, viewId = v.Id.IntegerValue };
            };
        }

    }
}


using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using mcp_app.Core;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using static System.Net.Mime.MediaTypeNames;

namespace mcp_app.Actions
{
    internal class MepPipeConduitActions
    {
        private class Pt2 { public double x; public double y; }
        private static double ToFt_m(double m) => Core.Units.MetersToFt(m);

        // ===== Helpers comunes (conectores) =====
        private static IEnumerable<Connector> CollectConnectors(Element e)
        {
            var list = new List<Connector>();
            if (e is MEPCurve mc)
            {
                var set = mc.ConnectorManager?.Connectors;
                if (set != null) foreach (Connector c in set) list.Add(c);
            }
            else if (e is FamilyInstance fi)
            {
                var cm = fi.MEPModel?.ConnectorManager;
                if (cm?.Connectors != null) foreach (Connector c in cm.Connectors) list.Add(c);
            }
            return list;
        }
        private static (Connector a, Connector b) FindClosestPair(IEnumerable<Connector> aa, IEnumerable<Connector> bb)
        {
            Connector bestA = null, bestB = null;
            double best = double.MaxValue;
            foreach (var a in aa)
                foreach (var b in bb)
                {
                    var d = a.Origin.DistanceTo(b.Origin);
                    if (d < best) { best = d; bestA = a; bestB = b; }
                }
            return (bestA, bestB);
        }
        private static bool Nearly(double x, double y, double eps = 1e-6) => Math.Abs(x - y) <= eps;
        private static bool SameSize(Connector c1, Connector c2)
        {
            try
            {
                if (c1.Shape == ConnectorProfileType.Round && c2.Shape == ConnectorProfileType.Round)
                    return Nearly(c1.Radius, c2.Radius);
                bool w = c1.Width > 0 && c2.Width > 0 ? Nearly(c1.Width, c2.Width) : true;
                bool h = c1.Height > 0 && c2.Height > 0 ? Nearly(c1.Height, c2.Height) : true;
                return w && h;
            }
            catch { return false; }
        }
        private static bool IsFreeEnd(Connector c) { try { return c != null && !c.IsConnected; } catch { return true; } }
        private static Element Require(Document doc, int id, string label)
        {
            var e = doc.GetElement(new ElementId(id));
            if (e == null) throw new Exception($"{label} {id} not found.");
            return e;
        }
        private static bool TrySetDoubleParamByNames(Element e, string[] names, double val)
        {
            foreach (Parameter p in e.Parameters)
            {
                var n = p.Definition?.Name;
                if (n == null) continue;
                foreach (var want in names)
                {
                    if (n.Equals(want, StringComparison.OrdinalIgnoreCase))
                    {
                        if (!p.IsReadOnly && p.StorageType == StorageType.Double)
                        { p.Set(val); return true; }
                    }
                }
            }
            return false;
        }

        // ===== Compat (reflexión) para fittings de PIPE =====
        private static FamilyInstance CreatePipeElbowFittingCompat(Document doc, Connector c1, Connector c2)
        {
            var asm = typeof(Pipe).Assembly;
            var t = asm.GetType("Autodesk.Revit.DB.Plumbing.PlumbingUtils");
            if (t != null)
            {
                var m1 = t.GetMethod("CreateElbowFitting",
                    BindingFlags.Public | BindingFlags.Static, null,
                    new Type[] { typeof(Document), typeof(Connector), typeof(Connector) }, null);
                if (m1 != null) return m1.Invoke(null, new object[] { doc, c1, c2 }) as FamilyInstance;

                var m2 = t.GetMethod("CreateElbowFitting",
                    BindingFlags.Public | BindingFlags.Static, null,
                    new Type[] { typeof(Connector), typeof(Connector) }, null);
                if (m2 != null) return m2.Invoke(null, new object[] { c1, c2 }) as FamilyInstance;
            }
            throw new Exception("PlumbingUtils.CreateElbowFitting no disponible en esta versión del API.");
        }
        private static FamilyInstance CreatePipeTransitionFittingCompat(Document doc, Connector c1, Connector c2)
        {
            var asm = typeof(Pipe).Assembly;
            var t = asm.GetType("Autodesk.Revit.DB.Plumbing.PlumbingUtils");
            if (t != null)
            {
                var m1 = t.GetMethod("CreateTransitionFitting",
                    BindingFlags.Public | BindingFlags.Static, null,
                    new Type[] { typeof(Document), typeof(Connector), typeof(Connector) }, null);
                if (m1 != null) return m1.Invoke(null, new object[] { doc, c1, c2 }) as FamilyInstance;

                var m2 = t.GetMethod("CreateTransitionFitting",
                    BindingFlags.Public | BindingFlags.Static, null,
                    new Type[] { typeof(Connector), typeof(Connector) }, null);
                if (m2 != null) return m2.Invoke(null, new object[] { c1, c2 }) as FamilyInstance;
            }
            throw new Exception("PlumbingUtils.CreateTransitionFitting no disponible en esta versión del API.");
        }

        // ===== Compat (reflexión) para PipeInsulation =====
        private static ElementId CreatePipeInsulationCompat(Document doc, ElementId curveId, ElementId typeId, double? thicknessFt, out Element created)
        {
            created = null;
            var t = typeof(PipeInsulation);
            var mNew = t.GetMethod("Create",
                BindingFlags.Public | BindingFlags.Static, null,
                new Type[] { typeof(Document), typeof(ElementId), typeof(ElementId), typeof(double) }, null);
            if (mNew != null)
            {
                created = mNew.Invoke(null, new object[] { doc, curveId, typeId, thicknessFt ?? 0.0 }) as Element;
                if (created == null) throw new Exception("PipeInsulation.Create devolvió null.");
                return created.Id;
            }
            var mOld = t.GetMethod("Create",
                BindingFlags.Public | BindingFlags.Static, null,
                new Type[] { typeof(Document), typeof(ElementId), typeof(ElementId) }, null);
            if (mOld != null)
            {
                created = mOld.Invoke(null, new object[] { doc, curveId, typeId }) as Element;
                if (created == null) throw new Exception("PipeInsulation.Create devolvió null.");
                return created.Id;
            }
            throw new Exception("No se encontró PipeInsulation.Create (firmas soportadas).");
        }

        // =========================
        // 0) Pipe.Create
        // =========================
        private class PipeCreateRequest
        {
            public string level { get; set; }
            public string systemType { get; set; }      // nombre (PipingSystemType) opcional
            public string pipeType { get; set; }        // opcional
            public double elevation_m { get; set; } = 2.5;
            public Pt2 start { get; set; }
            public Pt2 end { get; set; }
            public double? diameter_mm { get; set; }    // opcional
        }

        public static Func<UIApplication, object> PipeCreate(JObject args)
        {
            var req = args.ToObject<PipeCreateRequest>() ?? throw new Exception("Invalid args for mep.pipe.create.");
            return (UIApplication app) =>
            {
                var uidoc = app.ActiveUIDocument ?? throw new Exception("No active document.");
                var doc = uidoc.Document;
                var active = uidoc.ActiveView;

                var level = ViewHelpers.ResolveLevel(doc, req.level, active);
                double z = level.Elevation + Core.Units.MetersToFt(req.elevation_m);

                var systems = new FilteredElementCollector(doc).OfClass(typeof(PipingSystemType)).Cast<PipingSystemType>().ToList();
                var sys = string.IsNullOrWhiteSpace(req.systemType)
                    ? systems.FirstOrDefault()
                    : systems.FirstOrDefault(s => s.Name.Equals(req.systemType, StringComparison.OrdinalIgnoreCase));
                if (sys == null) throw new Exception("No PipingSystemType found.");

                var types = new FilteredElementCollector(doc).OfClass(typeof(PipeType)).Cast<PipeType>().ToList();
                PipeType ptype = null;
                if (!string.IsNullOrWhiteSpace(req.pipeType))
                {
                    ptype = types.FirstOrDefault(t =>
                        t.Name.Equals(req.pipeType, StringComparison.OrdinalIgnoreCase) ||
                        $"{t.FamilyName}: {t.Name}".Equals(req.pipeType, StringComparison.OrdinalIgnoreCase));
                }
                if (ptype == null) ptype = types.FirstOrDefault();
                if (ptype == null)
                {
                    var sample = string.Join(", ", types.Take(10).Select(t => $"{t.FamilyName}: {t.Name}"));
                    throw new Exception("No PipeType found." + (sample.Length > 0 ? $" Available examples: {sample}" : ""));
                }

                double ToFt(double m) => Core.Units.MetersToFt(m);
                var p1 = new XYZ(ToFt(req.start.x), ToFt(req.start.y), z);
                var p2 = new XYZ(ToFt(req.end.x), ToFt(req.end.y), z);

                int id;
                using (var t = new Transaction(doc, "MCP: MEP.Pipe.Create"))
                {
                    t.Start();
                    var pipe = Pipe.Create(doc, sys.Id, ptype.Id, level.Id, p1, p2);
                    id = pipe.Id.IntegerValue;

                    if (req.diameter_mm.HasValue)
                    {
                        var diam = Core.Units.MetersToFt(req.diameter_mm.Value / 1000.0);
                        pipe.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM)?.Set(diam);
                    }

                    t.Commit();
                }

                return new
                {
                    elementId = id,
                    used = new
                    {
                        level = level.Name,
                        systemType = sys.Name,
                        pipeType = $"{ptype.FamilyName}: {ptype.Name}",
                        elevation_m = req.elevation_m
                    }
                };
            };
        }

        // =========================
        // 1) Conduit.Create
        // =========================
        private class ConduitCreateRequest
        {
            public string level { get; set; }
            public string conduitType { get; set; }     // opcional
            public double elevation_m { get; set; } = 2.4;
            public Pt2 start { get; set; }
            public Pt2 end { get; set; }
            public double? diameter_mm { get; set; }    // opcional
        }

        public static Func<UIApplication, object> ConduitCreate(JObject args)
        {
            var req = args.ToObject<ConduitCreateRequest>() ?? throw new Exception("Invalid args for mep.conduit.create.");
            return (UIApplication app) =>
            {
                var uidoc = app.ActiveUIDocument ?? throw new Exception("No active document.");
                var doc = uidoc.Document;
                var active = uidoc.ActiveView;

                var level = ViewHelpers.ResolveLevel(doc, req.level, active);
                double z = level.Elevation + Core.Units.MetersToFt(req.elevation_m);

                var ctypes = new FilteredElementCollector(doc).OfClass(typeof(ConduitType)).Cast<ConduitType>().ToList();
                ConduitType ctype = null;
                if (!string.IsNullOrWhiteSpace(req.conduitType))
                {
                    ctype = ctypes.FirstOrDefault(t =>
                        t.Name.Equals(req.conduitType, StringComparison.OrdinalIgnoreCase) ||
                        $"{t.FamilyName}: {t.Name}".Equals(req.conduitType, StringComparison.OrdinalIgnoreCase));
                }
                if (ctype == null) ctype = ctypes.FirstOrDefault();
                if (ctype == null)
                {
                    var sample = string.Join(", ", ctypes.Take(10).Select(t => t.Name));
                    throw new Exception("No ConduitType found." + (sample.Length > 0 ? $" Available examples: {sample}" : ""));
                }

                double ToFt(double m) => Core.Units.MetersToFt(m);
                var p1 = new XYZ(ToFt(req.start.x), ToFt(req.start.y), z);
                var p2 = new XYZ(ToFt(req.end.x), ToFt(req.end.y), z);

                int id;
                using (var t = new Transaction(doc, "MCP: MEP.Conduit.Create"))
                {
                    t.Start();
                    var conduit = Conduit.Create(doc, ctype.Id, p1, p2, level.Id);
                    id = conduit.Id.IntegerValue;

                    if (req.diameter_mm.HasValue)
                    {
                        var diam = Core.Units.MetersToFt(req.diameter_mm.Value / 1000.0);
                        conduit.get_Parameter(BuiltInParameter.RBS_CONDUIT_DIAMETER_PARAM)?.Set(diam);
                    }

                    t.Commit();
                }

                return new
                {
                    elementId = id,
                    used = new { level = level.Name, conduitType = ctype.Name, elevation_m = req.elevation_m }
                };
            };
        }

        // =========================
        // 2) CableTray.Create
        // =========================
        private class CableTrayCreateRequest
        {
            public string level { get; set; }
            public string cableTrayType { get; set; }   // opcional
            public double elevation_m { get; set; } = 2.7;
            public Pt2 start { get; set; }
            public Pt2 end { get; set; }
        }

        public static Func<UIApplication, object> CableTrayCreate(JObject args)
        {
            var req = args.ToObject<CableTrayCreateRequest>() ?? throw new Exception("Invalid args for mep.cabletray.create.");
            return (UIApplication app) =>
            {
                var uidoc = app.ActiveUIDocument ?? throw new Exception("No active document.");
                var doc = uidoc.Document;

                var ctypes = new FilteredElementCollector(doc)
                    .OfClass(typeof(CableTrayType)).Cast<CableTrayType>().ToList();

                CableTrayType ctype = null;
                if (string.IsNullOrWhiteSpace(req.cableTrayType))
                    ctype = ctypes.FirstOrDefault();
                else
                    ctype = ctypes.FirstOrDefault(t => t.Name.Equals(req.cableTrayType, StringComparison.OrdinalIgnoreCase));

                if (ctype == null) throw new Exception("No CableTrayType found.");

                var level = ViewHelpers.ResolveLevel(doc, req.level, uidoc.ActiveView);

                double ToFt(double m) => Core.Units.MetersToFt(m);
                var z = level.Elevation + ToFt(req.elevation_m);
                var p1 = new XYZ(ToFt(req.start.x), ToFt(req.start.y), z);
                var p2 = new XYZ(ToFt(req.end.x), ToFt(req.end.y), z);
                if (p1.IsAlmostEqualTo(p2)) throw new Exception("Start and end points are the same.");

                int id;
                using (var t = new Transaction(doc, "MCP: MEP.CableTray.Create"))
                {
                    t.Start();
                    var tray = CableTray.Create(doc, ctype.Id, p1, p2, level.Id);
                    id = tray.Id.IntegerValue;
                    t.Commit();
                }

                return new { elementId = id, used = new { cableTrayType = ctype.Name, level = level.Name, elevation_m = req.elevation_m } };
            };
        }

        // =========================
        // 3) Pipe.Connect (elbow/transition)
        // =========================
        private class PipeConnectReq
        {
            public int aId { get; set; }
            public int bId { get; set; }
            public string mode { get; set; } // "auto" (default), "elbow", "transition"
        }

        public static Func<UIApplication, object> PipeConnect(JObject args)
        {
            var req = args.ToObject<PipeConnectReq>() ?? throw new Exception("Invalid args for mep.pipe.connect.");
            return (UIApplication app) =>
            {
                var doc = app.ActiveUIDocument?.Document ?? throw new Exception("No active document.");
                var a = Require(doc, req.aId, "Element");
                var b = Require(doc, req.bId, "Element");

                var conA = CollectConnectors(a).Where(IsFreeEnd).ToList();
                var conB = CollectConnectors(b).Where(IsFreeEnd).ToList();
                if (conA.Count == 0 || conB.Count == 0) throw new Exception("No free connectors to connect.");

                var (c1, c2) = FindClosestPair(conA, conB);
                if (c1 == null || c2 == null) throw new Exception("Could not resolve connector pair.");

                var mode = (req.mode ?? "auto").ToLowerInvariant();
                bool useTransition = mode == "transition" || (mode == "auto" && !SameSize(c1, c2));

                int fittingId;
                using (var t = new Transaction(doc, "MCP: MEP.Pipe.Connect"))
                {
                    t.Start();
                    FamilyInstance fit = useTransition
                        ? CreatePipeTransitionFittingCompat(doc, c1, c2)
                        : CreatePipeElbowFittingCompat(doc, c1, c2);
                    fittingId = fit.Id.IntegerValue;
                    t.Commit();
                }

                return new { createdFittingId = fittingId, mode = useTransition ? "transition" : "elbow" };
            };
        }

        // =========================
        // 4) Pipe.AddInsulation (bulk)
        // =========================
        private class PipeInsulationReq
        {
            public int[] pipeIds { get; set; } = Array.Empty<int>();
            public string typeName { get; set; }         // opcional
            public double? thickness_mm { get; set; }    // opcional
        }

        public static Func<UIApplication, object> PipeAddInsulation(JObject args)
        {
            var req = args.ToObject<PipeInsulationReq>() ?? new PipeInsulationReq();
            return (UIApplication app) =>
            {
                var doc = app.ActiveUIDocument?.Document ?? throw new Exception("No active document.");

                var types = new FilteredElementCollector(doc).OfClass(typeof(PipeInsulationType)).Cast<PipeInsulationType>().ToList();
                if (types.Count == 0) throw new Exception("No PipeInsulationType in project.");

                PipeInsulationType tSel = null;
                if (!string.IsNullOrWhiteSpace(req.typeName))
                    tSel = types.FirstOrDefault(t => t.Name.Equals(req.typeName, StringComparison.OrdinalIgnoreCase));
                if (tSel == null) tSel = types.First();

                var created = new List<int>();
                var failed = new List<object>();
                double? thicknessFt = req.thickness_mm.HasValue ? Core.Units.MetersToFt(req.thickness_mm.Value / 1000.0) : (double?)null;

                using (var t = new Transaction(doc, "MCP: MEP.Pipe.AddInsulation"))
                {
                    t.Start();
                    foreach (var id in req.pipeIds ?? Array.Empty<int>())
                    {
                        try
                        {
                            if (doc.GetElement(new ElementId(id)) is Pipe p)
                            {
                                Element createdElem;
                                var insId = CreatePipeInsulationCompat(doc, p.Id, tSel.Id, thicknessFt, out createdElem);

                                if (thicknessFt.HasValue)
                                {
                                    TrySetDoubleParamByNames(createdElem,
                                        new[] { "Thickness", "Espesor", "Épaisseur", "Dicke" },
                                        thicknessFt.Value);
                                }
                                created.Add(insId.IntegerValue);
                            }
                        }
                        catch (Exception ex)
                        {
                            failed.Add(new { pipeId = id, error = ex.Message });
                        }
                    }
                    t.Commit();
                }
                return new { type = tSel.Name, createdCount = created.Count, created, failed };
            };
        }

        // =========================
        // 5) Pipe.SetOffset (bulk)
        // =========================
        private class PipeSetOffsetReq
        {
            public int[] pipeIds { get; set; } = Array.Empty<int>();
            public double offset_m { get; set; }
        }

        public static Func<UIApplication, object> PipeSetOffset(JObject args)
        {
            var req = args.ToObject<PipeSetOffsetReq>() ?? throw new Exception("Invalid args for mep.pipe.set_offset.");
            return (UIApplication app) =>
            {
                var doc = app.ActiveUIDocument?.Document ?? throw new Exception("No active document.");
                var updated = new List<int>();
                var failed = new List<object>();
                var val = Core.Units.MetersToFt(req.offset_m);

                using (var t = new Transaction(doc, "MCP: MEP.Pipe.SetOffset"))
                {
                    t.Start();
                    foreach (var id in req.pipeIds ?? Array.Empty<int>())
                    {
                        try
                        {
                            if (doc.GetElement(new ElementId(id)) is Pipe p)
                            {
                                var par = p.get_Parameter(BuiltInParameter.RBS_OFFSET_PARAM);
                                if (par != null && !par.IsReadOnly && par.StorageType == StorageType.Double)
                                { par.Set(val); updated.Add(id); }
                                else failed.Add(new { pipeId = id, reason = "Offset param not settable" });
                            }
                        }
                        catch (Exception ex)
                        { failed.Add(new { pipeId = id, error = ex.Message }); }
                    }
                    t.Commit();
                }
                return new { updatedCount = updated.Count, updated, failed, offset_m = req.offset_m };
            };
        }

        // =========================
        // 6) Conduit.SetOffset (bulk)
        // =========================
        private class ConduitSetOffsetReq
        {
            public int[] conduitIds { get; set; } = Array.Empty<int>();
            public double offset_m { get; set; }
        }

        public static Func<UIApplication, object> ConduitSetOffset(JObject args)
        {
            var req = args.ToObject<ConduitSetOffsetReq>() ?? throw new Exception("Invalid args for mep.conduit.set_offset.");
            return (UIApplication app) =>
            {
                var doc = app.ActiveUIDocument?.Document ?? throw new Exception("No active document.");
                var updated = new List<int>();
                var failed = new List<object>();
                var val = Core.Units.MetersToFt(req.offset_m);

                using (var t = new Transaction(doc, "MCP: MEP.Conduit.SetOffset"))
                {
                    t.Start();
                    foreach (var id in req.conduitIds ?? Array.Empty<int>())
                    {
                        try
                        {
                            if (doc.GetElement(new ElementId(id)) is Conduit c)
                            {
                                var par = c.get_Parameter(BuiltInParameter.RBS_OFFSET_PARAM);
                                if (par != null && !par.IsReadOnly && par.StorageType == StorageType.Double)
                                { par.Set(val); updated.Add(id); }
                                else failed.Add(new { conduitId = id, reason = "Offset param not settable" });
                            }
                        }
                        catch (Exception ex)
                        { failed.Add(new { conduitId = id, error = ex.Message }); }
                    }
                    t.Commit();
                }
                return new { updatedCount = updated.Count, updated, failed, offset_m = req.offset_m };
            };
        }

        // =========================
        // 7) Conduit.SetDiameter (bulk)
        // =========================
        private class ConduitSetDiameterReq
        {
            public int[] conduitIds { get; set; } = Array.Empty<int>();
            public double diameter_mm { get; set; }
        }

        public static Func<UIApplication, object> ConduitSetDiameter(JObject args)
        {
            var req = args.ToObject<ConduitSetDiameterReq>() ?? throw new Exception("Invalid args for mep.conduit.set_diameter.");
            return (UIApplication app) =>
            {
                var doc = app.ActiveUIDocument?.Document ?? throw new Exception("No active document.");
                var updated = new List<int>();
                var failed = new List<object>();
                var val = Core.Units.MetersToFt(req.diameter_mm / 1000.0);

                using (var t = new Transaction(doc, "MCP: MEP.Conduit.SetDiameter"))
                {
                    t.Start();
                    foreach (var id in req.conduitIds ?? Array.Empty<int>())
                    {
                        try
                        {
                            if (doc.GetElement(new ElementId(id)) is Conduit c)
                            {
                                var par = c.get_Parameter(BuiltInParameter.RBS_CONDUIT_DIAMETER_PARAM);
                                if (par != null && !par.IsReadOnly && par.StorageType == StorageType.Double)
                                { par.Set(val); updated.Add(id); }
                                else failed.Add(new { conduitId = id, reason = "Diameter param not settable" });
                            }
                        }
                        catch (Exception ex)
                        { failed.Add(new { conduitId = id, error = ex.Message }); }
                    }
                    t.Commit();
                }
                return new { updatedCount = updated.Count, updated, failed, diameter_mm = req.diameter_mm };
            };
        }

        // =========================
        // 8) CableTray.SetSize (bulk)
        // =========================
        private class CableTraySetSizeReq
        {
            public int[] trayIds { get; set; } = Array.Empty<int>();
            public double? width_mm { get; set; }   // opcional
            public double? height_mm { get; set; }  // opcional
        }

        public static Func<UIApplication, object> CableTraySetSize(JObject args)
        {
            var req = args.ToObject<CableTraySetSizeReq>() ?? throw new Exception("Invalid args for mep.cabletray.set_size.");
            return (UIApplication app) =>
            {
                var doc = app.ActiveUIDocument?.Document ?? throw new Exception("No active document.");
                var updated = new List<int>();
                var failed = new List<object>();

                double? w = req.width_mm.HasValue ? Core.Units.MetersToFt(req.width_mm.Value / 1000.0) : (double?)null;
                double? h = req.height_mm.HasValue ? Core.Units.MetersToFt(req.height_mm.Value / 1000.0) : (double?)null;

                using (var t = new Transaction(doc, "MCP: MEP.CableTray.SetSize"))
                {
                    t.Start();
                    foreach (var id in req.trayIds ?? Array.Empty<int>())
                    {
                        try
                        {
                            if (doc.GetElement(new ElementId(id)) is CableTray tray)
                            {
                                bool okAny = false;

                                if (w.HasValue)
                                {
                                    var pW = tray.get_Parameter(BuiltInParameter.RBS_CABLETRAY_WIDTH_PARAM);
                                    if (pW != null && !pW.IsReadOnly && pW.StorageType == StorageType.Double)
                                    { pW.Set(w.Value); okAny = true; }
                                }
                                if (h.HasValue)
                                {
                                    var pH = tray.get_Parameter(BuiltInParameter.RBS_CABLETRAY_HEIGHT_PARAM);
                                    if (pH != null && !pH.IsReadOnly && pH.StorageType == StorageType.Double)
                                    { pH.Set(h.Value); okAny = true; }
                                }

                                if (okAny) updated.Add(id);
                                else failed.Add(new { trayId = id, reason = "Width/Height not settable" });
                            }
                        }
                        catch (Exception ex)
                        { failed.Add(new { trayId = id, error = ex.Message }); }
                    }
                    t.Commit();
                }
                return new { updatedCount = updated.Count, updated, failed, width_mm = req.width_mm, height_mm = req.height_mm };
            };
        }

        // =========================
        // 9) CableTray.SetOffset (bulk)
        // =========================
        private class CableTraySetOffsetReq
        {
            public int[] trayIds { get; set; } = Array.Empty<int>();
            public double offset_m { get; set; }
        }

        public static Func<UIApplication, object> CableTraySetOffset(JObject args)
        {
            var req = args.ToObject<CableTraySetOffsetReq>() ?? throw new Exception("Invalid args for mep.cabletray.set_offset.");
            return (UIApplication app) =>
            {
                var doc = app.ActiveUIDocument?.Document ?? throw new Exception("No active document.");
                var updated = new List<int>();
                var failed = new List<object>();
                var val = Core.Units.MetersToFt(req.offset_m);

                using (var t = new Transaction(doc, "MCP: MEP.CableTray.SetOffset"))
                {
                    t.Start();
                    foreach (var id in req.trayIds ?? Array.Empty<int>())
                    {
                        try
                        {
                            if (doc.GetElement(new ElementId(id)) is CableTray tray)
                            {
                                var par = tray.get_Parameter(BuiltInParameter.RBS_OFFSET_PARAM);
                                if (par != null && !par.IsReadOnly && par.StorageType == StorageType.Double)
                                { par.Set(val); updated.Add(id); }
                                else failed.Add(new { trayId = id, reason = "Offset param not settable" });
                            }
                        }
                        catch (Exception ex)
                        { failed.Add(new { trayId = id, error = ex.Message }); }
                    }
                    t.Commit();
                }
                return new { updatedCount = updated.Count, updated, failed, offset_m = req.offset_m };
            };
        }
    }
}
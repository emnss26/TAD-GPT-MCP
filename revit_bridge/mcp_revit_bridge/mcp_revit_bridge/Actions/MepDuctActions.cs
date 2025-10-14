using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
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
    internal class MepDuctActions
    {
        private class Pt2 { public double x; public double y; }

        // =========================
        // Helpers generales
        // =========================
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
                {
                    return Nearly(c1.Radius, c2.Radius);
                }
                else
                {
                    bool w = c1.Width > 0 && c2.Width > 0 ? Nearly(c1.Width, c2.Width) : true;
                    bool h = c1.Height > 0 && c2.Height > 0 ? Nearly(c1.Height, c2.Height) : true;
                    return w && h;
                }
            }
            catch { return false; }
        }

        private static bool IsFreeEnd(Connector c)
        {
            try { return c != null && !c.IsConnected; } catch { return true; }
        }

        private static Element Require(Document doc, int id, string label)
        {
            var e = doc.GetElement(new ElementId(id));
            if (e == null) throw new Exception($"{label} {id} not found.");
            return e;
        }

        private static double ToFt_m(double m) => Core.Units.MetersToFt(m);

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
                        {
                            p.Set(val);
                            return true;
                        }
                    }
                }
            }
            return false;
        }

        // =========================
        // Helpers compat (reflexión)
        // =========================

        // Crea codo vía reflexión (distintas builds)
        private static FamilyInstance CreateElbowFittingCompat(Document doc, Connector c1, Connector c2)
        {
            var asm = typeof(Duct).Assembly;
            // Tipo esperado: Autodesk.Revit.DB.Mechanical.MechanicalUtils
            var t = asm.GetType("Autodesk.Revit.DB.Mechanical.MechanicalUtils");
            if (t != null)
            {
                // Firma 1: (Document, Connector, Connector)
                var m1 = t.GetMethod("CreateElbowFitting",
                    BindingFlags.Public | BindingFlags.Static,
                    null,
                    new Type[] { typeof(Document), typeof(Connector), typeof(Connector) },
                    null);
                if (m1 != null)
                {
                    var o = m1.Invoke(null, new object[] { doc, c1, c2 }) as FamilyInstance;
                    if (o != null) return o;
                }

                // Firma 2: (Connector, Connector)
                var m2 = t.GetMethod("CreateElbowFitting",
                    BindingFlags.Public | BindingFlags.Static,
                    null,
                    new Type[] { typeof(Connector), typeof(Connector) },
                    null);
                if (m2 != null)
                {
                    var o = m2.Invoke(null, new object[] { c1, c2 }) as FamilyInstance;
                    if (o != null) return o;
                }
            }
            throw new Exception("CreateElbowFitting no disponible en esta versión del API.");
        }

        // Crea transición vía reflexión (distintas builds)
        private static FamilyInstance CreateTransitionFittingCompat(Document doc, Connector c1, Connector c2)
        {
            var asm = typeof(Duct).Assembly;
            var t = asm.GetType("Autodesk.Revit.DB.Mechanical.MechanicalUtils");
            if (t != null)
            {
                // Firma 1: (Document, Connector, Connector)
                var m1 = t.GetMethod("CreateTransitionFitting",
                    BindingFlags.Public | BindingFlags.Static,
                    null,
                    new Type[] { typeof(Document), typeof(Connector), typeof(Connector) },
                    null);
                if (m1 != null)
                {
                    var o = m1.Invoke(null, new object[] { doc, c1, c2 }) as FamilyInstance;
                    if (o != null) return o;
                }

                // Firma 2: (Connector, Connector)
                var m2 = t.GetMethod("CreateTransitionFitting",
                    BindingFlags.Public | BindingFlags.Static,
                    null,
                    new Type[] { typeof(Connector), typeof(Connector) },
                    null);
                if (m2 != null)
                {
                    var o = m2.Invoke(null, new object[] { c1, c2 }) as FamilyInstance;
                    if (o != null) return o;
                }
            }
            throw new Exception("CreateTransitionFitting no disponible en esta versión del API.");
        }

        // Crea Insulation con compat de firmas
        private static ElementId CreateDuctInsulationCompat(Document doc, ElementId curveId, ElementId typeId, double? thicknessFt, out Element created)
        {
            created = null;
            var t = typeof(DuctInsulation);
            // Firma nueva: Create(Document, ElementId, ElementId, double thickness)
            var mNew = t.GetMethod("Create",
                BindingFlags.Public | BindingFlags.Static,
                null,
                new Type[] { typeof(Document), typeof(ElementId), typeof(ElementId), typeof(double) },
                null);

            if (mNew != null)
            {
                double th = thicknessFt ?? 0.0;
                created = mNew.Invoke(null, new object[] { doc, curveId, typeId, th }) as Element;
                if (created == null) throw new Exception("DuctInsulation.Create devolvió null.");
                return created.Id;
            }

            // Firma antigua: Create(Document, ElementId, ElementId)
            var mOld = t.GetMethod("Create",
                BindingFlags.Public | BindingFlags.Static,
                null,
                new Type[] { typeof(Document), typeof(ElementId), typeof(ElementId) },
                null);

            if (mOld != null)
            {
                created = mOld.Invoke(null, new object[] { doc, curveId, typeId }) as Element;
                if (created == null) throw new Exception("DuctInsulation.Create devolvió null.");
                return created.Id;
            }

            throw new Exception("No se encontró DuctInsulation.Create (firmas soportadas).");
        }

        // Crea Lining con compat de firmas
        private static ElementId CreateDuctLiningCompat(Document doc, ElementId curveId, ElementId typeId, double? thicknessFt, out Element created)
        {
            created = null;
            var t = typeof(DuctLining);
            // Firma nueva: Create(Document, ElementId, ElementId, double thickness)
            var mNew = t.GetMethod("Create",
                BindingFlags.Public | BindingFlags.Static,
                null,
                new Type[] { typeof(Document), typeof(ElementId), typeof(ElementId), typeof(double) },
                null);

            if (mNew != null)
            {
                double th = thicknessFt ?? 0.0;
                created = mNew.Invoke(null, new object[] { doc, curveId, typeId, th }) as Element;
                if (created == null) throw new Exception("DuctLining.Create devolvió null.");
                return created.Id;
            }

            // Firma antigua: Create(Document, ElementId, ElementId)
            var mOld = t.GetMethod("Create",
                BindingFlags.Public | BindingFlags.Static,
                null,
                new Type[] { typeof(Document), typeof(ElementId), typeof(ElementId) },
                null);

            if (mOld != null)
            {
                created = mOld.Invoke(null, new object[] { doc, curveId, typeId }) as Element;
                if (created == null) throw new Exception("DuctLining.Create devolvió null.");
                return created.Id;
            }

            throw new Exception("No se encontró DuctLining.Create (firmas soportadas).");
        }

        // =========================
        // 0) Crear ducto
        // =========================
        private class DuctCreateRequest
        {
            public string level { get; set; }               // opcional
            public string systemType { get; set; }          // nombre o clasificación (SupplyAir, ReturnAir, etc) opcional
            public string ductType { get; set; }            // opcional
            public double elevation_m { get; set; } = 2.7;  // altura sobre nivel
            public Pt2 start { get; set; }
            public Pt2 end { get; set; }
            public double? width_mm { get; set; }           // opcional (rectangular)
            public double? height_mm { get; set; }          // opcional (rectangular)
            public double? diameter_mm { get; set; }        // opcional (redondo)
        }

        public static Func<UIApplication, object> DuctCreate(JObject args)
        {
            var req = args.ToObject<DuctCreateRequest>() ?? throw new Exception("Invalid args for mep.duct.create.");
            return (UIApplication app) =>
            {
                var uidoc = app.ActiveUIDocument ?? throw new Exception("No active document.");
                var doc = uidoc.Document;
                var active = uidoc.ActiveView;

                var level = ViewHelpers.ResolveLevel(doc, req.level, active);
                double z = level.Elevation + Core.Units.MetersToFt(req.elevation_m);

                var systems = new FilteredElementCollector(doc)
                    .OfClass(typeof(MEPSystemType)).Cast<MEPSystemType>().ToList();

                MEPSystemType sys = null;
                if (!string.IsNullOrWhiteSpace(req.systemType))
                {
                    sys = systems.FirstOrDefault(s =>
                        s.Name.Equals(req.systemType, StringComparison.OrdinalIgnoreCase) ||
                        s.SystemClassification.ToString().Equals(req.systemType, StringComparison.OrdinalIgnoreCase));
                }
                if (sys == null) sys = systems.FirstOrDefault();
                if (sys == null) throw new Exception("No MEPSystemType available for ducts.");

                if (req.start == null || req.end == null)
                    throw new Exception("Duct requires start and end points.");

                var dtypes = new FilteredElementCollector(doc).OfClass(typeof(DuctType)).Cast<DuctType>().ToList();
                DuctType dtype = null;
                if (!string.IsNullOrWhiteSpace(req.ductType))
                {
                    dtype = dtypes.FirstOrDefault(t =>
                        t.Name.Equals(req.ductType, StringComparison.OrdinalIgnoreCase) ||
                        $"{t.FamilyName}: {t.Name}".Equals(req.ductType, StringComparison.OrdinalIgnoreCase));
                }
                if (dtype == null) dtype = dtypes.FirstOrDefault();

                if (dtype == null)
                {
                    var sample = string.Join(", ", dtypes.Take(10).Select(t => $"{t.FamilyName}: {t.Name}"));
                    throw new Exception("No suitable DuctType found." + (sample.Length > 0 ? $" Available examples: {sample}" : ""));
                }

                double ToFt(double m) => Core.Units.MetersToFt(m);
                var p1 = new XYZ(ToFt(req.start.x), ToFt(req.start.y), z);
                var p2 = new XYZ(ToFt(req.end.x), ToFt(req.end.y), z);
                if (p1.IsAlmostEqualTo(p2)) throw new Exception("Start and end points are the same.");

                int id;
                using (var t = new Transaction(doc, "MCP: MEP.Duct.Create"))
                {
                    t.Start();
                    var duct = Duct.Create(doc, sys.Id, dtype.Id, level.Id, p1, p2);
                    id = duct.Id.IntegerValue;

                    if (req.diameter_mm.HasValue)
                    {
                        var diam = Core.Units.MetersToFt(req.diameter_mm.Value / 1000.0);
                        duct.get_Parameter(BuiltInParameter.RBS_CURVE_DIAMETER_PARAM)?.Set(diam);
                    }
                    else
                    {
                        if (req.width_mm.HasValue)
                        {
                            var w = Core.Units.MetersToFt(req.width_mm.Value / 1000.0);
                            duct.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM)?.Set(w);
                        }
                        if (req.height_mm.HasValue)
                        {
                            var h = Core.Units.MetersToFt(req.height_mm.Value / 1000.0);
                            duct.get_Parameter(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM)?.Set(h);
                        }
                    }

                    t.Commit();
                }

                return new
                {
                    elementId = id,
                    used = new
                    {
                        level = level.Name,
                        systemType = $"{sys.SystemClassification} / {sys.Name}",
                        ductType = $"{dtype.FamilyName}: {dtype.Name}",
                        elevation_m = req.elevation_m
                    }
                };
            };
        }

        // =========================
        // 1) Conectar (elbow/transition)
        // =========================
        private class DuctConnectReq
        {
            public int aId { get; set; }
            public int bId { get; set; }
            public string mode { get; set; } // "auto" (default), "elbow", "transition"
        }

        public static Func<UIApplication, object> DuctConnect(JObject args)
        {
            var req = args.ToObject<DuctConnectReq>() ?? throw new Exception("Invalid args for mep.duct.connect.");
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
                using (var t = new Transaction(doc, "MCP: MEP.Duct.Connect"))
                {
                    t.Start();
                    FamilyInstance fit = useTransition
                        ? CreateTransitionFittingCompat(doc, c1, c2)
                        : CreateElbowFittingCompat(doc, c1, c2);
                    fittingId = fit.Id.IntegerValue;
                    t.Commit();
                }

                return new { createdFittingId = fittingId, mode = useTransition ? "transition" : "elbow" };
            };
        }

        // =========================
        // 2) Insulation (bulk)
        // =========================
        private class DuctInsulationReq
        {
            public int[] ductIds { get; set; } = Array.Empty<int>();
            public string typeName { get; set; }
            public double? thickness_mm { get; set; }
        }

        public static Func<UIApplication, object> DuctAddInsulation(JObject args)
        {
            var req = args.ToObject<DuctInsulationReq>() ?? new DuctInsulationReq();
            return (UIApplication app) =>
            {
                var doc = app.ActiveUIDocument?.Document ?? throw new Exception("No active document.");
                var types = new FilteredElementCollector(doc).OfClass(typeof(DuctInsulationType)).Cast<DuctInsulationType>().ToList();
                if (types.Count == 0) throw new Exception("No DuctInsulationType in project.");

                DuctInsulationType tSel = null;
                if (!string.IsNullOrWhiteSpace(req.typeName))
                    tSel = types.FirstOrDefault(t => t.Name.Equals(req.typeName, StringComparison.OrdinalIgnoreCase));
                if (tSel == null) tSel = types.First();

                var created = new List<int>();
                var failed = new List<object>();
                double? thicknessFt = req.thickness_mm.HasValue ? Core.Units.MetersToFt(req.thickness_mm.Value / 1000.0) : (double?)null;

                using (var t = new Transaction(doc, "MCP: MEP.Duct.AddInsulation"))
                {
                    t.Start();
                    foreach (var id in req.ductIds ?? Array.Empty<int>())
                    {
                        try
                        {
                            if (doc.GetElement(new ElementId(id)) is Duct d)
                            {
                                Element createdElem;
                                var insId = CreateDuctInsulationCompat(doc, d.Id, tSel.Id, thicknessFt, out createdElem);

                                // Si fue la firma "antigua", setear espesor por parámetro si se pidió
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
                            failed.Add(new { ductId = id, error = ex.Message });
                        }
                    }
                    t.Commit();
                }
                return new { type = tSel.Name, createdCount = created.Count, created, failed };
            };
        }

        // =========================
        // 3) Lining (bulk)
        // =========================
        private class DuctLiningReq
        {
            public int[] ductIds { get; set; } = Array.Empty<int>();
            public string typeName { get; set; }
            public double? thickness_mm { get; set; }
        }

        public static Func<UIApplication, object> DuctAddLining(JObject args)
        {
            var req = args.ToObject<DuctLiningReq>() ?? new DuctLiningReq();
            return (UIApplication app) =>
            {
                var doc = app.ActiveUIDocument?.Document ?? throw new Exception("No active document.");
                var types = new FilteredElementCollector(doc).OfClass(typeof(DuctLiningType)).Cast<DuctLiningType>().ToList();
                if (types.Count == 0) throw new Exception("No DuctLiningType in project.");

                DuctLiningType tSel = null;
                if (!string.IsNullOrWhiteSpace(req.typeName))
                    tSel = types.FirstOrDefault(t => t.Name.Equals(req.typeName, StringComparison.OrdinalIgnoreCase));
                if (tSel == null) tSel = types.First();

                var created = new List<int>();
                var failed = new List<object>();
                double? thicknessFt = req.thickness_mm.HasValue ? Core.Units.MetersToFt(req.thickness_mm.Value / 1000.0) : (double?)null;

                using (var t = new Transaction(doc, "MCP: MEP.Duct.AddLining"))
                {
                    t.Start();
                    foreach (var id in req.ductIds ?? Array.Empty<int>())
                    {
                        try
                        {
                            if (doc.GetElement(new ElementId(id)) is Duct d)
                            {
                                Element createdElem;
                                var linId = CreateDuctLiningCompat(doc, d.Id, tSel.Id, thicknessFt, out createdElem);

                                if (thicknessFt.HasValue)
                                {
                                    TrySetDoubleParamByNames(createdElem,
                                        new[] { "Thickness", "Espesor", "Épaisseur", "Dicke" },
                                        thicknessFt.Value);
                                }

                                created.Add(linId.IntegerValue);
                            }
                        }
                        catch (Exception ex)
                        {
                            failed.Add(new { ductId = id, error = ex.Message });
                        }
                    }
                    t.Commit();
                }
                return new { type = tSel.Name, createdCount = created.Count, created, failed };
            };
        }

        // =========================
        // 4) Set Offset (bulk)
        // =========================
        private class DuctSetOffsetReq
        {
            public int[] ductIds { get; set; } = Array.Empty<int>();
            public double offset_m { get; set; }     // offset desde el nivel del ducto
        }

        public static Func<UIApplication, object> DuctSetOffset(JObject args)
        {
            var req = args.ToObject<DuctSetOffsetReq>() ?? throw new Exception("Invalid args for mep.duct.set_offset.");
            return (UIApplication app) =>
            {
                var doc = app.ActiveUIDocument?.Document ?? throw new Exception("No active document.");
                var updated = new List<int>();
                var failed = new List<object>();
                var val = Core.Units.MetersToFt(req.offset_m);

                using (var t = new Transaction(doc, "MCP: MEP.Duct.SetOffset"))
                {
                    t.Start();
                    foreach (var id in req.ductIds ?? Array.Empty<int>())
                    {
                        try
                        {
                            if (doc.GetElement(new ElementId(id)) is Duct d)
                            {
                                var p = d.get_Parameter(BuiltInParameter.RBS_OFFSET_PARAM);
                                if (p != null && !p.IsReadOnly && p.StorageType == StorageType.Double)
                                {
                                    p.Set(val);
                                    updated.Add(id);
                                }
                                else
                                {
                                    failed.Add(new { ductId = id, reason = "Offset param not settable" });
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            failed.Add(new { ductId = id, error = ex.Message });
                        }
                    }
                    t.Commit();
                }
                return new { updatedCount = updated.Count, updated, failed, offset_m = req.offset_m };
            };
        }
    }
}
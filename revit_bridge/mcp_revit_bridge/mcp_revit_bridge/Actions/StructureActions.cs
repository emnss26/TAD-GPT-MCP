using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using mcp_app.Core;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection.Emit;
using System.Security.Cryptography;
using System.Xml.Linq;
using static System.Net.Mime.MediaTypeNames;

namespace mcp_app.Actions
{
    internal class StructureActions
    {
        private class Pt2 { public double x; public double y; }

        // =========================
        // Helpers
        // =========================
        static double ToFt(double m) => Core.Units.MetersToFt(m);

        static Level ResolveLevel(Document doc, string levelName, View active)
            => ViewHelpers.ResolveLevel(doc, levelName, active);

        static FamilySymbol FindSymbolByName(Document doc, BuiltInCategory bic, string familyTypeOrNull)
        {
            var symbols = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>()
                .Where(fs => fs.Category != null && fs.Category.Id.IntegerValue == (int)bic)
                .ToList();

            if (symbols.Count == 0) return null;

            if (string.IsNullOrWhiteSpace(familyTypeOrNull))
                return symbols.First();

            return symbols.FirstOrDefault(fs =>
                fs.Name.Equals(familyTypeOrNull, StringComparison.OrdinalIgnoreCase) ||
                $"{fs.FamilyName}: {fs.Name}".Equals(familyTypeOrNull, StringComparison.OrdinalIgnoreCase));
        }

        static IList<Curve> BuildClosedProfileXY(Pt2[] poly, double z)
        {
            if (poly == null || poly.Length < 3)
                throw new Exception("Profile requires at least 3 points.");

            var curves = new List<Curve>(poly.Length);
            for (int i = 0; i < poly.Length; i++)
            {
                var a = poly[i];
                var b = poly[(i + 1) % poly.Length];
                var p1 = new XYZ(ToFt(a.x), ToFt(a.y), z);
                var p2 = new XYZ(ToFt(b.x), ToFt(b.y), z);
                if (p1.IsAlmostEqualTo(p2)) continue;
                curves.Add(Line.CreateBound(p1, p2));
            }
            if (curves.Count < 3) throw new Exception("Invalid profile.");
            return curves;
        }

        static bool TrySetOffsetByName(Element e, double offsetFt)
        {
            if (e == null) return false;

            // Prioridad: nombres más comunes en fundaciones de muro (varía por idioma/template)
            var preferred = new[]
            {
        "Elevation at Top",     // EN
        "Top Elevation",
        "Level Offset",
        "Top Offset",
        "Offset",
        "Elevation",
        "Cota superior",        // ES posibles
        "Desfase de nivel",
        "Desfase",
        "Elevación"
    };

            // 1) Coincidencia por nombre de parámetro
            foreach (Parameter p in e.Parameters)
            {
                if (p == null || p.Definition == null) continue;
                if (p.StorageType != StorageType.Double || p.IsReadOnly) continue;

                var name = p.Definition.Name ?? "";
                foreach (var key in preferred)
                {
                    if (name.IndexOf(key, StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        try { p.Set(offsetFt); return true; } catch { /* intenta siguiente */ }
                    }
                }
            }

            // 2) Fallback: intenta algunos nombres genéricos cortos
            foreach (Parameter p in e.Parameters)
            {
                if (p == null || p.Definition == null) continue;
                if (p.StorageType != StorageType.Double || p.IsReadOnly) continue;

                var name = p.Definition.Name ?? "";
                if (name.Equals("Offset", StringComparison.OrdinalIgnoreCase))
                {
                    try { p.Set(offsetFt); return true; } catch { }
                }
            }

            return false;
        }


        // =========================
        // struct.beam.create
        // =========================
        private class BeamCreateRequest
        {
            public string level { get; set; }         // opcional
            public string familyType { get; set; }    // e.g. "W-Wide Flange : W12x26"
            public double elevation_m { get; set; } = 3.0;
            public Pt2 start { get; set; }
            public Pt2 end { get; set; }
        }

        public static Func<UIApplication, object> BeamCreate(JObject args)
        {
            var req = args.ToObject<BeamCreateRequest>() ?? throw new Exception("Invalid args for struct.beam.create.");
            return (UIApplication app) =>
            {
                var uidoc = app.ActiveUIDocument ?? throw new Exception("No active document.");
                var doc = uidoc.Document;
                var active = uidoc.ActiveView;

                var level = ResolveLevel(doc, req.level, active);
                double z = level.Elevation + ToFt(req.elevation_m);

                var sym = FindSymbolByName(doc, BuiltInCategory.OST_StructuralFraming, req.familyType)
                          ?? throw new Exception("No Structural Framing FamilySymbol found.");

                var p1 = new XYZ(ToFt(req.start.x), ToFt(req.start.y), z);
                var p2 = new XYZ(ToFt(req.end.x), ToFt(req.end.y), z);
                var line = Line.CreateBound(p1, p2);

                int id;
                using (var t = new Transaction(doc, "MCP: Struct.Beam.Create"))
                {
                    t.Start();
                    if (!sym.IsActive) sym.Activate();
                    var fi = doc.Create.NewFamilyInstance(line, sym, level, StructuralType.Beam);
                    id = fi.Id.IntegerValue;
                    t.Commit();
                }
                return new { elementId = id, used = new { level = level.Name, familyType = $"{sym.FamilyName}: {sym.Name}" } };
            };
        }

        // =========================
        // struct.column.create
        // =========================
        private class ColumnCreateRequest
        {
            public string level { get; set; }
            public string familyType { get; set; } // e.g. "Concrete-Rectangular Column : 300 x 600"
            public double elevation_m { get; set; } = 0.0; // base offset
            public Pt2 point { get; set; }                 // XY en m
        }

        public static Func<UIApplication, object> ColumnCreate(JObject args)
        {
            var req = args.ToObject<ColumnCreateRequest>() ?? throw new Exception("Invalid args for struct.column.create.");
            return (UIApplication app) =>
            {
                var uidoc = app.ActiveUIDocument ?? throw new Exception("No active document.");
                var doc = uidoc.Document;
                var active = uidoc.ActiveView;

                var level = ResolveLevel(doc, req.level, active);
                double z = level.Elevation + ToFt(req.elevation_m);

                var sym = FindSymbolByName(doc, BuiltInCategory.OST_StructuralColumns, req.familyType)
                          ?? throw new Exception("No Structural Column FamilySymbol found.");

                var p = new XYZ(ToFt(req.point.x), ToFt(req.point.y), z);

                int id;
                using (var t = new Transaction(doc, "MCP: Struct.Column.Create"))
                {
                    t.Start();
                    if (!sym.IsActive) sym.Activate();
                    var fi = doc.Create.NewFamilyInstance(p, sym, level, StructuralType.Column);
                    id = fi.Id.IntegerValue;
                    t.Commit();
                }
                return new { elementId = id, used = new { level = level.Name, familyType = $"{sym.FamilyName}: {sym.Name}" } };
            };
        }

        // =========================
        // struct.floor.create (estructural)
        // =========================
        private class SFloorCreateRequest
        {
            public string level { get; set; }
            public string floorType { get; set; } // opcional
            public Pt2[] profile { get; set; }    // polígono cerrado en XY (m)
        }

        public static Func<UIApplication, object> StructuralFloorCreate(JObject args)
        {
            var req = args.ToObject<SFloorCreateRequest>() ?? throw new Exception("Invalid args for struct.floor.create.");
            return (UIApplication app) =>
            {
                var uidoc = app.ActiveUIDocument ?? throw new Exception("No active document.");
                var doc = uidoc.Document;
                var level = ResolveLevel(doc, req.level, uidoc.ActiveView);

                var ftypes = new FilteredElementCollector(doc).OfClass(typeof(FloorType)).Cast<FloorType>().ToList();
                var ftype = string.IsNullOrWhiteSpace(req.floorType)
                    ? ftypes.FirstOrDefault()
                    : ftypes.FirstOrDefault(t =>
                        t.Name.Equals(req.floorType, StringComparison.OrdinalIgnoreCase) ||
                        $"{t.FamilyName}: {t.Name}".Equals(req.floorType, StringComparison.OrdinalIgnoreCase));
                if (ftype == null) throw new Exception("FloorType not found.");

                var curves = BuildClosedProfileXY(req.profile, level.Elevation);

                int id;
                using (var t = new Transaction(doc, "MCP: Struct.Floor.Create"))
                {
                    t.Start();
                    var loop = CurveLoop.Create(curves);
                    var fl = Floor.Create(doc, new List<CurveLoop> { loop }, ftype.Id, level.Id);
                    id = fl.Id.IntegerValue;

                    // Marcar como estructural si existe
                    var p = fl.get_Parameter(BuiltInParameter.FLOOR_PARAM_IS_STRUCTURAL);
                    if (p != null && !p.IsReadOnly) p.Set(1);

                    t.Commit();
                }
                return new { elementId = id, used = new { level = level.Name, floorType = $"{ftype.FamilyName}: {ftype.Name}" } };
            };
        }

        // =========================
        // struct.columns.place_on_grid  (ya tenías)
        // =========================
        private class ColumnsPlaceOnGridRequest
        {
            public string baseLevel { get; set; }
            public string topLevel { get; set; }
            public string familyType { get; set; }
            public string[] gridX { get; set; }
            public string[] gridY { get; set; }
            public string[] gridNames { get; set; }
            public double? baseOffset_m { get; set; }
            public double? topOffset_m { get; set; }
            public bool? onlyIntersectionsInsideActiveCrop { get; set; }
            public double? tolerance_m { get; set; }
            public bool? skipIfColumnExistsNearby { get; set; }
            public string worksetName { get; set; }
            public bool? pinned { get; set; }
            public string orientationRelativeTo { get; set; } // "X"|"Y"|"None"
        }

        public static Func<UIApplication, object> ColumnsPlaceOnGrid(JObject args)
        {
            var req = args.ToObject<ColumnsPlaceOnGridRequest>() ?? throw new Exception("Invalid args for struct.columns.place_on_grid.");

            return (UIApplication app) =>
            {
                // (… tu implementación existente sin cambios …)
                throw new NotImplementedException("ColumnsPlaceOnGrid is already defined in your previous file; deja esta signatura como estaba.");
            };
        }

        // =========================
        // struct.foundation.isolated.create
        // =========================
        private class IsolatedFoundationReq
        {
            public string level { get; set; }           // requerido
            public string familyType { get; set; }      // "Family: Type" o sólo Type (opcional)
            public Pt2 point { get; set; }              // XY en m
            public double? baseOffset_m { get; set; }   // opcional
            public bool? pinned { get; set; }           // opcional
        }

        public static Func<UIApplication, object> FoundationIsolatedCreate(JObject args)
        {
            var req = args.ToObject<IsolatedFoundationReq>() ?? throw new Exception("Invalid args for struct.foundation.isolated.create.");
            return (UIApplication app) =>
            {
                var uidoc = app.ActiveUIDocument ?? throw new Exception("No active document.");
                var doc = uidoc.Document;
                var level = ResolveLevel(doc, req.level, uidoc.ActiveView);

                var sym = FindSymbolByName(doc, BuiltInCategory.OST_StructuralFoundation, req.familyType)
                          ?? throw new Exception("No Structural Foundation FamilySymbol (isolated) found.");

                var p = new XYZ(ToFt(req.point.x), ToFt(req.point.y), level.Elevation);

                int id;
                using (var t = new Transaction(doc, "MCP: Struct.Foundation.Isolated.Create"))
                {
                    t.Start();
                    if (!sym.IsActive) sym.Activate();
                    var fi = doc.Create.NewFamilyInstance(p, sym, level, StructuralType.Footing);

                    // base offset si existe el parámetro
                    if (req.baseOffset_m.HasValue)
                    {
                        var par = fi.get_Parameter(BuiltInParameter.SCHEDULE_BASE_LEVEL_OFFSET_PARAM)
                                  ?? fi.get_Parameter(BuiltInParameter.FAMILY_BASE_LEVEL_OFFSET_PARAM);
                        if (par != null && !par.IsReadOnly) par.Set(ToFt(req.baseOffset_m.Value));
                    }

                    if (req.pinned == true) fi.Pinned = true;

                    id = fi.Id.IntegerValue;
                    t.Commit();
                }
                return new { elementId = id, used = new { level = level.Name, familyType = $"{sym.FamilyName}: {sym.Name}" } };
            };
        }

        // =========================
        // struct.foundation.wall.create  (continuous footing bajo muro)
        // =========================
        private class WallFoundationReq
        {
            public int wallId { get; set; }                 // requerido
            public string typeName { get; set; }            // opcional (WallFoundationType name)
            public double? baseOffset_m { get; set; }       // opcional
        }

        public static Func<UIApplication, object> WallFoundationCreate(JObject args)
        {
            var req = args.ToObject<WallFoundationReq>() ?? throw new Exception("Invalid args for struct.foundation.wall.create.");

            return (UIApplication app) =>
            {
                var uidoc = app.ActiveUIDocument ?? throw new Exception("No active document.");
                var doc = uidoc.Document;

                var wall = doc.GetElement(new ElementId(req.wallId)) as Wall
                           ?? throw new Exception($"Wall {req.wallId} not found.");

                // Tipos de foundation continua
                ElementType wfType = null;
                var wallFTypes = new FilteredElementCollector(doc)
                    .OfClass(typeof(WallFoundationType)) // DB.Structure
                    .Cast<ElementType>()
                    .ToList();

                if (wallFTypes.Count == 0)
                    throw new Exception("No WallFoundationType types in model.");

                if (string.IsNullOrWhiteSpace(req.typeName))
                    wfType = wallFTypes.First();
                else
                    wfType = wallFTypes.FirstOrDefault(t => t.Name.Equals(req.typeName, StringComparison.OrdinalIgnoreCase))
                             ?? wallFTypes.First();

                int id;
                using (var t = new Transaction(doc, "MCP: Struct.Foundation.Wall.Create"))
                {
                    t.Start();

                    // API moderna
                    WallFoundation wf = null;
                    var create = typeof(WallFoundation).GetMethod(
                        "Create",
                        new Type[] { typeof(Document), typeof(ElementId), typeof(ElementId) });

                    if (create == null)
                        throw new Exception("WallFoundation.Create API not available in this Revit version.");

                    wf = (WallFoundation)create.Invoke(null, new object[] { doc, wall.Id, wfType.Id });

                    // base offset si existe parámetro
                    if (req.baseOffset_m.HasValue && wf != null)
                    {
                        var ok = TrySetOffsetByName(wf, ToFt(req.baseOffset_m.Value));
                        // (opcional) no es error si no encontró el parámetro
                    }

                    id = wf?.Id.IntegerValue ?? 0;
                    t.Commit();
                }

                return new { elementId = id, used = new { wallId = wall.Id.IntegerValue, foundationType = wfType?.Name } };
            };
        }

        // =========================
        // struct.beamsystem.create
        // =========================
        private class BeamSystemReq
        {
            public string level { get; set; }               // requerido
            public Pt2[] profile { get; set; }              // polígono en XY (m)
            public string beamType { get; set; }            // opcional: "Family: Type"
            public string direction { get; set; }           // opcional: "X" | "Y"
            public double? spacing_m { get; set; }          // opcional (si no, usa del tipo)
            public bool? is3D { get; set; }                 // opcional
        }

        public static Func<UIApplication, object> BeamSystemCreate(JObject args)
        {
            var req = args.ToObject<BeamSystemReq>() ?? throw new Exception("Invalid args for struct.beamsystem.create.");

            return (UIApplication app) =>
            {
                var uidoc = app.ActiveUIDocument ?? throw new Exception("No active document.");
                var doc = uidoc.Document;
                var active = uidoc.ActiveView;

                var level = ResolveLevel(doc, req.level, active);
                var curves = BuildClosedProfileXY(req.profile, level.Elevation);

                // Dirección
                XYZ dir = null;
                if (!string.IsNullOrWhiteSpace(req.direction))
                {
                    if (req.direction.Equals("X", StringComparison.OrdinalIgnoreCase)) dir = XYZ.BasisX;
                    else if (req.direction.Equals("Y", StringComparison.OrdinalIgnoreCase)) dir = XYZ.BasisY;
                }
                if (dir == null)
                {
                    // heurística: usar la primera arista del perfil
                    var first = curves.First() as Line;
                    dir = (first != null) ? first.Direction : XYZ.BasisX;
                }

                // Tipo de viga para el sistema (opcional)
                var beamSym = FindSymbolByName(doc, BuiltInCategory.OST_StructuralFraming, req.beamType);

                int id = 0;
                using (var t = new Transaction(doc, "MCP: Struct.BeamSystem.Create"))
                {
                    t.Start();

                    // SketchPlane en el nivel
                    var plane = Plane.CreateByNormalAndOrigin(XYZ.BasisZ, new XYZ(0, 0, level.Elevation));
                    var sp = SketchPlane.Create(doc, plane);

                    // Intentar varios overloads de BeamSystem.Create (según versión)
                    BeamSystem bs = null;
                    var bsType = typeof(BeamSystem);
                    var methods = bsType.GetMethods(System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Static)
                        .Where(m => m.Name == "Create").ToList();

                    bool created = false;

                    foreach (var m in methods)
                    {
                        var ps = m.GetParameters();
                        try
                        {
                            object obj = null;

                            // (Document, IList<Curve>, SketchPlane, XYZ, bool)
                            if (!created && ps.Length == 5 &&
                                ps[0].ParameterType == typeof(Document) &&
                                typeof(System.Collections.Generic.IList<Curve>).IsAssignableFrom(ps[1].ParameterType) &&
                                ps[2].ParameterType == typeof(SketchPlane) &&
                                ps[3].ParameterType == typeof(XYZ) &&
                                ps[4].ParameterType == typeof(bool))
                            {
                                obj = m.Invoke(null, new object[] { doc, curves, sp, dir, req.is3D ?? false });
                                bs = obj as BeamSystem; created = bs != null;
                            }

                            // (Document, IList<Curve>, ElementId viewId, XYZ)
                            if (!created && ps.Length == 4 &&
                                ps[0].ParameterType == typeof(Document) &&
                                typeof(System.Collections.Generic.IList<Curve>).IsAssignableFrom(ps[1].ParameterType) &&
                                ps[2].ParameterType == typeof(ElementId) &&
                                ps[3].ParameterType == typeof(XYZ))
                            {
                                obj = m.Invoke(null, new object[] { doc, curves, active?.Id ?? ElementId.InvalidElementId, dir });
                                bs = obj as BeamSystem; created = bs != null;
                            }

                            // (Document, CurveArray, SketchPlane, bool)  — APIs antiguas
                            if (!created && ps.Length == 4 &&
                                ps[0].ParameterType == typeof(Document) &&
                                ps[1].ParameterType.Name.Contains("CurveArray") &&
                                ps[2].ParameterType == typeof(SketchPlane) &&
                                ps[3].ParameterType == typeof(bool))
                            {
                                // Construir CurveArray vía reflexión
                                var curveArrayType = ps[1].ParameterType;
                                dynamic ca = Activator.CreateInstance(curveArrayType);
                                foreach (var c in curves) ca.Append((dynamic)c);
                                obj = m.Invoke(null, new object[] { doc, ca, sp, req.is3D ?? false });
                                bs = obj as BeamSystem; created = bs != null;
                            }

                            if (created) break;
                        }
                        catch { /* probar siguiente overload */ }
                    }

                    if (!created)
                        throw new Exception("BeamSystem.Create API not available/compatible in this Revit version.");

                    // Intentar asignar spacing si se solicitó
                    if (req.spacing_m.HasValue)
                    {
                        try
                        {
                            var acc = bs?.LayoutRule; // puede ser null
                            // Versiones con API de acceso:
                            var propSpacing = bs?.GetType().GetProperty("Spacing");
                            if (propSpacing != null && propSpacing.CanWrite)
                                propSpacing.SetValue(bs, ToFt(req.spacing_m.Value));
                            else
                            {
                                // shape-driven accessor (según versión)
                                var gAcc = bs?.GetType().GetMethod("GetBeamSystemSpacing");
                                if (gAcc != null) gAcc.Invoke(bs, new object[] { ToFt(req.spacing_m.Value) });
                            }
                        }
                        catch { /* no fatal */ }
                    }

                    // Intentar asignar tipo de viga a las vigas creadas
                    if (beamSym != null)
                    {
                        try
                        {
                            var idsProp = bs.GetType().GetProperty("BeamIds");
                            var ids = (idsProp != null) ? (idsProp.GetValue(bs) as ICollection<ElementId>) : null;
                            if (ids != null)
                            {
                                foreach (var idBeam in ids)
                                {
                                    var fam = doc.GetElement(idBeam) as FamilyInstance;
                                    if (fam != null && !Equals(fam.Symbol?.Id, beamSym.Id))
                                        fam.Symbol = beamSym;
                                }
                            }
                        }
                        catch { /* best-effort */ }
                    }

                    id = bs.Id.IntegerValue;
                    t.Commit();
                }

                return new
                {
                    elementId = id,
                    used = new
                    {
                        level = level.Name,
                        beamType = beamSym != null ? $"{beamSym.FamilyName}: {beamSym.Name}" : "(system default)",
                        spacing_m = req.spacing_m,
                        direction = req.direction ?? "auto"
                    }
                };
            };
        }

        // =========================
        // (OPCIONAL) struct.rebar.add_straight_on_beam
        // Best-effort: coloca barras rectas a lo largo del eje de una viga
        // =========================
        private class RebarOnBeamReq
        {
            public int hostId { get; set; }                 // viga con LocationCurve
            public string barType { get; set; }             // nombre de RebarBarType (opcional)
            public int? number { get; set; }                // default 1
            public double? spacing_m { get; set; }          // si number>1 y se quiere espaciar
            public double? zOffset_m { get; set; }          // offset vertical simple
        }

        public static Func<UIApplication, object> RebarAddStraightOnBeam(JObject args)
        {
            var req = args.ToObject<RebarOnBeamReq>() ?? new RebarOnBeamReq();
            return (UIApplication app) =>
            {
                var uidoc = app.ActiveUIDocument ?? throw new Exception("No active document.");
                var doc = uidoc.Document;

                var host = doc.GetElement(new ElementId(req.hostId)) as FamilyInstance
                           ?? throw new Exception($"Host {req.hostId} not found or not FamilyInstance.");
                if (host.Category?.Id.IntegerValue != (int)BuiltInCategory.OST_StructuralFraming)
                    throw new Exception("Host must be a Structural Framing (beam).");

                var lc = host.Location as LocationCurve;
                if (lc?.Curve == null) throw new Exception("Host beam has no LocationCurve.");

                var p0 = lc.Curve.GetEndPoint(0);
                var p1 = lc.Curve.GetEndPoint(1);

                // Offset Z simple (no local axes)
                if (req.zOffset_m.HasValue)
                {
                    var dz = ToFt(req.zOffset_m.Value);
                    p0 = new XYZ(p0.X, p0.Y, p0.Z + dz);
                    p1 = new XYZ(p1.X, p1.Y, p1.Z + dz);
                }

                var barTypes = new FilteredElementCollector(doc).OfClass(typeof(RebarBarType)).Cast<RebarBarType>().ToList();
                var bt = string.IsNullOrWhiteSpace(req.barType)
                    ? barTypes.FirstOrDefault()
                    : barTypes.FirstOrDefault(b => b.Name.Equals(req.barType, StringComparison.OrdinalIgnoreCase));
                if (bt == null) throw new Exception("No RebarBarType found.");

                int id;
                using (var t = new Transaction(doc, "MCP: Struct.Rebar.AddStraightOnBeam"))
                {
                    t.Start();

                    var line = Line.CreateBound(p0, p1);
                    var curves = new List<Curve> { line };
                    var normal = XYZ.BasisZ; // best-effort

                    var rebar = Rebar.CreateFromCurves(
                        doc,
                        RebarStyle.Standard,
                        bt,
                        null, null,
                        host,
                        normal,
                        curves,
                        RebarHookOrientation.Left,
                        RebarHookOrientation.Left,
                        true, true);

                    // Layout
                    var n = Math.Max(1, req.number ?? 1);
                    if (n > 1)
                    {
                        var acc = rebar.GetShapeDrivenAccessor();
                        if (req.spacing_m.HasValue)
                            acc.SetLayoutAsNumberWithSpacing(n, ToFt(req.spacing_m.Value), true, true, true);
                        else
                            acc.SetLayoutAsFixedNumber(n, 0, true, true, true);
                    }

                    id = rebar.Id.IntegerValue;
                    t.Commit();
                }

                return new { elementId = id, hostId = host.Id.IntegerValue, barType = bt.Name };
            };
        }
    }
}

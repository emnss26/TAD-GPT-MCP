using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using mcp_app.Core;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Policy;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using static System.Net.Mime.MediaTypeNames;

namespace mcp_app.Actions
{
    internal class QtoActions
    {
        // ========= Helpers de conversión =========
        static double FtToM(double ft) => UnitUtils.ConvertFromInternalUnits(ft, UnitTypeId.Meters);
        static double Ft2ToM2(double ft2) => UnitUtils.ConvertFromInternalUnits(ft2, UnitTypeId.SquareMeters);
        static double Ft3ToM3(double ft3) => UnitUtils.ConvertFromInternalUnits(ft3, UnitTypeId.CubicMeters);
        static double FtToMm(double ft) => UnitUtils.ConvertFromInternalUnits(ft, UnitTypeId.Millimeters);

        static double? GetDoubleParam(Element e, BuiltInParameter bip)
        {
            var p = e?.get_Parameter(bip);
            if (p == null || p.StorageType != StorageType.Double) return null;
            try { return p.AsDouble(); } catch { return null; }
        }

        static double? GetAnyVolume(Element e)
        {
            // 1) HOST_VOLUME_COMPUTED
            var v = GetDoubleParam(e, BuiltInParameter.HOST_VOLUME_COMPUTED);
            if (v.HasValue && v.Value > 1e-9) return v;

            // 2) Fallback: sumar volúmenes de sólidos
            try
            {
                var opt = new Options
                {
                    ComputeReferences = false,
                    DetailLevel = ViewDetailLevel.Fine,
                    IncludeNonVisibleObjects = true
                };

                var ge = e.get_Geometry(opt);
                if (ge == null) return null;

                double sum = 0.0;
                foreach (var obj in ge)
                    sum += VolumeOf(obj);

                return sum > 1e-9 ? sum : (double?)null;
            }
            catch { return null; }
        }

        static double VolumeOf(GeometryObject go)
        {
            double acc = 0.0;

            if (go is Solid s)
            {
                if (s.Volume > 1e-9) acc += s.Volume;
            }
            else if (go is GeometryInstance gi)
            {
                var inst = gi.GetInstanceGeometry();
                if (inst != null)
                    foreach (var obj in inst)
                        acc += VolumeOf(obj);
            }
            else if (go is GeometryElement ge)
            {
                foreach (var obj in ge)
                    acc += VolumeOf(obj);
            }

            return acc;
        }

        static double? GetAnyLengthFt(Element e)
        {
            // 1) Si el elemento tiene curva (vigas, barandales, tuberías, ductos, etc.)
            try
            {
                if (e?.Location is LocationCurve lc && lc.Curve != null)
                    return lc.Curve.Length;
            }
            catch { }

            // 2) Parámetros de longitud comunes y estables entre versiones
            var bipCandidates = new[]
            {
        BuiltInParameter.CURVE_ELEM_LENGTH,   // largo de elementos con curva
        BuiltInParameter.INSTANCE_LENGTH_PARAM, // muchas familias reportan longitud aquí
        BuiltInParameter.FAMILY_HEIGHT_PARAM   // a veces se usa como "altura/longitud" en familias
    };

            foreach (var bip in bipCandidates)
            {
                try
                {
                    var p = e?.get_Parameter(bip);
                    if (p != null && p.StorageType == StorageType.Double)
                    {
                        var d = p.AsDouble();
                        if (d > 1e-9) return d;
                    }
                }
                catch { /* seguir probando */ }
            }

            // 3) Fallback geométrico para columnas: altura = delta Z del bounding box
            try
            {
                if (e?.Category?.Id?.IntegerValue == (int)BuiltInCategory.OST_StructuralColumns)
                {
                    var bb = e.get_BoundingBox(null);
                    if (bb != null)
                    {
                        var dz = Math.Abs(bb.Max.Z - bb.Min.Z);
                        if (dz > 1e-9) return dz;
                    }
                }
            }
            catch { }

            return null;
        }

        static string GetTypeName(Element el, Document doc)
        {
            var tid = el.GetTypeId();
            if (tid == ElementId.InvalidElementId) return null;
            return doc.GetElement(tid)?.Name;
        }

        static string GetLevelName(Element el, Document doc)
        {
            ElementId levelId = ElementId.InvalidElementId;

            // Instancia de familia
            if (el is FamilyInstance fi && fi.LevelId != null && fi.LevelId != ElementId.InvalidElementId)
                levelId = fi.LevelId;
            else if (el.LevelId != null && el.LevelId != ElementId.InvalidElementId)
                levelId = el.LevelId;

            // MEP Curves suelen tener RBS_START_LEVEL_PARAM
            if (levelId == ElementId.InvalidElementId)
            {
                var p = el.get_Parameter(BuiltInParameter.RBS_START_LEVEL_PARAM);
                if (p != null && p.StorageType == StorageType.ElementId)
                {
                    try { levelId = p.AsElementId(); } catch { }
                }
            }

            if (levelId == ElementId.InvalidElementId) return null;
            return doc.GetElement(levelId)?.Name;
        }

        static string GetCreatedPhaseName(Element el, Document doc)
        {
            var p = el.get_Parameter(BuiltInParameter.PHASE_CREATED);
            if (p == null || p.StorageType != StorageType.ElementId) return null;
            try
            {
                var id = p.AsElementId();
                if (id == ElementId.InvalidElementId) return null;
                return (doc.GetElement(id) as Phase)?.Name;
            }
            catch { return null; }
        }

        static string KeyFromGroupBy(Dictionary<string, string> parts, string[] groupBy)
        {
            if (groupBy == null || groupBy.Length == 0) return "(total)";
            var vals = new List<string>();
            foreach (var g in groupBy)
            {
                parts.TryGetValue(g ?? "", out var v);
                vals.Add($"{g}:{(v ?? "-")}");
            }
            return string.Join("|", vals);
        }

        static object KeyObjFromGroupBy(Dictionary<string, string> parts, string[] groupBy)
        {
            if (groupBy == null || groupBy.Length == 0) return new { key = "(total)" };
            var dict = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
            foreach (var g in groupBy) dict[g] = parts.TryGetValue(g, out var v) ? (object)(v ?? "") : "";
            return dict;
        }

        static bool TryParseBuiltInCategory(string token, out BuiltInCategory bic)
        {
            bic = default;
            if (string.IsNullOrWhiteSpace(token)) return false;
            if (Enum.TryParse(token, true, out bic)) return true;

            var norm = token.StartsWith("OST_", StringComparison.OrdinalIgnoreCase) ? token : "OST_" + token;
            return Enum.TryParse(norm, true, out bic);
        }

        // =========================================================
        // qto.walls
        // Args: { groupBy?: ("type"|"level"|"phase")[], includeIds?: bool, filter?:{...} }
        // =========================================================

        public class QtoWallsFilter
        {
            public int[] typeIds { get; set; } = Array.Empty<int>();
            public string[] typeNames { get; set; } = Array.Empty<string>();
            public string nameRegex { get; set; } // regex contra el name del tipo
            public string[] levels { get; set; } = Array.Empty<string>();
            public string phase { get; set; }     // nombre de fase creada
            public bool useSelection { get; set; } = false;
        }

        public class QtoWallsReq
        {
            public string[] groupBy { get; set; } = Array.Empty<string>();
            public bool includeIds { get; set; } = false;
            public QtoWallsFilter filter { get; set; } = null;
        }

        public static Func<UIApplication, object> QtoWalls(JObject args)
        {
            var req = args.ToObject<QtoWallsReq>() ?? new QtoWallsReq();

            var fTok = args["filter"] as JObject;

            var typeIds = fTok?["typeIds"]?.Values<int>()?.ToArray() ?? Array.Empty<int>();
            var typeNames = fTok?["typeNames"]?.Values<string>()?.Where(s => !string.IsNullOrWhiteSpace(s)).ToArray() ?? Array.Empty<string>();
            var nameRegex = fTok?["nameRegex"]?.Value<string>();
            var levels = fTok?["levels"]?.Values<string>()?.Where(s => !string.IsNullOrWhiteSpace(s)).ToArray() ?? Array.Empty<string>();
            var phaseFilter = fTok?["phase"]?.Value<string>();
            var useSelection = fTok?["useSelection"]?.Value<bool?>() ?? false;

            var typeNameSet = typeNames.Length > 0 ? new HashSet<string>(typeNames, StringComparer.OrdinalIgnoreCase) : null;
            var levelSet = levels.Length > 0 ? new HashSet<string>(levels, StringComparer.OrdinalIgnoreCase) : null;
            System.Text.RegularExpressions.Regex rx = null;
            if (!string.IsNullOrWhiteSpace(nameRegex))
            {
                try { rx = new System.Text.RegularExpressions.Regex(nameRegex, System.Text.RegularExpressions.RegexOptions.IgnoreCase); } catch { rx = null; }
            }

            return (UIApplication app) =>
            {
                var uidoc = app.ActiveUIDocument ?? throw new Exception("No active document.");
                var doc = uidoc.Document ?? throw new Exception("No active document.");

                IEnumerable<Wall> wallsEnum = new FilteredElementCollector(doc)
                    .OfClass(typeof(Wall)).Cast<Wall>()
                    .Where(w => w.Category != null && w.Category.Id.IntegerValue == (int)BuiltInCategory.OST_Walls);

                if (useSelection)
                {
                    var selIds = uidoc.Selection?.GetElementIds() ?? new List<ElementId>();
                    var selSet = new HashSet<int>(selIds.Select(id => id.IntegerValue));
                    wallsEnum = wallsEnum.Where(w => selSet.Contains(w.Id.IntegerValue));
                }

                var walls = wallsEnum.ToList();

                if (typeIds.Length > 0 || (typeNameSet?.Count ?? 0) > 0 || rx != null || (levelSet?.Count ?? 0) > 0 || !string.IsNullOrWhiteSpace(phaseFilter))
                {
                    walls = walls.Where(w =>
                    {
                        var typeId = w.GetTypeId();
                        int tid = (typeId != null && typeId != ElementId.InvalidElementId) ? typeId.IntegerValue : -1;
                        var tname = GetTypeName(w, doc) ?? "";
                        var lvl = GetLevelName(w, doc) ?? "";
                        var ph = GetCreatedPhaseName(w, doc) ?? "";

                        bool pass = true;

                        if (typeIds.Length > 0) pass &= typeIds.Contains(tid);
                        if ((typeNameSet?.Count ?? 0) > 0) pass &= typeNameSet.Contains(tname);
                        if (rx != null) pass &= rx.IsMatch(tname);
                        if ((levelSet?.Count ?? 0) > 0) pass &= levelSet.Contains(lvl);
                        if (!string.IsNullOrWhiteSpace(phaseFilter)) pass &= string.Equals(phaseFilter, ph, StringComparison.OrdinalIgnoreCase);

                        return pass;
                    }).ToList();
                }

                double totalLenM = 0, totalAreaM2 = 0, totalVolM3 = 0;
                var groups = new Dictionary<string, (double len, double area, double vol, int count, object keyObj)>();
                var rows = new List<object>();

                foreach (var w in walls)
                {
                    var len_ft = (w.Location as LocationCurve)?.Curve?.Length ?? 0.0;
                    var area_ft2 = GetDoubleParam(w, BuiltInParameter.HOST_AREA_COMPUTED) ?? 0.0;
                    var vol_ft3 = GetDoubleParam(w, BuiltInParameter.HOST_VOLUME_COMPUTED) ?? 0.0;

                    var len_m = FtToM(len_ft);
                    var area_m2 = Ft2ToM2(area_ft2);
                    var vol_m3 = Ft3ToM3(vol_ft3);

                    totalLenM += len_m;
                    totalAreaM2 += area_m2;
                    totalVolM3 += vol_m3;

                    var type = GetTypeName(w, doc);
                    var level = GetLevelName(w, doc);
                    var phase = GetCreatedPhaseName(w, doc);
                    bool roomBounding = false;
                    try
                    {
                        var rb = w.get_Parameter(BuiltInParameter.WALL_ATTR_ROOM_BOUNDING);
                        if (rb != null && rb.StorageType == StorageType.Integer) roomBounding = rb.AsInteger() == 1;
                    }
                    catch { }

                    var parts = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                    {
                        ["type"] = type,
                        ["level"] = level,
                        ["phase"] = phase
                    };

                    var gk = KeyFromGroupBy(parts, req.groupBy);
                    if (!groups.TryGetValue(gk, out var acc))
                    {
                        acc = (0, 0, 0, 0, KeyObjFromGroupBy(parts, req.groupBy));
                    }
                    acc.len += len_m;
                    acc.area += area_m2;
                    acc.vol += vol_m3;
                    acc.count += 1;
                    groups[gk] = acc;

                    if (req.includeIds)
                    {
                        rows.Add(new
                        {
                            id = w.Id.IntegerValue,
                            type,
                            level,
                            phase,
                            roomBounding,
                            length_m = len_m,
                            area_m2 = area_m2,
                            volume_m3 = vol_m3
                        });
                    }
                }

                var groupRows = groups.Select(kv => new
                {
                    key = kv.Value.keyObj,
                    count = kv.Value.count,
                    length_m = kv.Value.len,
                    area_m2 = kv.Value.area,
                    volume_m3 = kv.Value.vol
                }).ToList();

                return new
                {
                    summary = new
                    {
                        totalCount = walls.Count,
                        totalLength_m = totalLenM,
                        totalArea_m2 = totalAreaM2,
                        totalVolume_m3 = totalVolM3
                    },
                    groups = groupRows,
                    rows = req.includeIds ? rows : null
                };
            };
        }

        // =========================================================
        // qto.floors
        // Args: { groupBy?: ("type"|"level")[], includeIds?: bool }
        // =========================================================
        public class QtoFloorsReq
        {
            public string[] groupBy { get; set; } = Array.Empty<string>();
            public bool includeIds { get; set; } = false;
        }

        public static Func<UIApplication, object> QtoFloors(JObject args)
        {
            var req = args.ToObject<QtoFloorsReq>() ?? new QtoFloorsReq();
            return (UIApplication app) =>
            {
                var doc = app.ActiveUIDocument?.Document ?? throw new Exception("No active document.");

                var floors = new FilteredElementCollector(doc)
                    .OfClass(typeof(Floor)).Cast<Floor>()
                    .Where(f => f.Category != null && f.Category.Id.IntegerValue == (int)BuiltInCategory.OST_Floors)
                    .ToList();

                double totalAreaM2 = 0, totalVolM3 = 0;
                var groups = new Dictionary<string, (double area, double vol, int count, object keyObj)>();
                var rows = new List<object>();

                foreach (var f in floors)
                {
                    var area_ft2 = GetDoubleParam(f, BuiltInParameter.HOST_AREA_COMPUTED) ?? 0.0;
                    var vol_ft3 = GetAnyVolume(f) ?? 0.0;

                    var area_m2 = Ft2ToM2(area_ft2);
                    var vol_m3 = Ft3ToM3(vol_ft3);

                    totalAreaM2 += area_m2;
                    totalVolM3 += vol_m3;

                    var type = GetTypeName(f, doc);
                    var level = GetLevelName(f, doc);

                    var parts = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                    {
                        ["type"] = type,
                        ["level"] = level
                    };

                    var gk = KeyFromGroupBy(parts, req.groupBy);
                    if (!groups.TryGetValue(gk, out var acc))
                    {
                        acc = (0, 0, 0, KeyObjFromGroupBy(parts, req.groupBy));
                    }
                    acc.area += area_m2;
                    acc.vol += vol_m3;
                    acc.count += 1;
                    groups[gk] = acc;

                    if (req.includeIds)
                    {
                        rows.Add(new
                        {
                            id = f.Id.IntegerValue,
                            type,
                            level,
                            area_m2 = area_m2,
                            volume_m3 = vol_m3
                        });
                    }
                }

                var groupRows = groups.Select(kv => new
                {
                    key = kv.Value.keyObj,
                    count = kv.Value.count,
                    area_m2 = kv.Value.area,
                    volume_m3 = kv.Value.vol
                }).ToList();

                return new
                {
                    summary = new
                    {
                        totalCount = floors.Count,
                        totalArea_m2 = totalAreaM2,
                        totalVolume_m3 = totalVolM3
                    },
                    groups = groupRows,
                    rows = req.includeIds ? rows : null
                };
            };
        }

        // =========================================================
        // qto.ceilings (plafones) – SOLO ÁREA
        // Args: { groupBy?: ("type"|"level")[], includeIds?: bool }
        // =========================================================
        public class QtoCeilingsReq
        {
            public string[] groupBy { get; set; } = Array.Empty<string>();
            public bool includeIds { get; set; } = false;
        }

        public static Func<UIApplication, object> QtoCeilings(JObject args)
        {
            var req = args.ToObject<QtoCeilingsReq>() ?? new QtoCeilingsReq();
            return (UIApplication app) =>
            {
                var doc = app.ActiveUIDocument?.Document ?? throw new Exception("No active document.");

                var items = new FilteredElementCollector(doc)
                    .OfClass(typeof(Ceiling)).Cast<Ceiling>()
                    .Where(c => c.Category != null && c.Category.Id.IntegerValue == (int)BuiltInCategory.OST_Ceilings)
                    .ToList();

                double totalArea = 0;
                var groups = new Dictionary<string, (double area, int count, object keyObj)>();
                var rows = new List<object>();

                foreach (var c in items)
                {
                    var area_m2 = Ft2ToM2(GetDoubleParam(c, BuiltInParameter.HOST_AREA_COMPUTED) ?? 0.0);
                    totalArea += area_m2;

                    var type = GetTypeName(c, doc);
                    var level = GetLevelName(c, doc);

                    var parts = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                    {
                        ["type"] = type,
                        ["level"] = level
                    };

                    var gk = KeyFromGroupBy(parts, req.groupBy);
                    if (!groups.TryGetValue(gk, out var acc))
                        acc = (0, 0, KeyObjFromGroupBy(parts, req.groupBy));
                    acc.area += area_m2; acc.count++;
                    groups[gk] = acc;

                    if (req.includeIds)
                        rows.Add(new { id = c.Id.IntegerValue, type, level, area_m2 = area_m2 });
                }

                var groupRows = groups.Select(k => new { key = k.Value.keyObj, count = k.Value.count, area_m2 = k.Value.area }).ToList();

                return new
                {
                    summary = new { totalCount = items.Count, totalArea_m2 = totalArea },
                    groups = groupRows,
                    rows = req.includeIds ? rows : null
                };
            };
        }

        // =========================================================
        // qto.railings (barandales) – m.l.
        // Args: { groupBy?: ("type"|"level")[], includeIds?: bool }
        // =========================================================
        public class QtoRailingsReq
        {
            public string[] groupBy { get; set; } = Array.Empty<string>();
            public bool includeIds { get; set; } = false;
        }

        public static Func<UIApplication, object> QtoRailings(JObject args)
        {
            var req = args.ToObject<QtoRailingsReq>() ?? new QtoRailingsReq();
            return (UIApplication app) =>
            {
                var doc = app.ActiveUIDocument?.Document ?? throw new Exception("No active document.");

                var rails = new FilteredElementCollector(doc)
                    .OfClass(typeof(Railing)).Cast<Railing>()
                    .ToList();

                var groups = new Dictionary<string, (double ml, int count, object keyObj)>();
                var rows = new List<object>();

                foreach (var r in rails)
                {
                    var len_m = FtToM(GetAnyLengthFt(r) ?? 0.0);
                    var type = GetTypeName(r, doc);
                    var level = GetLevelName(r, doc);

                    var parts = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                    {
                        ["type"] = type,
                        ["level"] = level
                    };

                    var gk = KeyFromGroupBy(parts, req.groupBy);
                    if (!groups.TryGetValue(gk, out var acc))
                        acc = (0, 0, KeyObjFromGroupBy(parts, req.groupBy));
                    acc.ml += len_m; acc.count++;
                    groups[gk] = acc;

                    if (req.includeIds)
                        rows.Add(new { id = r.Id.IntegerValue, type, level, length_m = len_m });
                }

                var groupRows = groups.Select(k => new { key = k.Value.keyObj, count = k.Value.count, length_m = k.Value.ml }).ToList();
                return new { groups = groupRows, rows = req.includeIds ? rows : null };
            };
        }

        // =========================================================
        // qto.struct.beams – opciones m.l. y/o volumen (sin material)
        // Args: { groupBy?: ("type"|"level"), includeIds?: bool, includeLength?: bool=true, includeVolume?: bool=true }
        // =========================================================
        public class QtoStructBeamsReq
        {
            public string[] groupBy { get; set; } = Array.Empty<string>();
            public bool includeIds { get; set; } = false;
            public bool includeLength { get; set; } = true;
            public bool includeVolume { get; set; } = true;
        }

        public static Func<UIApplication, object> QtoStructBeams(JObject args)
        {
            var req = args.ToObject<QtoStructBeamsReq>() ?? new QtoStructBeamsReq();
            return (UIApplication app) =>
            {
                var doc = app.ActiveUIDocument?.Document ?? throw new Exception("No active document.");

                var elems = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_StructuralFraming)
                    .WhereElementIsNotElementType()
                    .ToElements();

                var groups = new Dictionary<string, (double ml, double vol, int count, object keyObj)>();
                var rows = new List<object>();

                foreach (var e in elems)
                {
                    double len_m = req.includeLength ? FtToM(GetAnyLengthFt(e) ?? 0.0) : 0.0;
                    double vol_m3 = req.includeVolume ? Ft3ToM3(GetAnyVolume(e) ?? 0.0) : 0.0;
                    var type = GetTypeName(e, doc);
                    var level = GetLevelName(e, doc);

                    var parts = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                    {
                        ["type"] = type,
                        ["level"] = level
                    };

                    var gk = KeyFromGroupBy(parts, req.groupBy);
                    if (!groups.TryGetValue(gk, out var acc))
                        acc = (0, 0, 0, KeyObjFromGroupBy(parts, req.groupBy));
                    acc.ml += len_m; acc.vol += vol_m3; acc.count++;
                    groups[gk] = acc;

                    if (req.includeIds)
                        rows.Add(new { id = e.Id.IntegerValue, type, level, length_m = len_m, volume_m3 = vol_m3 });
                }

                var groupRows = groups.Select(k => new { key = k.Value.keyObj, count = k.Value.count, length_m = k.Value.ml, volume_m3 = k.Value.vol }).ToList();
                return new { groups = groupRows, rows = req.includeIds ? rows : null };
            };
        }

        // =========================================================
        // qto.struct.columns – opciones m.l. y/o volumen (sin material)
        // Args: { groupBy?: ("type"|"level"), includeIds?: bool, includeLength?: bool=true, includeVolume?: bool=true }
        // =========================================================
        public class QtoStructColumnsReq
        {
            public string[] groupBy { get; set; } = Array.Empty<string>();
            public bool includeIds { get; set; } = false;
            public bool includeLength { get; set; } = true;
            public bool includeVolume { get; set; } = true;
        }

        public static Func<UIApplication, object> QtoStructColumns(JObject args)
        {
            var req = args.ToObject<QtoStructColumnsReq>() ?? new QtoStructColumnsReq();
            return (UIApplication app) =>
            {
                var doc = app.ActiveUIDocument?.Document ?? throw new Exception("No active document.");

                var elems = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_StructuralColumns)
                    .WhereElementIsNotElementType()
                    .ToElements();

                var groups = new Dictionary<string, (double ml, double vol, int count, object keyObj)>();
                var rows = new List<object>();

                foreach (var e in elems)
                {
                    double len_m = req.includeLength ? FtToM(GetAnyLengthFt(e) ?? 0.0) : 0.0;
                    double vol_m3 = req.includeVolume ? Ft3ToM3(GetAnyVolume(e) ?? 0.0) : 0.0;
                    var type = GetTypeName(e, doc);
                    var level = GetLevelName(e, doc);

                    var parts = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                    {
                        ["type"] = type,
                        ["level"] = level
                    };

                    var gk = KeyFromGroupBy(parts, req.groupBy);
                    if (!groups.TryGetValue(gk, out var acc))
                        acc = (0, 0, 0, KeyObjFromGroupBy(parts, req.groupBy));
                    acc.ml += len_m; acc.vol += vol_m3; acc.count++;
                    groups[gk] = acc;

                    if (req.includeIds)
                        rows.Add(new { id = e.Id.IntegerValue, type, level, length_m = len_m, volume_m3 = vol_m3 });
                }

                var groupRows = groups.Select(k => new { key = k.Value.keyObj, count = k.Value.count, length_m = k.Value.ml, volume_m3 = k.Value.vol }).ToList();
                return new { groups = groupRows, rows = req.includeIds ? rows : null };
            };
        }

        // =========================================================
        // qto.struct.foundations – conteo y volumen
        // Args: { groupBy?: ("type"|"level"), includeIds?: bool }
        // =========================================================
        public class QtoStructFoundationsReq
        {
            public string[] groupBy { get; set; } = Array.Empty<string>();
            public bool includeIds { get; set; } = false;
        }

        public static Func<UIApplication, object> QtoStructFoundations(JObject args)
        {
            var req = args.ToObject<QtoStructFoundationsReq>() ?? new QtoStructFoundationsReq();
            return (UIApplication app) =>
            {
                var doc = app.ActiveUIDocument?.Document ?? throw new Exception("No active document.");

                var elems = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_StructuralFoundation)
                    .WhereElementIsNotElementType()
                    .ToElements();

                var groups = new Dictionary<string, (int count, double vol, object keyObj)>();
                var rows = new List<object>();

                foreach (var e in elems)
                {
                    var vol_m3 = Ft3ToM3(GetAnyVolume(e) ?? 0.0);
                    var type = GetTypeName(e, doc);
                    var level = GetLevelName(e, doc);

                    var parts = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                    {
                        ["type"] = type,
                        ["level"] = level
                    };

                    var gk = KeyFromGroupBy(parts, req.groupBy);
                    if (!groups.TryGetValue(gk, out var acc))
                        acc = (0, 0, KeyObjFromGroupBy(parts, req.groupBy));
                    acc.count++; acc.vol += vol_m3;
                    groups[gk] = acc;

                    if (req.includeIds)
                        rows.Add(new { id = e.Id.IntegerValue, type, level, volume_m3 = vol_m3 });
                }

                var groupRows = groups.Select(k => new { key = k.Value.keyObj, count = k.Value.count, volume_m3 = k.Value.vol }).ToList();
                return new { groups = groupRows, rows = req.includeIds ? rows : null };
            };
        }

        // =========================================================
        // qto.families.count – conteo por familia/tipo/categoría/level
        // Args: { groupBy?: ("category"|"family"|"type"|"level")[], includeIds?: bool, categories?: string[] (OST_* o sin OST_) }
        // =========================================================
        public class QtoFamiliesCountReq
        {
            public string[] groupBy { get; set; } = Array.Empty<string>();
            public bool includeIds { get; set; } = false;
            public string[] categories { get; set; } = Array.Empty<string>();
        }

        public static Func<UIApplication, object> QtoFamiliesCount(JObject args)
        {
            var req = args.ToObject<QtoFamiliesCountReq>() ?? new QtoFamiliesCountReq();
            return (UIApplication app) =>
            {
                var doc = app.ActiveUIDocument?.Document ?? throw new Exception("No active document.");

                IEnumerable<Element> insts = new FilteredElementCollector(doc)
                    .OfClass(typeof(FamilyInstance))
                    .WhereElementIsNotElementType()
                    .Cast<Element>();

                var bicSet = new HashSet<BuiltInCategory>();
                foreach (var t in (req.categories ?? Array.Empty<string>()))
                    if (TryParseBuiltInCategory(t, out var bic)) bicSet.Add(bic);

                if (bicSet.Count > 0)
                {
                    insts = insts.Where(e =>
                    {
                        var cid = e.Category?.Id?.IntegerValue ?? int.MinValue;
                        try { return bicSet.Contains((BuiltInCategory)cid); } catch { return false; }
                    });
                }

                var groups = new Dictionary<string, (int count, object keyObj)>();
                var rows = new List<object>();

                foreach (var e in insts)
                {
                    var cat = e.Category?.Name;
                    var type = GetTypeName(e, doc);
                    var level = GetLevelName(e, doc);
                    string fam = null;
                    try { fam = (e as FamilyInstance)?.Symbol?.Family?.Name; } catch { }

                    var parts = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                    {
                        ["category"] = cat,
                        ["family"] = fam,
                        ["type"] = type,
                        ["level"] = level
                    };

                    var gk = KeyFromGroupBy(parts, req.groupBy);
                    if (!groups.TryGetValue(gk, out var acc))
                        acc = (0, KeyObjFromGroupBy(parts, req.groupBy));
                    acc.count++;
                    groups[gk] = acc;

                    if (req.includeIds)
                        rows.Add(new { id = e.Id.IntegerValue, category = cat, family = fam, type, level });
                }

                var groupRows = groups.Select(k => new { key = k.Value.keyObj, count = k.Value.count }).ToList();

                return new { groups = groupRows, rows = req.includeIds ? rows : null };
            };
        }

        // =========================================================
        // qto.mep.pipes – m.l. + buckets por diámetro
        // Args: { groupBy?: ("system"|"type"|"level"), diameterBucketsMm?: double[], includeIds?: bool }
        // =========================================================
        public class QtoPipesReq
        {
            public string[] groupBy { get; set; } = Array.Empty<string>();
            public double[] diameterBucketsMm { get; set; } = Array.Empty<double>();
            public bool includeIds { get; set; } = false;
        }

        public static Func<UIApplication, object> QtoMepPipes(JObject args)
        {
            var req = args.ToObject<QtoPipesReq>() ?? new QtoPipesReq();
            var buckets = (req.diameterBucketsMm ?? Array.Empty<double>()).OrderBy(x => x).ToArray();

            string BucketName(double mm)
            {
                if (buckets.Length == 0) return "all";
                foreach (var b in buckets)
                    if (mm <= b) return $"≤{b}mm";
                return $">{buckets.Last()}mm";
            }

            return (UIApplication app) =>
            {
                var doc = app.ActiveUIDocument?.Document ?? throw new Exception("No active document.");

                var pipes = new FilteredElementCollector(doc)
                    .OfClass(typeof(Pipe)).Cast<Pipe>()
                    .ToList();

                var groups = new Dictionary<string, (double ml, Dictionary<string, double> byBucket, int count, object keyObj)>();
                var rows = new List<object>();

                foreach (var p in pipes)
                {
                    var len_ft = (p.Location as LocationCurve)?.Curve?.Length ?? 0.0;
                    var len_m = FtToM(len_ft);

                    var d_ft = GetDoubleParam(p, BuiltInParameter.RBS_PIPE_DIAMETER_PARAM) ?? 0.0;
                    var d_mm = FtToMm(d_ft);
                    var bucket = BucketName(d_mm);

                    var sysName = p.MEPSystem?.Name;
                    var type = GetTypeName(p, doc);
                    var level = GetLevelName(p, doc);

                    var parts = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                    {
                        ["system"] = sysName,
                        ["type"] = type,
                        ["level"] = level
                    };

                    var gk = KeyFromGroupBy(parts, req.groupBy);
                    if (!groups.TryGetValue(gk, out var acc))
                        acc = (0, new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase), 0, KeyObjFromGroupBy(parts, req.groupBy));
                    acc.ml += len_m;
                    acc.count += 1;
                    acc.byBucket[bucket] = (acc.byBucket.TryGetValue(bucket, out var old) ? old : 0) + len_m;
                    groups[gk] = acc;

                    if (req.includeIds)
                        rows.Add(new { id = p.Id.IntegerValue, system = sysName, type, level, diameter_mm = d_mm, length_m = len_m });
                }

                var groupRows = groups.Select(kv => new
                {
                    key = kv.Value.keyObj,
                    count = kv.Value.count,
                    length_m = kv.Value.ml,
                    buckets = kv.Value.byBucket.Select(b => new { bucket = b.Key, length_m = b.Value }).OrderBy(x => x.bucket).ToList()
                }).ToList();

                return new { groups = groupRows, rows = req.includeIds ? rows : null };
            };
        }

        // =========================================================
        // qto.mep.ducts – m.l. + (opcional) área superficial estimada
        // Args: { groupBy?: ("system"|"type"|"level"), roundVsRect?: bool=false, includeIds?: bool }
        // =========================================================
        public class QtoDuctsReq
        {
            public string[] groupBy { get; set; } = Array.Empty<string>();
            public bool roundVsRect { get; set; } = false;
            public bool includeIds { get; set; } = false;
        }

        public static Func<UIApplication, object> QtoMepDucts(JObject args)
        {
            var req = args.ToObject<QtoDuctsReq>() ?? new QtoDuctsReq();

            return (UIApplication app) =>
            {
                var doc = app.ActiveUIDocument?.Document ?? throw new Exception("No active document.");

                var ducts = new FilteredElementCollector(doc)
                    .OfClass(typeof(Duct)).Cast<Duct>()
                    .ToList();

                var groups = new Dictionary<string, (double ml, double areaSurf, int count, object keyObj)>();
                var rows = new List<object>();

                foreach (var d in ducts)
                {
                    var len_ft = (d.Location as LocationCurve)?.Curve?.Length ?? 0.0;
                    var len_m = FtToM(len_ft);

                    var diam_ft = GetDoubleParam(d, BuiltInParameter.RBS_CURVE_DIAMETER_PARAM);
                    var w_ft = GetDoubleParam(d, BuiltInParameter.RBS_CURVE_WIDTH_PARAM);
                    var h_ft = GetDoubleParam(d, BuiltInParameter.RBS_CURVE_HEIGHT_PARAM);

                    bool isRound = diam_ft.HasValue && diam_ft.Value > 1e-9;
                    double areaSurf_m2 = 0.0;

                    if (isRound)
                    {
                        var d_m = FtToM(diam_ft.Value);
                        areaSurf_m2 = Math.PI * d_m * len_m; // ~ área lateral
                    }
                    else if (w_ft.HasValue && h_ft.HasValue && w_ft.Value > 0 && h_ft.Value > 0)
                    {
                        var w_m = FtToM(w_ft.Value);
                        var h_m = FtToM(h_ft.Value);
                        var per_m = 2 * (w_m + h_m);
                        areaSurf_m2 = per_m * len_m;
                    }

                    var sysName = d.MEPSystem?.Name;
                    var type = GetTypeName(d, doc);
                    var level = GetLevelName(d, doc);

                    var parts = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                    {
                        ["system"] = sysName,
                        ["type"] = type,
                        ["level"] = level
                    };
                    if (req.roundVsRect) parts["shape"] = isRound ? "round" : "rect";

                    var gb = req.groupBy.Concat(req.roundVsRect ? new[] { "shape" } : Array.Empty<string>()).ToArray();
                    var gk = KeyFromGroupBy(parts, gb);
                    if (!groups.TryGetValue(gk, out var acc))
                        acc = (0, 0, 0, KeyObjFromGroupBy(parts, gb));
                    acc.ml += len_m;
                    acc.areaSurf += areaSurf_m2;
                    acc.count += 1;
                    groups[gk] = acc;

                    if (req.includeIds)
                        rows.Add(new { id = d.Id.IntegerValue, system = sysName, type, level, shape = isRound ? "round" : "rect", length_m = len_m, surface_area_m2 = areaSurf_m2 });
                }

                var groupRows = groups.Select(kv => new
                {
                    key = kv.Value.keyObj,
                    count = kv.Value.count,
                    length_m = kv.Value.ml,
                    surface_area_m2 = kv.Value.areaSurf
                }).ToList();

                return new { groups = groupRows, rows = req.includeIds ? rows : null };
            };
        }

        // =========================================================
        // qto.cabletrays – m.l.
        // Args: { groupBy?: ("type"|"level"), includeIds?: bool }
        // =========================================================
        public class QtoCableTraysReq
        {
            public string[] groupBy { get; set; } = Array.Empty<string>();
            public bool includeIds { get; set; } = false;
        }

        public static Func<UIApplication, object> QtoCableTrays(JObject args)
        {
            var req = args.ToObject<QtoCableTraysReq>() ?? new QtoCableTraysReq();
            return (UIApplication app) =>
            {
                var doc = app.ActiveUIDocument?.Document ?? throw new Exception("No active document.");

                var trays = new FilteredElementCollector(doc)
                    .OfClass(typeof(CableTray)).Cast<CableTray>()
                    .ToList();

                var groups = new Dictionary<string, (double ml, int count, object keyObj)>();
                var rows = new List<object>();

                foreach (var t in trays)
                {
                    var len_m = FtToM(GetAnyLengthFt(t) ?? 0.0);
                    var type = GetTypeName(t, doc);
                    var level = GetLevelName(t, doc);

                    var parts = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                    {
                        ["type"] = type,
                        ["level"] = level
                    };

                    var gk = KeyFromGroupBy(parts, req.groupBy);
                    if (!groups.TryGetValue(gk, out var acc))
                        acc = (0, 0, KeyObjFromGroupBy(parts, req.groupBy));
                    acc.ml += len_m; acc.count++;
                    groups[gk] = acc;

                    if (req.includeIds)
                        rows.Add(new { id = t.Id.IntegerValue, type, level, length_m = len_m });
                }

                var groupRows = groups.Select(k => new { key = k.Value.keyObj, count = k.Value.count, length_m = k.Value.ml }).ToList();
                return new { groups = groupRows, rows = req.includeIds ? rows : null };
            };
        }

        // =========================================================
        // qto.conduits – m.l.
        // Args: { groupBy?: ("type"|"level"), includeIds?: bool }
        // =========================================================
        public class QtoConduitsReq
        {
            public string[] groupBy { get; set; } = Array.Empty<string>();
            public bool includeIds { get; set; } = false;
        }

        public static Func<UIApplication, object> QtoConduits(JObject args)
        {
            var req = args.ToObject<QtoConduitsReq>() ?? new QtoConduitsReq();
            return (UIApplication app) =>
            {
                var doc = app.ActiveUIDocument?.Document ?? throw new Exception("No active document.");

                var conduits = new FilteredElementCollector(doc)
                    .OfClass(typeof(Conduit)).Cast<Conduit>()
                    .ToList();

                var groups = new Dictionary<string, (double ml, int count, object keyObj)>();
                var rows = new List<object>();

                foreach (var c in conduits)
                {
                    var len_m = FtToM(GetAnyLengthFt(c) ?? 0.0);
                    var type = GetTypeName(c, doc);
                    var level = GetLevelName(c, doc);

                    var parts = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                    {
                        ["type"] = type,
                        ["level"] = level
                    };

                    var gk = KeyFromGroupBy(parts, req.groupBy);
                    if (!groups.TryGetValue(gk, out var acc))
                        acc = (0, 0, KeyObjFromGroupBy(parts, req.groupBy));
                    acc.ml += len_m; acc.count++;
                    groups[gk] = acc;

                    if (req.includeIds)
                        rows.Add(new { id = c.Id.IntegerValue, type, level, length_m = len_m });
                }

                var groupRows = groups.Select(k => new { key = k.Value.keyObj, count = k.Value.count, length_m = k.Value.ml }).ToList();
                return new { groups = groupRows, rows = req.includeIds ? rows : null };
            };
        }

        // =========================================================
        /* qto.counts.mep_electrical – conteos de:
           - DuctFitting, DuctTerminal (air terminals), MechanicalEquipment
           - PipeFitting, PipeAccessory
           - PlumbingEquipment, PlumbingFixtures
           - ElectricalEquipment, ElectricalFixtures
           - LightingFixtures, LightingDevices
           Args:
           {
             groupBy?: ("category"|"type"|"level"|"family")[],
             includeIds?: bool,
             categories?: string[] // si no se pasa, usa el set por defecto arriba
           }
        */
        // =========================================================
        public class QtoCountsMepElecReq
        {
            public string[] groupBy { get; set; } = Array.Empty<string>();
            public bool includeIds { get; set; } = false;
            public string[] categories { get; set; } = Array.Empty<string>();
        }

        public static Func<UIApplication, object> QtoCountsMepElectrical(JObject args)
        {
            var req = args.ToObject<QtoCountsMepElecReq>() ?? new QtoCountsMepElecReq();

            var defaultCats = new[]
            {
                BuiltInCategory.OST_DuctFitting,
                BuiltInCategory.OST_DuctTerminal,
                BuiltInCategory.OST_MechanicalEquipment,
                BuiltInCategory.OST_PipeFitting,
                BuiltInCategory.OST_PipeAccessory,
                BuiltInCategory.OST_PlumbingEquipment,
                BuiltInCategory.OST_PlumbingFixtures,
                BuiltInCategory.OST_ElectricalEquipment,
                BuiltInCategory.OST_ElectricalFixtures,
                BuiltInCategory.OST_LightingFixtures,
                BuiltInCategory.OST_LightingDevices
            };

            var cats = new List<BuiltInCategory>();
            if (req.categories != null && req.categories.Length > 0)
            {
                foreach (var t in req.categories)
                    if (TryParseBuiltInCategory(t, out var bic)) cats.Add(bic);
            }
            if (cats.Count == 0) cats.AddRange(defaultCats);

            return (UIApplication app) =>
            {
                var doc = app.ActiveUIDocument?.Document ?? throw new Exception("No active document.");
                var groups = new Dictionary<string, (int count, object keyObj)>();
                var rows = new List<object>();

                foreach (var bic in cats)
                {
                    var elems = new FilteredElementCollector(doc)
                        .OfCategory(bic)
                        .WhereElementIsNotElementType()
                        .ToElements();

                    foreach (var e in elems)
                    {
                        var category = e.Category?.Name;
                        var type = GetTypeName(e, doc);
                        var level = GetLevelName(e, doc);
                        string family = null;
                        try { family = (e as FamilyInstance)?.Symbol?.Family?.Name; } catch { }

                        var parts = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                        {
                            ["category"] = category,
                            ["family"] = family,
                            ["type"] = type,
                            ["level"] = level
                        };

                        var gk = KeyFromGroupBy(parts, req.groupBy);
                        if (!groups.TryGetValue(gk, out var acc))
                            acc = (0, KeyObjFromGroupBy(parts, req.groupBy));
                        acc.count++;
                        groups[gk] = acc;

                        if (req.includeIds)
                            rows.Add(new { id = e.Id.IntegerValue, category, family, type, level });
                    }
                }

                var groupRows = groups.Select(k => new { key = k.Value.keyObj, count = k.Value.count }).ToList();
                return new { groups = groupRows, rows = req.includeIds ? rows : null };
            };
        }

        // =========================================================
        // qto.struct.concrete (conjunto rápido) – volumen + m.l. beams
        // (se mantiene por compatibilidad; para control fino usa beams/columns/foundations)
        // Args: { includeBeams?:bool, includeColumns?:bool, includeFoundation?:bool, groupBy?: ("type"|"level"), includeIds?: bool }
        // =========================================================
        public class QtoStructConcreteReq
        {
            public bool includeBeams { get; set; } = true;
            public bool includeColumns { get; set; } = true;
            public bool includeFoundation { get; set; } = true;
            public string[] groupBy { get; set; } = Array.Empty<string>();
            public bool includeIds { get; set; } = false;
        }

        public static Func<UIApplication, object> QtoStructConcrete(JObject args)
        {
            var req = args.ToObject<QtoStructConcreteReq>() ?? new QtoStructConcreteReq();
            return (UIApplication app) =>
            {
                var doc = app.ActiveUIDocument?.Document ?? throw new Exception("No active document.");

                var rows = new List<object>();
                var groups = new Dictionary<string, (double vol, double ml, int count, object keyObj)>();

                void Acc(string type, string level, int id, double? vol_ft3, double? ml_ft, bool addMl)
                {
                    var vol_m3 = Ft3ToM3(vol_ft3 ?? 0.0);
                    var ml_m = FtToM(addMl ? (ml_ft ?? 0.0) : 0.0);

                    var parts = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                    {
                        ["type"] = type,
                        ["level"] = level
                    };
                    var gk = KeyFromGroupBy(parts, req.groupBy);
                    if (!groups.TryGetValue(gk, out var acc))
                        acc = (0, 0, 0, KeyObjFromGroupBy(parts, req.groupBy));
                    acc.vol += vol_m3;
                    acc.ml += ml_m;
                    acc.count += 1;
                    groups[gk] = acc;

                    if (req.includeIds)
                        rows.Add(new { id, type, level, volume_m3 = vol_m3, ml_m = ml_m });
                }

                if (req.includeBeams)
                {
                    var beams = new FilteredElementCollector(doc)
                        .OfCategory(BuiltInCategory.OST_StructuralFraming)
                        .WhereElementIsNotElementType()
                        .ToElements();

                    foreach (var b in beams)
                    {
                        var type = GetTypeName(b, doc);
                        var level = GetLevelName(b, doc);
                        var vol = GetAnyVolume(b);
                        var len = (b.Location as LocationCurve)?.Curve?.Length;
                        Acc(type, level, b.Id.IntegerValue, vol, len, addMl: true);
                    }
                }

                if (req.includeColumns)
                {
                    var cols = new FilteredElementCollector(doc)
                        .OfCategory(BuiltInCategory.OST_StructuralColumns)
                        .WhereElementIsNotElementType()
                        .ToElements();

                    foreach (var c in cols)
                    {
                        var type = GetTypeName(c, doc);
                        var level = GetLevelName(c, doc);
                        var vol = GetAnyVolume(c);
                        Acc(type, level, c.Id.IntegerValue, vol, ml_ft: null, addMl: false);
                    }
                }

                if (req.includeFoundation)
                {
                    var fnds = new FilteredElementCollector(doc)
                        .OfCategory(BuiltInCategory.OST_StructuralFoundation)
                        .WhereElementIsNotElementType()
                        .ToElements();

                    foreach (var f in fnds)
                    {
                        var type = GetTypeName(f, doc);
                        var level = GetLevelName(f, doc);
                        var vol = GetAnyVolume(f);
                        Acc(type, level, f.Id.IntegerValue, vol, ml_ft: null, addMl: false);
                    }
                }

                var groupRows = groups.Select(kv => new
                {
                    key = kv.Value.keyObj,
                    count = kv.Value.count,
                    volume_m3 = kv.Value.vol,
                    ml_m = kv.Value.ml
                }).ToList();

                return new { groups = groupRows, rows = req.includeIds ? rows : null };
            };
        }
    }
}

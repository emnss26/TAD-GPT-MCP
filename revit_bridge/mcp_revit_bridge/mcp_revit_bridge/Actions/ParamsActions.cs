using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection.Emit;
using System.Text;
using System.Threading.Tasks;
using static System.Net.Mime.MediaTypeNames;


namespace mcp_app.Actions
{
    internal class ParamsActions
    {
        // ===== Models =====

        private class BulkReq
        {
            // El server TS debe convertir CSV/XLSX a updates[] antes de llegar aquí.
            public UpdateItem[] updates { get; set; } = Array.Empty<UpdateItem>();

            // Campos opcionales que pueden venir en la petición original (no se usan aquí).
            public string tableCsv { get; set; }
            public string tableXlsxPath { get; set; }
            public string idColumn { get; set; }
            public string paramColumn { get; set; }
            public string valueColumn { get; set; }
        }

        private class GetReq
        {
            public int[] elementIds { get; set; } = Array.Empty<int>();
            public string[] paramNames { get; set; } = Array.Empty<string>();
            public bool includeValueString { get; set; } = true;
        }

        private class UpdateItem
        {
            public int elementId { get; set; }                 // Si viene <=0, se usará la selección actual (uno o varios)
            public string param { get; set; }                  // nombre visible (p.ej. "Comments") o fallback si no hay bip/guid
            public JToken value { get; set; }                  // string | number | boolean | null
        }

        private class SetReq
        {
            public UpdateItem[] updates { get; set; } = Array.Empty<UpdateItem>();
        }

        // === Tokens de parámetro multi-método ===
        private class ParamSetItem
        {
            public string param { get; set; }                  // visible
            public string bip { get; set; }                    // BuiltInParameter (texto) p.ej. "ALL_MODEL_INSTANCE_COMMENTS"
            public string guid { get; set; }                   // GUID de parámetro compartido
            public JToken value { get; set; }
        }

        // === Filtros WHERE ===
        private class WhereClause
        {
            public int[] elementIds { get; set; } = Array.Empty<int>();
            public bool? useSelection { get; set; }            // si true, usa selección actual
            public int[] typeIds { get; set; } = Array.Empty<int>();
            public string[] typeNames { get; set; } = Array.Empty<string>();
            public string[] familyNames { get; set; } = Array.Empty<string>();
            public string[] categories { get; set; } = Array.Empty<string>();   // preferible OST_...; intenta también por nombre visible
            public int[] categoryIds { get; set; } = Array.Empty<int>();
            public int[] levelIds { get; set; } = Array.Empty<int>();
            public string[] levelNames { get; set; } = Array.Empty<string>();
            public int? viewId { get; set; }                   // si se pasa, colecta elementos visibles en esa vista
            public string viewName { get; set; }
        }

        private class SetWhereReq
        {
            public WhereClause where { get; set; } = new WhereClause();
            public ParamSetItem[] set { get; set; } = Array.Empty<ParamSetItem>();
            public bool allowTypeParams { get; set; } = true;  // NUEVO: evitar tocar params de tipo si false
        }

        private static string GetTypeName(Element el, Document doc)
        {
            var tid = el.GetTypeId();
            if (tid == ElementId.InvalidElementId) return null;
            return (doc.GetElement(tid) as ElementType)?.Name;
        }

        private static string GetFamilyName(Element el, Document doc)
        {
            var tid = el.GetTypeId();
            if (tid == ElementId.InvalidElementId) return null;
            var et = doc.GetElement(tid) as ElementType;
            return et?.FamilyName;
        }

        private static string GetLevelName(Element el, Document doc)
        {
            ElementId levelId = TryGetLevelId(el);

            if (levelId == ElementId.InvalidElementId)
                return null;

            return (doc.GetElement(levelId) as Level)?.Name;
        }

        private static ElementId TryGetLevelId(Element el)
        {
            // 1) FamilyInstance.LevelId
            if (el is FamilyInstance fi && fi.LevelId != null && fi.LevelId != ElementId.InvalidElementId)
                return fi.LevelId;

            // 2) Propiedad directa LevelId (algunos elementos la tienen)
            if (el.LevelId != null && el.LevelId != ElementId.InvalidElementId)
                return el.LevelId;

            // 3) Parámetros comunes de nivel (en orden de probabilidad)
            var paramIds = new[]
            {
                BuiltInParameter.INSTANCE_REFERENCE_LEVEL_PARAM, // muy común
                BuiltInParameter.LEVEL_PARAM,                    // genérico
                BuiltInParameter.WALL_BASE_CONSTRAINT,           // muros
                BuiltInParameter.FAMILY_LEVEL_PARAM,             // familias basadas en cara
                BuiltInParameter.RBS_START_LEVEL_PARAM           // MEP
            };

            foreach (var bip in paramIds)
            {
                try
                {
                    var p = el.get_Parameter(bip);
                    if (p != null && p.StorageType == StorageType.ElementId)
                    {
                        var id = p.AsElementId();
                        if (id != null && id != ElementId.InvalidElementId)
                            return id;
                    }
                }
                catch { /* continuar */ }
            }

            return ElementId.InvalidElementId;
        }

        /* ------------------- GET ------------------- */
        public static Func<UIApplication, object> ParamsGet(JObject args)
        {
            var req = args.ToObject<GetReq>() ?? new GetReq();

            return (UIApplication app) =>
            {
                var doc = app.ActiveUIDocument?.Document ?? throw new Exception("No active document.");
                var rows = new List<object>();

                foreach (var eid in req.elementIds ?? Array.Empty<int>())
                {
                    var el = doc.GetElement(new ElementId(eid));
                    if (el == null)
                    {
                        foreach (var pname in req.paramNames ?? Array.Empty<string>())
                            rows.Add(new { elementId = eid, param = pname, error = "Element not found" });
                        continue;
                    }

                    foreach (var pname in req.paramNames ?? Array.Empty<string>())
                    {
                        var p = FindParam(el, pname, null, null, out Element targetEl, out string source);
                        if (p == null)
                        {
                            rows.Add(new { elementId = eid, param = pname, error = "Parameter not found" });
                            continue;
                        }

                        rows.Add(ParamRow(p, eid, pname, source, req.includeValueString));
                    }
                }

                return new { count = rows.Count, items = rows };
            };
        }

        /* ------------------- SET por lista (usa selección si elementId <= 0) ------------------- */
        public static Func<UIApplication, object> ParamsSet(JObject args)
        {
            var req = args.ToObject<SetReq>() ?? new SetReq();

            return (UIApplication app) =>
            {
                var uidoc = app.ActiveUIDocument ?? throw new Exception("No active document.");
                var doc = uidoc.Document;

                // Expande updates que no traen elementId usando la selección actual
                var selIds = uidoc.Selection.GetElementIds()?.Select(id => id.IntegerValue).ToList() ?? new List<int>();
                var expanded = new List<UpdateItem>();

                foreach (var u in req.updates ?? Array.Empty<UpdateItem>())
                {
                    if (u.elementId > 0)
                    {
                        expanded.Add(u);
                    }
                    else
                    {
                        if (selIds.Count == 0)
                            throw new Exception("Update sin elementId y no hay elementos seleccionados.");
                        foreach (var sid in selIds)
                            expanded.Add(new UpdateItem { elementId = sid, param = u.param, value = u.value });
                    }
                }

                var results = new List<object>();
                int ok = 0, failed = 0;

                using (var t = new Transaction(doc, "MCP: Params.Set"))
                {
                    t.Start();

                    foreach (var u in expanded)
                    {
                        try
                        {
                            var el = doc.GetElement(new ElementId(u.elementId));
                            if (el == null) { failed++; results.Add(new { u.elementId, u.param, ok = false, error = "Element not found" }); continue; }

                            var p = FindParam(el, u.param, null, null, out Element targetEl, out string source);
                            if (p == null) { failed++; results.Add(new { u.elementId, u.param, ok = false, error = "Parameter not found" }); continue; }
                            if (p.IsReadOnly) { failed++; results.Add(new { u.elementId, u.param, ok = false, error = "Parameter is read-only" }); continue; }

                            ApplyParamValue(p, u.value);

                            ok++;
                            results.Add(new { u.elementId, u.param, ok = true, target = source });
                        }
                        catch (Exception ex)
                        {
                            failed++;
                            results.Add(new { u.elementId, u.param, ok = false, error = ex.Message });
                        }
                    }

                    t.Commit();
                }

                return new { updated = ok, failed, results };
            };
        }

        /* ------------------- SET por filtros WHERE ------------------- */
        public static Func<UIApplication, object> ParamsSetWhere(JObject args)
        {
            var req = args.ToObject<SetWhereReq>() ?? new SetWhereReq();
            if (req.set == null || req.set.Length == 0)
                throw new Exception("params.set_where requiere 'set[]'.");

            return (UIApplication app) =>
            {
                var uidoc = app.ActiveUIDocument ?? throw new Exception("No active document.");
                var doc = uidoc.Document;

                var targets = ResolveTargets(doc, uidoc, req.where);
                if (targets.Count == 0) throw new Exception("No se resolvieron elementos objetivo.");

                var results = new List<object>();
                int ok = 0, failed = 0;

                using (var t = new Transaction(doc, "MCP: Params.SetWhere"))
                {
                    t.Start();

                    foreach (var el in targets)
                    {
                        foreach (var s in req.set)
                        {
                            var label = s.bip ?? s.guid ?? s.param;

                            try
                            {
                                var p = FindParam(el, s.param, s.bip, s.guid, out Element targetEl, out string source);
                                if (p == null) { failed++; results.Add(new { id = el.Id.IntegerValue, param = label, ok = false, error = "Parameter not found" }); continue; }

                                // NUEVO: opcionalmente evita tocar parámetros de TIPO
                                if (!req.allowTypeParams && string.Equals(source, "type", StringComparison.OrdinalIgnoreCase))
                                {
                                    failed++;
                                    results.Add(new { id = el.Id.IntegerValue, param = label, ok = false, error = "Type parameter skipped (allowTypeParams=false)" });
                                    continue;
                                }

                                if (p.IsReadOnly) { failed++; results.Add(new { id = el.Id.IntegerValue, param = label, ok = false, error = "Parameter is read-only" }); continue; }

                                ApplyParamValue(p, s.value);

                                ok++;
                                results.Add(new { id = el.Id.IntegerValue, param = label, ok = true, target = source });
                            }
                            catch (Exception ex)
                            {
                                failed++;
                                results.Add(new { id = el.Id.IntegerValue, param = label, ok = false, error = ex.Message });
                            }
                        }
                    }

                    t.Commit();
                }

                return new { targeted = targets.Count, updated = ok, failed, results };
            };
        }

        // ===== NUEVO: listar parámetros (inspección rápida) =====
        public static Func<UIApplication, object> ParamsList(JObject args)
        {
            var elementIds = args["elementIds"]?.ToObject<int[]>() ?? Array.Empty<int>();
            bool includeType = args["includeType"]?.ToObject<bool?>() ?? true;

            return (UIApplication app) =>
            {
                var doc = app.ActiveUIDocument?.Document ?? throw new Exception("No active document.");
                var items = new List<object>();

                foreach (var eid in elementIds)
                {
                    var el = doc.GetElement(new ElementId(eid));
                    if (el == null) { items.Add(new { elementId = eid, error = "Element not found" }); continue; }

                    // instancia
                    foreach (Parameter p in el.Parameters)
                        items.Add(DescribeParam(p, eid, "instance"));

                    // tipo
                    if (includeType)
                    {
                        var typ = doc.GetElement(el.GetTypeId()) as Element;
                        if (typ != null)
                            foreach (Parameter p in typ.Parameters)
                                items.Add(DescribeParam(p, eid, "type"));
                    }
                }

                return new { count = items.Count, items };
            };
        }

        private static object DescribeParam(Parameter p, int elementId, string scope)
        {
            string name = null;
            string storage = null;
            bool readOnly = false;
            string builtInName = null;
            string guidStr = null;

            try { name = p.Definition?.Name; } catch { }
            try { storage = p.StorageType.ToString(); } catch { }
            try { readOnly = p.IsReadOnly; } catch { }

            try
            {
                // GUID solo para shared params
                if (p.Definition is ExternalDefinition ext) guidStr = ext.GUID.ToString();
            }
            catch { }

            try
            {
                // BuiltInParameter si aplica (InternalDefinition)
                if (p.Definition is InternalDefinition idf)
                    builtInName = idf.BuiltInParameter.ToString();
            }
            catch { }

            return new
            {
                elementId,
                scope,               // "instance" | "type"
                name,
                storageType = storage,
                readOnly,
                guid = guidStr,
                builtIn = builtInName
            };
        }

        // ===== Helpers =====

        private static object ParamRow(Parameter p, int elementId, string paramToken, string source, bool includeValueString)
        {
            object raw = null;
            string storage = p.StorageType.ToString();
            string valueString = null;

            try
            {
                switch (p.StorageType)
                {
                    case StorageType.Integer: raw = p.AsInteger(); break;
                    case StorageType.Double: raw = p.AsDouble(); break;
                    case StorageType.String: raw = p.AsString(); break;
                    case StorageType.ElementId: raw = p.AsElementId()?.IntegerValue; break;
                }
            }
            catch { /* ignore */ }

            if (includeValueString)
            {
                try { valueString = p.AsValueString(); } catch { }
            }

            return new
            {
                elementId,
                param = paramToken,
                storageType = storage,
                value = raw,
                valueString,
                source
            };
        }

        private static void ApplyParamValue(Parameter p, JToken tok)
        {
            if (p == null) throw new ArgumentNullException(nameof(p));

            switch (p.StorageType)
            {
                case StorageType.String:
                    p.Set(tok?.Type == JTokenType.Null ? null : tok.ToString());
                    break;

                case StorageType.Integer:
                    if (tok?.Type == JTokenType.Boolean)
                        p.Set(tok.Value<bool>() ? 1 : 0);
                    else
                        p.Set(tok?.Type == JTokenType.Null ? 0 : tok.Value<int>());
                    break;

                case StorageType.Double:
                    if (tok?.Type == JTokenType.String)
                    {
                        var s = tok.Value<string>();
                        bool done = false;
                        try { p.SetValueString(s); done = true; } catch { /* fall back */ }
                        if (!done)
                        {
                            if (double.TryParse(s, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var d))
                                p.Set(d);
                            else
                                throw new Exception($"Cannot parse numeric value '{s}'.");
                        }
                    }
                    else
                    {
                        // UNIT-SMART: si llega número, intenta primero SetValueString (usa unidades del proyecto)
                        var n = tok?.Type == JTokenType.Null ? 0.0 : tok.Value<double>();
                        var sNum = n.ToString(System.Globalization.CultureInfo.InvariantCulture);
                        bool done = false;
                        try { p.SetValueString(sNum); done = true; } catch { /* fall back a unidades internas */ }
                        if (!done) p.Set(n);
                    }
                    break;

                case StorageType.ElementId:
                    if (tok?.Type == JTokenType.Null) { p.Set(ElementId.InvalidElementId); }
                    else p.Set(new ElementId(tok.Value<int>()));
                    break;

                default:
                    throw new Exception($"Unsupported storage type: {p.StorageType}");
            }
        }

        // --- Param finder extendido: bip / guid / visible ---
        private static Parameter FindParam(Element el, string token, string bipToken, string guidToken, out Element targetElement, out string source)
        {
            targetElement = el; source = "instance";

            // GUID (shared parameter)
            if (!string.IsNullOrWhiteSpace(guidToken) && Guid.TryParse(guidToken, out var g))
            {
                var p = FindParamByGuid(el, g);
                if (p != null) return p;

                var typ = el.Document.GetElement(el.GetTypeId());
                var pt = FindParamByGuid(typ as Element, g);
                if (pt != null) { targetElement = typ; source = "type"; return pt; }
            }

            // BuiltInParameter explícito
            if (!string.IsNullOrWhiteSpace(bipToken) && TryParseBuiltInParameter(bipToken, out var bipExplicit))
            {
                var pBuiltIn = el.get_Parameter(bipExplicit);
                if (pBuiltIn != null) return pBuiltIn;

                var typ = el.Document.GetElement(el.GetTypeId());
                var pT = (typ as Element)?.get_Parameter(bipExplicit);
                if (pT != null) { targetElement = typ; source = "type"; return pT; }
            }

            // Si token viene, prueba como BuiltInParameter primero y luego nombre visible
            if (!string.IsNullOrWhiteSpace(token))
            {
                if (TryParseBuiltInParameter(token, out var bip))
                {
                    var pBuiltIn = el.get_Parameter(bip);
                    if (pBuiltIn != null) return pBuiltIn;

                    var typ = el.Document.GetElement(el.GetTypeId());
                    var pT = (typ as Element)?.get_Parameter(bip);
                    if (pT != null) { targetElement = typ; source = "type"; return pT; }
                }

                // visible (instancia)
                var pByName = el?.Parameters
                    ?.Cast<Parameter>()
                    ?.FirstOrDefault(p => string.Equals(p.Definition?.Name, token, StringComparison.OrdinalIgnoreCase));
                if (pByName != null) return pByName;

                // visible (tipo)
                var typeId = el.GetTypeId();
                if (typeId != ElementId.InvalidElementId)
                {
                    var typEl = el.Document.GetElement(typeId);
                    var pTypeByName = typEl?.Parameters
                        ?.Cast<Parameter>()
                        ?.FirstOrDefault(p => string.Equals(p.Definition?.Name, token, StringComparison.OrdinalIgnoreCase));
                    if (pTypeByName != null) { targetElement = typEl; source = "type"; return pTypeByName; }
                }
            }

            return null;
        }

        private static Parameter FindParamByGuid(Element e, Guid g)
        {
            if (e == null) return null;
            foreach (Parameter p in e.Parameters)
            {
                try
                {
                    // sólo para shared params (ExternalDefinition)
                    var ext = p.Definition as ExternalDefinition;
                    if (ext != null && ext.GUID == g) return p;
                }
                catch { }
            }
            return null;
        }

        private static bool TryParseBuiltInParameter(string token, out BuiltInParameter bip)
        {
            // acepta EXACTO o sin may/min
            if (Enum.TryParse(token, true, out bip)) return true;

            // intenta mapear si viene sin prefijos raros, etc. (muy básico)
            foreach (var name in Enum.GetNames(typeof(BuiltInParameter)))
            {
                if (string.Equals(name, token, StringComparison.OrdinalIgnoreCase))
                {
                    bip = (BuiltInParameter)Enum.Parse(typeof(BuiltInParameter), name);
                    return true;
                }
            }
            bip = default;
            return false;
        }

        /* ----------- Target resolver para WHERE ----------- */
        private static List<Element> ResolveTargets(Document doc, UIDocument uidoc, WhereClause w)
        {
            var result = new HashSet<int>();

            // 1) elementIds explícitos
            foreach (var id in w?.elementIds ?? Array.Empty<int>()) result.Add(id);

            // 2) selección
            if (w?.useSelection == true)
            {
                var sel = uidoc.Selection.GetElementIds();
                foreach (var id in sel) result.Add(id.IntegerValue);
            }

            // 3) base collector (opcionalmente por vista)
            FilteredElementCollector baseCollector;
            if (w != null && (w.viewId.HasValue || !string.IsNullOrWhiteSpace(w.viewName)))
            {
                View v = null;
                if (w.viewId.HasValue && w.viewId.Value > 0)
                    v = doc.GetElement(new ElementId(w.viewId.Value)) as View;
                if (v == null && !string.IsNullOrWhiteSpace(w.viewName))
                    v = new FilteredElementCollector(doc).OfClass(typeof(View)).Cast<View>()
                         .FirstOrDefault(x => string.Equals(x.Name, w.viewName, StringComparison.OrdinalIgnoreCase));
                if (v == null) throw new Exception("View not found for filter.");
                baseCollector = new FilteredElementCollector(doc, v.Id);
            }
            else
            {
                baseCollector = new FilteredElementCollector(doc);
            }

            baseCollector = baseCollector.WhereElementIsNotElementType();

            // 4) APLICAR CATEGORÍAS EN EL COLLECTOR (idioma-independiente)
            var bicList = ResolveBuiltInCategories(doc, w);
            if (bicList.Count > 0)
            {
                // Multicategoría
                var catFilter = new ElementMulticategoryFilter(bicList);
                baseCollector = baseCollector.WherePasses(catFilter);
            }

            // Fallback: si el usuario pasó nombres visibles de categorías y no BIC,
            // nos quedamos con sus IDs para un segundo filtro manual.
            var allowedCatIdsByName = new HashSet<int>();
            if (w != null && (w.categories?.Length ?? 0) > 0)
            {
                foreach (var token in w.categories)
                {
                    if (TryParseBuiltInCategory(token, out _)) continue; // ya cubierto arriba
                    try
                    {
                        foreach (Category c in doc.Settings.Categories)
                            if (string.Equals(c.Name, token, StringComparison.OrdinalIgnoreCase))
                                allowedCatIdsByName.Add(c.Id.IntegerValue);
                    }
                    catch { /* ignore */ }
                }
            }

            // 5) otros filtros
            var wantTypeIds = new HashSet<int>((w?.typeIds ?? Array.Empty<int>()));
            var wantTypeNames = new HashSet<string>((w?.typeNames ?? Array.Empty<string>()), StringComparer.OrdinalIgnoreCase);
            var wantFamilyNames = new HashSet<string>((w?.familyNames ?? Array.Empty<string>()), StringComparer.OrdinalIgnoreCase);
            var wantLevelIds = new HashSet<int>((w?.levelIds ?? Array.Empty<int>()));
            var wantLevelNames = new HashSet<string>((w?.levelNames ?? Array.Empty<string>()), StringComparer.OrdinalIgnoreCase);

            foreach (var el in baseCollector)
            {
                int id = el.Id.IntegerValue;
                if (result.Contains(id)) continue;

                // Si hubo nombres visibles de categoría (no BIC), aplicamos filtro manual aquí
                if (allowedCatIdsByName.Count > 0)
                {
                    var catId = el.Category?.Id?.IntegerValue ?? -1;
                    if (catId < 0 || !allowedCatIdsByName.Contains(catId)) continue;
                }

                // type ids
                if (wantTypeIds.Count > 0)
                {
                    var tid = el.GetTypeId()?.IntegerValue ?? -1;
                    if (tid < 0 || !wantTypeIds.Contains(tid)) continue;
                }

                // type names
                if (wantTypeNames.Count > 0)
                {
                    var tname = GetTypeName(el, doc);
                    if (string.IsNullOrWhiteSpace(tname) || !wantTypeNames.Contains(tname)) continue;
                }

                // family names
                if (wantFamilyNames.Count > 0)
                {
                    var fname = GetFamilyName(el, doc);
                    if (string.IsNullOrWhiteSpace(fname) || !wantFamilyNames.Contains(fname)) continue;
                }

                // level ids / names
                if (wantLevelIds.Count > 0 || wantLevelNames.Count > 0)
                {
                    var lvlId = TryGetLevelId(el);
                    string lvlName = lvlId != ElementId.InvalidElementId
                        ? (doc.GetElement(lvlId) as Level)?.Name
                        : null;

                    if (wantLevelIds.Count > 0)
                    {
                        if (lvlId == ElementId.InvalidElementId || !wantLevelIds.Contains(lvlId.IntegerValue)) continue;
                    }
                    if (wantLevelNames.Count > 0)
                    {
                        if (string.IsNullOrWhiteSpace(lvlName) || !wantLevelNames.Contains(lvlName)) continue;
                    }
                }

                result.Add(id);
            }

            return result.Select(i => doc.GetElement(new ElementId(i))).Where(e => e != null).ToList();
        }

        private static HashSet<int> ResolveCategoryIds(Document doc, WhereClause w)
        {
            var set = new HashSet<int>();
            if (w == null) return set;

            foreach (var cid in w.categoryIds ?? Array.Empty<int>()) set.Add(cid);

            foreach (var token in w.categories ?? Array.Empty<string>())
            {
                // Preferible: OST_...
                if (TryParseBuiltInCategory(token, out var bic))
                {
                    var cat = Category.GetCategory(doc, bic);
                    if (cat != null) set.Add(cat.Id.IntegerValue);
                    continue;
                }

                // Fallback: nombre visible (puede ser idioma dependiente)
                try
                {
                    foreach (Category c in doc.Settings.Categories)
                        if (string.Equals(c.Name, token, StringComparison.OrdinalIgnoreCase))
                            set.Add(c.Id.IntegerValue);
                }
                catch { /* ignore */ }
            }

            return set;
        }

        private static bool TryParseBuiltInCategory(string token, out BuiltInCategory bic)
        {
            // acepta exacto, sin / con "OST_"
            if (Enum.TryParse(token, true, out bic)) return true;

            var norm = token?.Trim();
            if (string.IsNullOrEmpty(norm)) { bic = default; return false; }
            if (!norm.StartsWith("OST_", StringComparison.OrdinalIgnoreCase))
                norm = "OST_" + norm;

            if (Enum.TryParse(norm, true, out bic)) return true;

            // Búsqueda tolerantita
            foreach (var name in Enum.GetNames(typeof(BuiltInCategory)))
            {
                if (string.Equals(name, token, StringComparison.OrdinalIgnoreCase))
                {
                    bic = (BuiltInCategory)Enum.Parse(typeof(BuiltInCategory), name);
                    return true;
                }
            }
            bic = default;
            return false;
        }

        public static Func<UIApplication, object> ParamsBulkFromTable(JObject args)
        {
            var req = args.ToObject<BulkReq>() ?? new BulkReq();

            if (req.updates != null && req.updates.Length > 0)
            {
                var pass = new JObject
                {
                    ["updates"] = JArray.FromObject(req.updates)
                };
                return ParamsSet(pass);
            }

            return (UIApplication _) => new
            {
                ok = false,
                message = "params.bulk_from_table expects 'updates[]'. The TS server should parse CSV/XLSX and send updates[]."
            };
        }

        private static IList<BuiltInCategory> ResolveBuiltInCategories(Document doc, WhereClause w)
        {
            var bics = new HashSet<BuiltInCategory>();
            if (w == null) return bics.ToList();

            // Por tokens tipo "OST_Walls"
            foreach (var token in w.categories ?? Array.Empty<string>())
            {
                if (TryParseBuiltInCategory(token, out var bic))
                    bics.Add(bic);
            }

            // Por IDs numéricos (int) -> castear a BIC si aplica
            foreach (var cid in w.categoryIds ?? Array.Empty<int>())
            {
                try
                {
                    var bic = (BuiltInCategory)cid; // en built-ins el Id es negativo y mapea al enum
                    if (Enum.IsDefined(typeof(BuiltInCategory), bic))
                        bics.Add(bic);
                }
                catch { /* ignorar */ }
            }

            return bics.ToList();
        }
    }
}
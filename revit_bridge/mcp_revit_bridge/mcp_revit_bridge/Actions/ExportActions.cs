using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using mcp_app.Contracts;
using mcp_app.Core;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using static System.Net.Mime.MediaTypeNames;

namespace mcp_app.Actions
{
    internal class ExportActions
    {
        public static Func<UIApplication, object> ExportNwc(JObject args)
        {
            var req = args.ToObject<ExportNwcRequest>() ?? throw new Exception("Invalid args for export.nwc.");
            return (UIApplication app) =>
            {
                var uidoc = app.ActiveUIDocument ?? throw new Exception("No active document.");
                var doc = uidoc.Document;

                var baseOpts = new NavisworksExportOptions
                {
                    ExportElementIds = true,
                    ConvertElementProperties = req.convertElementProperties
                };

                var folder = req.folder ?? Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                Directory.CreateDirectory(folder);

                // --- Resolver vistas 3D destino ---
                var targets = new List<View3D>();

                if (req.viewId.HasValue && req.viewId.Value > 0)
                {
                    if (doc.GetElement(new ElementId(req.viewId.Value)) is View3D v && !v.IsTemplate) targets.Add(v);
                }

                if (req.viewIds != null && req.viewIds.Length > 0)
                {
                    foreach (var id in req.viewIds)
                        if (doc.GetElement(new ElementId(id)) is View3D v && !v.IsTemplate &&
                            !targets.Any(x => x.Id.IntegerValue == v.Id.IntegerValue))
                            targets.Add(v);
                }

                if (!string.IsNullOrWhiteSpace(req.viewName))
                {
                    var v = new FilteredElementCollector(doc).OfClass(typeof(View3D)).Cast<View3D>()
                        .FirstOrDefault(x => !x.IsTemplate && x.Name.Equals(req.viewName, StringComparison.OrdinalIgnoreCase));
                    if (v != null && !targets.Any(x => x.Id.IntegerValue == v.Id.IntegerValue)) targets.Add(v);
                }

                if (req.viewNames != null && req.viewNames.Length > 0)
                {
                    var all3d = new FilteredElementCollector(doc).OfClass(typeof(View3D)).Cast<View3D>()
                        .Where(x => !x.IsTemplate).ToList();
                    foreach (var name in req.viewNames)
                    {
                        var v = all3d.FirstOrDefault(x => x.Name.Equals(name, StringComparison.OrdinalIgnoreCase));
                        if (v != null && !targets.Any(x => x.Id.IntegerValue == v.Id.IntegerValue)) targets.Add(v);
                    }
                }

                if (targets.Count == 0)
                {
                    if (uidoc.ActiveView is View3D av3d && !av3d.IsTemplate) targets.Add(av3d);
                    else
                    {
                        var any3d = new FilteredElementCollector(doc).OfClass(typeof(View3D)).Cast<View3D>()
                            .FirstOrDefault(x => !x.IsTemplate)
                            ?? throw new Exception("No 3D view available for NWC export.");
                        targets.Add(any3d);
                    }
                }

                // Intentar overload (folder, name, ICollection<ElementId>, NavisworksExportOptions) si existe
                var exportWithViewsMethod = typeof(Document).GetMethod(
                    "Export",
                    new Type[] { typeof(string), typeof(string), typeof(ICollection<ElementId>), typeof(NavisworksExportOptions) }
                );

                // Propiedades opcionales dependientes de versión
                var nwType = typeof(NavisworksExportOptions);
                var pViewId = nwType.GetProperty("ViewId");
                var pExportScope = nwType.GetProperty("ExportScope");
                var pExportLinks = nwType.GetProperty("ExportLinks")
                                  ?? nwType.GetProperty("ExportLinkedFiles")
                                  ?? nwType.GetProperty("ExportRevitLinks");

                // FIX: exportLinks es bool (no nullable) ⇒ asignación directa
                bool exportLinksFlag = req.exportLinks; // <-- aquí el cambio

                var results = new List<object>();

                foreach (var v in targets)
                {
                    try
                    {
                        // Nombre del archivo por vista / prefijo
                        var baseName =
                            (targets.Count > 1 && !string.IsNullOrWhiteSpace(req.filename))
                            ? req.filename + "_" + SafeFileName(v.Name)
                            : (!string.IsNullOrWhiteSpace(req.filename) ? req.filename : SafeFileName(v.Name));

                        var path = Path.Combine(folder, baseName + ".nwc");
                        bool ok = false;

                        if (exportWithViewsMethod != null)
                        {
                            // Overload con lista de vistas (puede devolver bool o void según versión)
                            var ids = new List<ElementId> { v.Id };
                            var ret = exportWithViewsMethod.Invoke(doc, new object[] { folder, baseName, ids, baseOpts });
                            if (exportWithViewsMethod.ReturnType == typeof(bool))
                                ok = (bool)ret;
                            else
                                ok = File.Exists(path); // si es void, inferimos por el archivo
                        }
                        else
                        {
                            // Usar opciones por vista si existen; de lo contrario, fallback a cambiar la vista activa
                            var localOpts = new NavisworksExportOptions
                            {
                                ExportElementIds = baseOpts.ExportElementIds,
                                ConvertElementProperties = baseOpts.ConvertElementProperties
                            };

                            // Intentar activar exportación de links si la propiedad existe
                            if (pExportLinks != null && pExportLinks.CanWrite)
                            {
                                try { pExportLinks.SetValue(localOpts, exportLinksFlag); } catch { /* ignore */ }
                            }

                            bool setView = false;

                            if (pViewId != null && pViewId.CanWrite)
                            {
                                try { pViewId.SetValue(localOpts, v.Id); setView = true; } catch { /* ignore */ }
                            }

                            if (pExportScope != null && pExportScope.CanWrite)
                            {
                                try
                                {
                                    var enumType = pExportScope.PropertyType;
                                    var viewEnum = Enum.Parse(enumType, "View", true);
                                    pExportScope.SetValue(localOpts, viewEnum);
                                }
                                catch { /* ignore */ }
                            }

                            if (!setView)
                            {
                                // Último recurso: cambiar vista activa (respetando settings de UI)
                                uidoc.RequestViewChange(v);
                            }

                            // Este overload devuelve VOID en algunas versiones ⇒ inferimos ok por el archivo
                            doc.Export(folder, baseName, localOpts);
                            ok = File.Exists(path);
                        }

                        results.Add(new { ok, viewId = v.Id.IntegerValue, viewName = v.Name, path });
                    }
                    catch (Exception ex)
                    {
                        results.Add(new { ok = false, viewId = v.Id.IntegerValue, viewName = v.Name, error = ex.Message });
                    }
                }

                return new { count = results.Count, results };
            };
        }

        // Reemplazar caracteres inválidos en nombres de archivo (Windows)
        private static string SafeFileName(string name)
        {
            var invalid = Path.GetInvalidFileNameChars();
            var sb = new System.Text.StringBuilder(name.Length);
            for (int i = 0; i < name.Length; i++)
            {
                var c = name[i];
                sb.Append(invalid.Contains(c) ? '_' : c);
            }
            return sb.ToString();
        }

        public static Func<UIApplication, object> ExportDwg(JObject args)
        {
            var req = args.ToObject<ExportDwgRequest>() ?? throw new Exception("Invalid args for export.dwg.");
            return (UIApplication app) =>
            {
                var uidoc = app.ActiveUIDocument ?? throw new Exception("No active document.");
                var doc = uidoc.Document;
                var folder = req.folder ?? Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                Directory.CreateDirectory(folder);
                var name = req.filename ?? "export";

                var views = new List<ElementId>();
                if (req.viewIds != null && req.viewIds.Length > 0)
                    views.AddRange(req.viewIds.Select(id => new ElementId(id)));
                else
                    views.Add(uidoc.ActiveView.Id);

                var opts = !string.IsNullOrWhiteSpace(req.exportSetupName)
                    ? (DWGExportOptions.GetPredefinedOptions(doc, req.exportSetupName) ?? new DWGExportOptions())
                    : new DWGExportOptions();

                var ok = doc.Export(folder, name, views, opts);
                return new { ok, folder, name };
            };
        }

        public static Func<UIApplication, object> ExportPdf(JObject args)
        {
            var req = args.ToObject<ExportPdfRequest>() ?? throw new Exception("Invalid args for export.pdf.");
            return (UIApplication app) =>
            {
                var doc = app.ActiveUIDocument?.Document ?? throw new Exception("No active document.");
#if REVIT2022_OR_GREATER
                var folder = req.folder ?? Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                Directory.CreateDirectory(folder);
                var name = (req.filename ?? "set") + ".pdf";
                var path = Path.Combine(folder, name);

                var pdfOpts = new PDFExportOptions
                {
                    Combine = req.combine,
                };

                var ids = (req.viewOrSheetIds != null && req.viewOrSheetIds.Length > 0)
                    ? req.viewOrSheetIds.Select(i => new ElementId(i)).ToList()
                    : new List<ElementId> { doc.ActiveView.Id };

                var ok = doc.Export(folder, (req.filename ?? "set"), ids, pdfOpts);
                return new { ok, path };
#else
                throw new Exception("PDF export requires Revit 2022 or newer.");
#endif
            };
        }
    }
}

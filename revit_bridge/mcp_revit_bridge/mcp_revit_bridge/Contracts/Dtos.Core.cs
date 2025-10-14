using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Autodesk.Revit.DB;

namespace mcp_app.Contracts
{
    // Puntos 2D (m) y 3D (m)
    internal class Pt2 { public double x; public double y; }
    internal class Pt3 { public double x; public double y; public double z; }

    // ----- CORE: WALL/DOOR/WINDOW/LEVEL/GRID/FLOOR/CEILING -----

    internal class CreateWallRequest
    {
        public string level { get; set; }          // opcional
        public string wallType { get; set; }       // opcional
        public Pt2 start { get; set; }
        public Pt2 end { get; set; }
        public double height_m { get; set; }       // opcional (default 3.0)
        public bool structural { get; set; } = false;
    }

    internal class CreateWallResponse
    {
        public int elementId { get; set; }
        public object used { get; set; }
    }

    internal class LevelCreateRequest
    {
        public double elevation_m { get; set; }    // requerido
        public string name { get; set; }           // opcional (Revit auto-nombra si null)
    }

    internal class GridCreateRequest
    {
        public Pt2 start { get; set; }             // requerido (m)
        public Pt2 end { get; set; }               // requerido (m)
        public string name { get; set; }           // opcional (A, 1, etc.)
    }

    internal class FloorCreateRequest
    {
        public string level { get; set; }          // opcional → usa vista activa si plan
        public string floorType { get; set; }      // opcional
        public Pt2[] profile { get; set; }         // polígono en m, sentido horario
    }



    // Puerta/Ventana: requieren MURO host
    internal class DoorPlaceRequest
    {
        public int? hostWallId { get; set; }
        public string level { get; set; }
        public string familySymbol { get; set; }
        public Pt2 point { get; set; }
        public double offset_m { get; set; } = 0;

        public double? offsetAlong_m { get; set; }
        public double? alongNormalized { get; set; }

        public bool flipHand { get; set; } = false;
        public bool flipFacing { get; set; } = false;
    }

    internal class WindowPlaceRequest : DoorPlaceRequest { }

    // ----- GRAPHICS -----

    internal class ViewIdArg { public int? viewId { get; set; } }

    internal class ViewApplyTemplateRequest : ViewIdArg
    {
        public string templateName { get; set; } // por nombre
        public int? templateId { get; set; }     // o por Id
    }

    internal class ViewSetScaleRequest : ViewIdArg
    {
        public int scale { get; set; }           // p.ej. 50, 100
    }

    internal class ViewSetDetailRequest : ViewIdArg
    {
        public string detailLevel { get; set; }  // "Coarse","Medium","Fine"
    }

    internal class ViewSetDisciplineRequest : ViewIdArg
    {
        public string discipline { get; set; }   // "Architectural","Structural","Mechanical","Coordination"
    }

    internal class ViewSetPhaseRequest : ViewIdArg
    {
        public string phase { get; set; }        // nombre de Fase
    }

    internal class ViewsDuplicateRequest
    {
        public int[] viewIds { get; set; }       // vistas origen
        public string mode { get; set; } = "duplicate"; // "duplicate","with_detailing","as_dependent"
    }

    internal class SheetsCreateRequest
    {
        public string number { get; set; }         // p.ej. "A-101"
        public string name { get; set; }           // p.ej. "Planta Arquitectura L1"
        public string titleBlockType { get; set; } // opcional, p.ej. "A1 Metric"
    }

    internal class SheetsAddViewsRequest
    {

        public int sheetId { get; set; }
        public string sheetName { get; set; }

        public int[] viewIds { get; set; }
        public string[] viewNames { get; set; }
    }

    // ----- EXPORTS -----

    internal class ExportNwcRequest
    {
        public string folder { get; set; }
        public string filename { get; set; }
        public int[] viewIds { get; set; }
        public int? viewId { get; set; }
        public string viewName { get; set; }
        public string[] viewNames { get; set; }
        public bool convertElementProperties { get; set; } = true;
        public bool exportLinks { get; set; } = true;
    }

    internal class ExportDwgRequest
    {
        public string folder { get; set; }
        public string filename { get; set; }
        public int[] viewIds { get; set; }        // si null → activa
        public string exportSetupName { get; set; } // opcional: nombre de setup
    }

    internal class ExportPdfRequest
    {
        public string folder { get; set; }
        public string filename { get; set; }
        public int[] viewOrSheetIds { get; set; } // vistas u hojas
        public bool combine { get; set; } = true; // 1 PDF combinado
    }

    internal class RoomsCreateOnLevelsRequest
    {
        public string[] levelNames { get; set; } = Array.Empty<string>();
        public bool? placeOnlyEnclosed { get; set; } = true;
    }

    internal class FloorsFromRoomsRequest
    {
        public int[] roomIds { get; set; } = Array.Empty<int>();
        public string floorType { get; set; }     // opcional ("Family: Type" o solo Type)
        public double? baseOffset_m { get; set; } // opcional (por defecto 0)
    }



    internal class RoofFootprintCreateRequest
    {
        public string level { get; set; }                 // requerido
        public string roofType { get; set; }              // opcional
        public Pt2[] profile { get; set; }                // requerido
        public double? slope { get; set; }                // opcional (grados)
    }

    internal class CeilingCreateRequest
    {
        public string level { get; set; }                // opcional: usa vista activa si falta
        public string ceilingType { get; set; }          // opcional: primer CeilingType disponible
        public Pt2[] profile { get; set; }               // requerido: polígono cerrado XY en metros
        public double? baseOffset_m { get; set; }        // opcional: offset sobre el nivel (m). default: 0.0
    }

    internal class CeilingsFromRoomsRequest
    {
        public int[] roomIds { get; set; }               // requerido: rooms de origen
        public string ceilingType { get; set; }          // opcional: primer CeilingType si falta
        public double? baseOffset_m { get; set; }        // opcional: altura sobre el nivel de cada room (m). ej: 2.70
        public bool? useFinishBoundaries { get; set; }   // default true (contorno por acabados)
    }

    internal class RailingCreateRequest
    {
        public string level { get; set; }
        public string railingType { get; set; }
        public Pt2[] path { get; set; }                  // >= 2 puntos en metros
    }

    internal class StairCreateRequest
    {
        public string baseLevel { get; set; }
        public string topLevel { get; set; }
        public double? runWidth_m { get; set; }          // opcional
        public double? riserHeight_m { get; set; }       // opcional
        public double? treadDepth_m { get; set; }        // opcional
    }

    internal class RampCreateRequest
    {
        public string baseLevel { get; set; }
        public string topLevel { get; set; }
        public double? slopePercent { get; set; }        // opcional
        public Pt2[] path { get; set; }                  // opcional (si no damos niveles)
    }

    internal class FamilyLoadRequest
    {
        public string path { get; set; }                  // ruta absoluta al .rfa (requerido)
        public bool? overwriteExisting { get; set; }      // default true
    }

    internal class FamilyPlaceRequest
    {
        public string familySymbol { get; set; }          // "Familia: Tipo" o "Tipo" (requerido)
        public Pt2 point { get; set; }                    // XY en metros (requerido)
        public string level { get; set; }                 // opcional (usa vista activa si falta)
    }

}

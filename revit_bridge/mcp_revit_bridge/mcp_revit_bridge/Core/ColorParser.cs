using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace mcp_app.Core
{
    internal class ColorParser
    {
        public static Color FromAny(object any)
        {
            if (any is string s) return FromHex(s);
            var t = any.GetType();
            var r = t.GetProperty("r")?.GetValue(any) ?? t.GetProperty("R")?.GetValue(any);
            var g = t.GetProperty("g")?.GetValue(any) ?? t.GetProperty("G")?.GetValue(any);
            var b = t.GetProperty("b")?.GetValue(any) ?? t.GetProperty("B")?.GetValue(any);
            if (r != null && g != null && b != null)
                return new Color(Convert.ToByte(r), Convert.ToByte(g), Convert.ToByte(b));
            throw new ArgumentException("Unsupported color format. Use #RRGGBB or {r,g,b}.");
        }

        public static Color FromHex(string hex)
        {
            var m = Regex.Match(hex.Trim(), "^#?([0-9A-Fa-f]{6})$");
            if (!m.Success) throw new ArgumentException("Hex color must be #RRGGBB");
            byte R = Convert.ToByte(m.Groups[1].Value.Substring(0, 2), 16);
            byte G = Convert.ToByte(m.Groups[1].Value.Substring(2, 2), 16);
            byte B = Convert.ToByte(m.Groups[1].Value.Substring(4, 2), 16);
            return new Color(R, G, B);
        }
    }
}

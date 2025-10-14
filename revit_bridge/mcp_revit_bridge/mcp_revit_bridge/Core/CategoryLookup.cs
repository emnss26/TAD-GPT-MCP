using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace mcp_app.Core
{
    internal class CategoryLookup
    {
        public static Category Require(Document doc, string token)
        {
            BuiltInCategory bic;
            if (Enum.TryParse(token, out bic))
            {
                var catB = Category.GetCategory(doc, bic);
                if (catB != null) return catB;
            }

            foreach (Category c in doc.Settings.Categories)
            {
                if (string.Equals(c.Name, token, StringComparison.OrdinalIgnoreCase))
                    return c;
            }

            throw new InvalidOperationException("Category not found: " + token);
        }
    }
}

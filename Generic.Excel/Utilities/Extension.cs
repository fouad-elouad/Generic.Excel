using Generic.Excel.Export;
using System.Collections.Generic;
using System.Linq;

namespace Generic.Excel.Utilities
{
    internal static class Extension
    {

        public static IList<string> Names(this IEnumerable<ExcelProperty> excelProperties)
        {
            return excelProperties.Select(e => e.Name).ToList();
        }

        public static IList<string> DisplayNames(this IEnumerable<ExcelProperty> excelProperties)
        {
            return excelProperties.Select(e => e.DisplayName).ToList();
        }
    }
}

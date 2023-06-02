using System.Collections.Generic;
using System.Linq;

namespace Generic.Excel.Export
{
    public struct ExcelProperty
    {
        public string Name { get; private set; }
        public string DisplayName { get; private set; }

        public static ExcelProperty Create(string name, string displayName)
        {
            return new ExcelProperty { Name = name, DisplayName = displayName };
        }

        public static IEnumerable<ExcelProperty> Create(IEnumerable<string> names, IEnumerable<string> displayNames)
        {
            if (names == null || displayNames == null)
                throw new System.ArgumentNullException();
            if (names.Count() != displayNames.Count())
                throw new System.ArgumentOutOfRangeException("Length missmatch");

            List<ExcelProperty> list = new List<ExcelProperty>();
            foreach (var pair in names?.Zip(displayNames, (name, displayName) => new { Name = name, DisplayName = displayName }))
            {
                list.Add(Create(pair.Name, pair.DisplayName));
            }

            return list;
        }
    }
}

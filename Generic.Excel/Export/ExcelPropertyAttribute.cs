using System;

namespace Generic.Excel.Export
{
    [AttributeUsage(AttributeTargets.Property | AttributeTargets.Field, AllowMultiple = false)]
    public class ExcelPropertyAttribute : Attribute
    {
        public string DisplayName { get; protected set; }
        public bool Ignore { get; protected set; }

        public int Order { get; protected set; }

        public string NestedProperty { get; protected set; }

        public ExcelPropertyAttribute(string displayName, int order = 999, bool ignore = false, string nestedProperty = null)
        {
            DisplayName = displayName;
            Ignore = ignore;
            Order = order;
            NestedProperty = nestedProperty;
        }

        public ExcelPropertyAttribute(string displayName, int order, string nestedProperty)
        {
            DisplayName = displayName;
            Order = order;
            NestedProperty = nestedProperty;
        }

        public ExcelPropertyAttribute(bool ignore, int order = 999, string nestedProperty = null)
        {
            Ignore = ignore;
            Order = order;
            NestedProperty = nestedProperty;
        }

        public ExcelPropertyAttribute(int order)
        {
            Order = order;
        }

        public ExcelPropertyAttribute(string displayName, string nestedProperty)
        {
            DisplayName = displayName;
            NestedProperty = nestedProperty;
        }
    }
}

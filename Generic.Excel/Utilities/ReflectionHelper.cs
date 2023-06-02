using Generic.Excel.Export;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection;

namespace Generic.Excel.Utilities
{
    /// <summary>
    /// Reflection Helper
    /// </summary>
    public static class ReflectionHelper
    {

        public const string _dot = ".";

        /// <summary>
        /// Get Property Value by name
        /// </summary>
        /// <param name="obj"></param>
        /// <param name="propName"></param>
        /// <returns></returns>
        public static object GetPropertyValue(this object obj, string propName)
        {
            string[] nameParts = propName.Split(_dot.ToCharArray());
            if (nameParts.Length == 1)
            {
                return obj.GetType().GetProperty(propName).GetValue(obj, null);
            }

            foreach (string part in nameParts)
            {
                if (obj == null) { return null; }

                Type type = obj.GetType();
                PropertyInfo info = type.GetProperty(part);
                if (info == null) { return null; }

                obj = info.GetValue(obj, null);
            }
            return obj;
        }

        /// <summary>
        /// Gets the Display attribute name from PropertyInfo otherwise return the property Name.
        /// </summary>
        /// <param name="property">The property.</param>
        /// <param name="propertypropertyName">The property Name.</param>
        /// <returns>
        /// Display Attribute Name
        /// </returns>
        public static bool TryGetDisplayAttributeName(this PropertyInfo property, out string propertyName)
        {
            ExcelPropertyAttribute excelPropertyAttribute = property.GetCustomAttribute<ExcelPropertyAttribute>(false);
            if (excelPropertyAttribute != null)
            {
                if (excelPropertyAttribute.Ignore)
                {
                    propertyName = null;
                    return false;
                }
                else
                {
                    if (excelPropertyAttribute.DisplayName != null)
                    {
                        propertyName = excelPropertyAttribute.DisplayName;
                        return true;
                    }
                }
            }
            propertyName = property.Name;
            return true;
        }

        /// <summary>
        /// Gets the Display attribute name from PropertyInfo otherwise return the property Name.
        /// </summary>
        /// <param name="property">The property.</param>
        /// <returns>
        /// Display Attribute Name
        /// </returns>
        public static string GetDisplayAttributeName(this PropertyInfo property)
        {
            ExcelPropertyAttribute excelPropertyAttribute = property.GetCustomAttribute<ExcelPropertyAttribute>(false);
            if (excelPropertyAttribute != null && excelPropertyAttribute.DisplayName != null)
            {
                return excelPropertyAttribute.DisplayName;
            }

            return property.Name;
        }

        public static bool IsIgnored(this PropertyInfo property)
        {
            ExcelPropertyAttribute excelPropertyAttribute = property.GetCustomAttribute<ExcelPropertyAttribute>(false);
            return excelPropertyAttribute != null && excelPropertyAttribute.Ignore;
        }

        /// <summary>
        /// Gets the Display attribute name from PropertyInfo otherwise return the property Name.
        /// </summary>
        /// <param name="property">The property.</param>
        /// <returns>
        /// Display Attribute Name
        /// </returns>
        public static bool TryGetNestedPropertyName(this PropertyInfo property, out string nestedPropertyName)
        {
            ExcelPropertyAttribute excelPropertyAttribute = property.GetCustomAttribute<ExcelPropertyAttribute>(false);
            if (excelPropertyAttribute == null || excelPropertyAttribute.Ignore || excelPropertyAttribute.NestedProperty == null)
            {
                nestedPropertyName = null;
                return false;
            }
            else
            {
                nestedPropertyName = string.Concat(property.Name, _dot, excelPropertyAttribute.NestedProperty);
                return true;
            }
        }

        /// <summary>
        /// Order
        /// </summary>
        /// <param name="properties">The properties.</param>
        /// <returns>
        /// Ordered array
        /// </returns>
        public static PropertyInfo[] Order(this PropertyInfo[] properties)
        {
            return properties.OrderBy(p => p.GetCustomAttribute<ExcelPropertyAttribute>(false) == null).ThenBy(p => p.GetCustomAttribute<ExcelPropertyAttribute>(false)?.Order).ToArray();
        }

        /// <summary>
        /// Gets the Display attribute name from PropertyInfo otherwise return the property Name.
        /// </summary>
        /// <param name="property">The property.</param>
        /// <returns>
        /// Display Attribute Name
        /// </returns>
        public static IList<ExcelProperty> GetExcelProperties(this Type type, IList<string> propertiesToExport = null, IList<string> propertiesDisplayName = null)
        {
            PropertyInfo[] headerProperties = type.GetProperties();
            bool ispropertiesToExport = propertiesToExport != null && propertiesToExport.Any();
            bool isPropertiesDisplayName = ispropertiesToExport && propertiesDisplayName != null && propertiesDisplayName.Any();
            if (propertiesToExport != null && propertiesToExport.Any())
            {
                var propertiesToExportSplited = propertiesToExport.Select(p => p.Split(_dot.ToCharArray()).FirstOrDefault()).ToList();
                headerProperties = headerProperties.Where(h => propertiesToExportSplited.Contains(h.Name))
                    .OrderBy(p => propertiesToExportSplited.IndexOf(p.Name)).ToArray();
                if (isPropertiesDisplayName && propertiesDisplayName.Count != propertiesToExport.Count)
                    throw new IndexOutOfRangeException("length mismatch");
            }
            else
                headerProperties = headerProperties.Order();

            var excelProperties = new List<ExcelProperty>();
            int index = 0;
            foreach (var property in headerProperties)
            {
                if (property.PropertyType.IsClass && property.PropertyType != typeof(string))
                {
                    if (property.TryGetNestedPropertyName(out string nestedProperty) || ispropertiesToExport)
                    {
                        string nestedPropertyName = ispropertiesToExport ? propertiesToExport[index] : nestedProperty;
                        string displayName = isPropertiesDisplayName ? propertiesDisplayName[index] : property.GetDisplayAttributeName();
                        excelProperties.Add(ExcelProperty.Create(nestedPropertyName, displayName));
                    }
                }

                else
                {
                    if (ispropertiesToExport && property.IsIgnored() && propertiesToExport.Contains(property.Name))
                    {
                        string displayName = isPropertiesDisplayName ? propertiesDisplayName[index] : property.GetDisplayAttributeName();
                        excelProperties.Add(ExcelProperty.Create(propertiesToExport[index], displayName));
                    }
                    else if (property.TryGetDisplayAttributeName(out string displayName))
                    {
                        displayName = isPropertiesDisplayName ? propertiesDisplayName[index] : displayName;
                        excelProperties.Add(ExcelProperty.Create(property.Name, displayName));
                    }
                }
                index++;
            }

            return excelProperties;
        }
    }
}
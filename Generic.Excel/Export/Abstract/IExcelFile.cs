using System;
using System.Collections.Generic;

namespace Generic.Excel.Export.Abstract
{

    public interface IExcelFile : IDisposable
    {

        /// <summary>
        /// Get nbr of sheets
        /// </summary>
        /// <returns></returns>
        int WorkSheetsCount();

        /// <summary>
        /// Create worksheet and add data
        /// headers are placed in first line
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="list">List to Export</param>
        /// <param name="workSheetName"></param>
        /// <param name="propertiesToExport">if is null all object properties are exported</param>
        /// <param name="propertiesDisplayName">if is null model display name is used otherwise model property name</param>
        /// <param name="horizontal">transpose result</param>
        void AddSheetList<T>(IEnumerable<T> list, string workSheetName,
                        IList<string> propertiesToExport = null, IList<string> propertiesDisplayName = null, bool horizontal = false) where T : class;

        /// <summary>
        /// Create worksheet and add data
        /// headers are placed in first line
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="list">List to Export</param>
        /// <param name="workSheetName"></param>
        /// <param name="propertiesNamesDictionary">if is null all object properties are exported</param>
        /// <param name="horizontal">transpose result</param>
        void AddSheetList<T>(IEnumerable<T> list, string workSheetName,
                            IDictionary<string, string> propertiesNamesDictionary, bool horizontal = false) where T : class;

        /// <summary>
        /// Create worksheet and add one column data
        /// used for primitive types (string, int ....)
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="list">List to Export</param>
        /// <param name="workSheetName"></param>
        /// <param name="header">used in column header, if is null only data is exported without header</param>
        /// <param name="horizontal">transpose result</param>
        void AddSheetList<T>(IEnumerable<T> list, string workSheetName, string header = null, bool horizontal = false) where T : struct;

        /// <summary>
        /// Create worksheet and add one column data
        /// </summary>
        /// <param name="list">List to Export</param>
        /// <param name="workSheetName"></param>
        /// <param name="header">used in column header, if is null only data is exported without header</param>
        /// <param name="horizontal">transpose result</param>
        void AddSheet(IEnumerable<string> list, string workSheetName, string header = null, bool horizontal = false);

        /// <summary>
        /// Create worksheet and add data
        /// headers are placed in first column
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="obj"></param>
        /// <param name="workSheetName"></param>
        /// <param name="propertiesToExport">if is null all object properties are exported</param>
        /// <param name="propertiesDisplayName">if is null model display name is used otherwise model property name</param>
        /// <param name="horizontal">transpose result</param>
        void AddSheetObj<T>(T obj, string workSheetName, IList<string> propertiesToExport = null,
                                IList<string> propertiesDisplayName = null, bool horizontal = false) where T : class;

        /// <summary>
        /// Add Blank WorkSheet
        /// </summary>
        /// <param name="workSheetName"></param>
        void AddBlankWorkSheet(string workSheetName);

        /// <summary>
        /// Create worksheet and add data
        /// used for dictionary data
        /// </summary>
        /// <param name="dictionary"></param>
        /// <param name="workSheetName"></param>
        /// <param name="keyHeader"></param>
        /// <param name="valueHeader"></param>
        /// <param name="horizontal">transpose result</param>
        void AddSheet<T1, T2>(IDictionary<T1, ICollection<T2>> dictionary, string workSheetName, string keyHeader = null,
                        string valueHeader = null, bool horizontal = false);

        /// <summary>
        /// Create worksheet and add data
        /// used for dictionary data
        /// </summary>
        /// <param name="dictionary"></param>
        /// <param name="workSheetName"></param>
        /// <param name="keyHeader"></param>
        /// <param name="valueHeader"></param>
        /// <param name="horizontal">transpose result</param>
        void AddSheet<T1, T2>(IDictionary<T1, T2> dictionary, string workSheetName, string keyHeader = null,
                        string valueHeader = null, bool horizontal = false);

    }
}
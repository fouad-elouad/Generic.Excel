using ClosedXML.Excel;
using Generic.Excel.Export.Abstract;
using Generic.Excel.Utilities;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Generic.Excel.Export
{
    public class ExcelFile : IExcelFile
    {
        protected XLWorkbook xLWorkbook { get; private set; }
        private bool _isDisposed = false;
        protected ExcelFile() { }


        /// <summary>
        /// Get nbr of sheets
        /// </summary>
        /// <returns></returns>
        public int WorkSheetsCount()
        {
            try
            {
                return xLWorkbook.Worksheets.Count();
            }
            catch (NullReferenceException)
            {
                return 0;
            }
        }


        /// <summary>
        /// Creates the specified file namesuffix
        /// must be used in using {} statement
        /// </summary>
        /// <returns> ExcelFile </returns>
        public static ExcelFile Create()
        {
            return new ExcelFile { xLWorkbook = new XLWorkbook() };
        }

        /// <summary>
        /// Saves the current ExcelFile to a file.
        /// </summary>
        /// <returns> ExcelFile </returns>
        public void SaveAs(string file)
        {
            xLWorkbook.SaveAs(file);
        }

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
        public void AddSheetList<T>(IEnumerable<T> list, string workSheetName,
                        IList<string> propertiesToExport = null, IList<string> propertiesDisplayName = null, bool horizontal = false) where T : class
        {
            InternalAddSheet(list, workSheetName, propertiesToExport, propertiesDisplayName, horizontal);
        }

        /// <summary>
        /// Create worksheet and add data
        /// headers are placed in first line
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="list">List to Export</param>
        /// <param name="workSheetName"></param>
        /// <param name="propertiesNamesDictionary">if is null all object properties are exported</param>
        /// <param name="horizontal">transpose result</param>
        public void AddSheetList<T>(IEnumerable<T> list, string workSheetName,
                            IDictionary<string, string> propertiesNamesDictionary, bool horizontal = false) where T : class
        {
            InternalAddSheet(list, workSheetName, propertiesNamesDictionary?.Keys?.ToList(), propertiesNamesDictionary?.Values?.ToList(), horizontal);
        }

        /// <summary>
        /// Create worksheet and add one column data
        /// used for primitive types (string, int ....)
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="list">List to Export</param>
        /// <param name="workSheetName"></param>
        /// <param name="header">used in column header, if is null only data is exported without header</param>
        /// <param name="horizontal">transpose result</param>
        public void AddSheetList<T>(IEnumerable<T> list, string workSheetName, string header = null, bool horizontal = false) where T : struct
        {
            InternalAddSheet<T>(list, workSheetName, header, horizontal);
        }

        /// <summary>
        /// Create worksheet and add one column data
        /// </summary>
        /// <param name="list">List to Export</param>
        /// <param name="workSheetName"></param>
        /// <param name="header">used in column header, if is null only data is exported without header</param>
        /// <param name="horizontal">transpose result</param>
        public void AddSheet(IEnumerable<string> list, string workSheetName, string header = null, bool horizontal = false)
        {
            InternalAddSheet<string>(list, workSheetName, header, horizontal);
        }

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
        public void AddSheetObj<T>(T obj, string workSheetName, IList<string> propertiesToExport = null,
                                IList<string> propertiesDisplayName = null, bool horizontal = false) where T : class
        {
            InternalAddSheet(new List<T> { obj }, workSheetName, propertiesToExport, propertiesDisplayName, horizontal);
        }

        /// <summary>
        /// Add Blank WorkSheet
        /// </summary>
        /// <param name="workSheetName"></param>
        public void AddBlankWorkSheet(string workSheetName)
        {
            xLWorkbook.Worksheets.Add(workSheetName);
        }

        /// <summary>
        /// Create worksheet and add data
        /// used for dictionary data
        /// </summary>
        /// <param name="dictionary"></param>
        /// <param name="workSheetName"></param>
        /// <param name="keyHeader"></param>
        /// <param name="valueHeader"></param>
        /// <param name="horizontal">transpose result</param>
        public void AddSheet<T1, T2>(IDictionary<T1, ICollection<T2>> dictionary, string workSheetName, string keyHeader = null,
                        string valueHeader = null, bool horizontal = false)
        {
            var workSheet = xLWorkbook.Worksheets.Add(workSheetName);

            int startRow = 1;
            // create header
            if (keyHeader != null && valueHeader != null)
            {
                workSheet.Cell(startRow, 1).SetValue(keyHeader);
                workSheet.Cell(startRow, 2).SetValue(valueHeader);
                startRow++;
            }

            int rowIndex = 0;
            foreach (var obj in dictionary)
            {
                foreach (var value in obj.Value)
                {
                    workSheet.Cell(rowIndex + startRow, 1).SetValue(obj.Key);
                    workSheet.Cell(rowIndex + startRow, 2).SetValue(value);
                    rowIndex++;
                }
            }

            if (horizontal)
                Transpose(workSheet);
            workSheet.Columns().AdjustToContents();
        }

        /// <summary>
        /// Create worksheet and add data
        /// used for dictionary data
        /// </summary>
        /// <param name="dictionary"></param>
        /// <param name="workSheetName"></param>
        /// <param name="keyHeader"></param>
        /// <param name="valueHeader"></param>
        /// <param name="horizontal">transpose result</param>
        public void AddSheet<T1, T2>(IDictionary<T1, T2> dictionary, string workSheetName, string keyHeader = null,
                        string valueHeader = null, bool horizontal = false)
        {
            var workSheet = xLWorkbook.Worksheets.Add(workSheetName);

            int startRow = 1;
            // create header
            if (keyHeader != null && valueHeader != null)
            {
                workSheet.Cell(startRow, 1).SetValue(keyHeader);
                workSheet.Cell(startRow, 2).SetValue(valueHeader);
                startRow++;
            }

            int rowIndex = 0;
            foreach (var obj in dictionary)
            {
                workSheet.Cell(rowIndex + startRow, 1).SetValue(obj.Key);
                workSheet.Cell(rowIndex + startRow, 2).SetValue(obj.Value);
                rowIndex++;
            }

            if (horizontal)
                Transpose(workSheet);
            workSheet.Columns().AdjustToContents();
        }

        protected virtual IXLWorksheet InternalAddSheet<T>(IEnumerable<T> list, string workSheetName, IList<string> propertiesToExport = null,
                    IList<string> propertiesDisplayName = null, bool horizontal = false) where T : class
        {
            var excelProperties = typeof(T).GetExcelProperties(propertiesToExport, propertiesDisplayName);

            var l_propertiesToExport = excelProperties.Names();
            var l_propertiesDisplayName = excelProperties.DisplayNames();

            var workSheet = xLWorkbook.Worksheets.Add(workSheetName);

            // create header
            for (int i = 0; i < l_propertiesDisplayName.Count; i++)
            {
                workSheet.Cell(1, i + 1).SetValue(l_propertiesDisplayName[i]);
            }

            int rowIndex = 0;
            foreach (var obj in list)
            {
                for (int cIndex = 0; cIndex < l_propertiesToExport.Count(); cIndex++)
                {
                    workSheet.Cell(rowIndex + 2, cIndex + 1).SetValue(obj.GetPropertyValue(l_propertiesToExport[cIndex]));
                }
                rowIndex++;
            }

            if (horizontal)
                Transpose(workSheet);
            workSheet.Columns().AdjustToContents();
            return workSheet;
        }

        protected IXLWorksheet InternalAddSheet<T>(IEnumerable<T> list, string workSheetName, string header = null, bool horizontal = false)
        {
            var workSheet = xLWorkbook.Worksheets.Add(workSheetName);
            int startIndex = 1;
            // create header
            if (header != null)
            {
                workSheet.Cell(startIndex, 1).SetValue(header);
                startIndex++;
            }

            foreach (var item in list)
            {
                workSheet.Cell(startIndex, 1).SetValue(item);
                startIndex++;
            }
            if (horizontal)
                Transpose(workSheet);
            workSheet.Columns().AdjustToContents();
            return workSheet;
        }

        protected virtual IXLWorksheet Transpose(IXLWorksheet worksheet)
        {
            worksheet.RangeUsed().Transpose(XLTransposeOptions.MoveCells);
            return worksheet;
        }

        /// <summary>
        /// Performs application-defined tasks associated with freeing, releasing, or resetting unmanaged resources.
        /// Flushes and saves the content, closes the document, and releases all resources.
        /// </summary>
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (_isDisposed)
                return;

            if (disposing)
            {
                if (xLWorkbook != null)
                    xLWorkbook.Dispose();
            }
            _isDisposed = true;
        }
    }
}
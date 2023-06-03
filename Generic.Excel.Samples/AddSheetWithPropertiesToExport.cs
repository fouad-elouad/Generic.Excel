using Generic.Excel.Export;
using Generic.Excel.Samples.Common;
using Generic.Excel.Samples.Utilities;
using System.Collections.Generic;

namespace Generic.Excel.Samples
{
    public class AddSheetWithPropertiesToExport : SampleBase
    {
        public AddSheetWithPropertiesToExport(DataFactory dataFactory) : base(dataFactory)
        {
            _DataFactory = dataFactory;
        }

        public AddSheetWithPropertiesToExport() : base()
        {
        }

        public override void Run()
        {
            IEnumerable<Employee> employees = _DataFactory.Generate(10);

            using (ExcelFile excelFile = ExcelFile.Create())
            {
                var propertiesToExport = new List<string> { "EmployeeID", "IsSupervisor", "Department.FullName", "Salary" };
                excelFile.AddSheetList<Employee>(employees, nameof(employees), propertiesToExport);

                string filePath = FilePath(nameof(AddSheetWithPropertiesToExport) + _ExcelExtension);
                excelFile.SaveAs(filePath);
            }
        }
    }
}

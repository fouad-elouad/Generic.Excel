using Generic.Excel.Export;
using Generic.Excel.Samples.Common;
using Generic.Excel.Samples.Utilities;
using System.Collections.Generic;

namespace Generic.Excel.Samples
{
    public class AddHorizontalSheetWithPropertiesDisplayName : SampleBase
    {
        public AddHorizontalSheetWithPropertiesDisplayName(DataFactory dataFactory) : base(dataFactory)
        {
            _DataFactory = dataFactory;
        }

        public AddHorizontalSheetWithPropertiesDisplayName() : base()
        {
        }

        public override void Run()
        {
            IEnumerable<Employee> employees = _DataFactory.Generate(10);

            using (ExcelFile excelFile = ExcelFile.Create())
            {
                var propertiesToExport = new List<string> { "EmployeeID", "IsSupervisor", "Department.FullName", "Salary" };
                var propertiesDisplayName = new List<string> { "display name 1", "display name 2", "display name 3", "display name 4" };
                excelFile.AddSheetList<Employee>(employees, nameof(employees), propertiesToExport, propertiesDisplayName, true);

                string filePath = FilePath(nameof(AddHorizontalSheetWithPropertiesDisplayName) + _ExcelExtension);
                excelFile.SaveAs(filePath);
            }
        }
    }
}

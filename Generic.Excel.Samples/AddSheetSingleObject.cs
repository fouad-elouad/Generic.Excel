using Generic.Excel.Export;
using Generic.Excel.Samples.Common;
using Generic.Excel.Samples.Utilities;
using System.Collections.Generic;

namespace Generic.Excel.Samples
{
    public class AddSheetSingleObject : SampleBase
    {
        public AddSheetSingleObject(DataFactory dataFactory) : base(dataFactory)
        {
            _DataFactory = dataFactory;
        }

        public AddSheetSingleObject() : base()
        {
        }

        public override void Run()
        {

            Employee employee = _DataFactory.Generate(false);

            using (ExcelFile excelFile = ExcelFile.Create())
            {
                excelFile.AddSheetObj<Employee>(employee, "Default");
                var propertiesToExport = new List<string> { "EmployeeID", "IsSupervisor", "Department.FullName", "Salary" };
                excelFile.AddSheetObj<Employee>(employee, "propertiesToExport", propertiesToExport);
                var propertiesDisplayName = new List<string> { "display name 1", "display name 2", "display name 3", "display name 4" };
                excelFile.AddSheetObj<Employee>(employee, "propertiesDisplayName", propertiesToExport, propertiesDisplayName);
                excelFile.AddSheetObj<Employee>(employee, "Horizontal", horizontal: true);

                string filePath = FilePath(nameof(AddSheetSingleObject) + _ExcelExtension);
                excelFile.SaveAs(filePath);
            }
        }
    }
}

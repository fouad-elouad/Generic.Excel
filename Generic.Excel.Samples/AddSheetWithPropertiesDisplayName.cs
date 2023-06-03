using Generic.Excel.Export;
using Generic.Excel.Samples.Common;
using Generic.Excel.Samples.Utilities;

namespace Generic.Excel.Samples
{
    public class AddSheetWithPropertiesDisplayName : SampleBase
    {
        public AddSheetWithPropertiesDisplayName(DataFactory dataFactory) : base(dataFactory)
        {
            _DataFactory = dataFactory;
        }

        public AddSheetWithPropertiesDisplayName() : base()
        {
        }

        public override void Run()
        {
            IEnumerable<Employee> employees = _DataFactory.Generate(10);

            using (ExcelFile excelFile = ExcelFile.Create())
            {
                var propertiesToExport = new List<string> { "EmployeeID", "IsSupervisor", "Department.FullName", "Salary" };
                var propertiesDisplayName = new List<string> { "display name 1", "display name 2", "display name 3", "display name 4" };
                excelFile.AddSheetList<Employee>(employees, nameof(employees), propertiesToExport, propertiesDisplayName);

                IDictionary<string, string> propertiesDictionary =
                    propertiesToExport.Zip(propertiesDisplayName, (k, v) => new { k, v }).ToDictionary(x => x.k, x => x.v);
                excelFile.AddSheetList<Employee>(employees, "Dictionary properties", propertiesDictionary);

                string filePath = FilePath(nameof(AddSheetWithPropertiesDisplayName) + _ExcelExtension);
                excelFile.SaveAs(filePath);
            }
        }
    }
}

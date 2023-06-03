using Generic.Excel.Export;
using Generic.Excel.Samples.Common;
using Generic.Excel.Samples.Utilities;

namespace Generic.Excel.Samples
{
    public class AddHorizontalSheetWithPropertiesToExport : SampleBase
    {
        public AddHorizontalSheetWithPropertiesToExport(DataFactory dataFactory) : base(dataFactory)
        {
            _DataFactory = dataFactory;
        }

        public AddHorizontalSheetWithPropertiesToExport() : base()
        {
        }

        public override void Run()
        {
            IEnumerable<Employee> employees = _DataFactory.Generate(10);

            using (ExcelFile excelFile = ExcelFile.Create())
            {
                var propertiesToExport = new List<string> { "EmployeeID", "IsSupervisor", "Department.FullName", "Salary" };
                excelFile.AddSheetList<Employee>(employees, nameof(employees), propertiesToExport, horizontal: true);

                string filePath = FilePath(nameof(AddHorizontalSheetWithPropertiesToExport) + _ExcelExtension);
                excelFile.SaveAs(filePath);
            }
        }
    }
}

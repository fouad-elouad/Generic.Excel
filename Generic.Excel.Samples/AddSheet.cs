using Generic.Excel.Export;
using Generic.Excel.Samples.Common;
using Generic.Excel.Samples.Utilities;

namespace Generic.Excel.Samples
{
    public class AddSheet : SampleBase
    {
        public AddSheet(DataFactory dataFactory) : base(dataFactory)
        {
            _DataFactory = dataFactory;
        }

        public AddSheet() : base()
        {
        }

        public override void Run()
        {
            IEnumerable<Employee> employees = _DataFactory.Generate(10);

            using (ExcelFile excelFile = ExcelFile.Create())
            {
                excelFile.AddSheetList<Employee>(employees, nameof(employees));

                string filePath = FilePath(nameof(AddSheet) + _ExcelExtension);
                excelFile.SaveAs(filePath);
            }
        }
    }
}

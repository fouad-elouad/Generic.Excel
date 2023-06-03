using Generic.Excel.Export;
using Generic.Excel.Samples.Common;
using Generic.Excel.Samples.Utilities;

namespace Generic.Excel.Samples
{
    public class WorkSheetsCount : SampleBase
    {
        public WorkSheetsCount(DataFactory dataFactory) : base(dataFactory)
        {
            _DataFactory = dataFactory;
        }

        public WorkSheetsCount() : base()
        {
        }

        public override void Run()
        {
            Employee employee = _DataFactory.Generate(false);
            using (ExcelFile excelFile = ExcelFile.Create())
            {
                excelFile.AddSheetObj<Employee>(employee, nameof(employee));
                var oneWorksheet = excelFile.WorkSheetsCount();

                excelFile.AddSheetObj<Employee>(employee, nameof(employee) + 2);
                excelFile.AddBlankWorkSheet(nameof(employee) + 3);
                var threeWorksheets = excelFile.WorkSheetsCount();
            }
        }
    }
}

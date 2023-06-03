using Generic.Excel.Export;
using Generic.Excel.Samples.Common;
using Generic.Excel.Samples.Utilities;
using System.Collections.Generic;

namespace Generic.Excel.Samples
{
    public class AddHorizontalSheet : SampleBase
    {
        public AddHorizontalSheet(DataFactory dataFactory) : base(dataFactory)
        {
            _DataFactory = dataFactory;
        }

        public AddHorizontalSheet() : base()
        {
        }

        public override void Run()
        {
            IEnumerable<Employee> employees = _DataFactory.Generate(10);

            using (ExcelFile excelFile = ExcelFile.Create())
            {
                excelFile.AddSheetList<Employee>(employees, nameof(employees), horizontal: true);

                string filePath = FilePath(nameof(AddHorizontalSheet) + _ExcelExtension);
                excelFile.SaveAs(filePath);
            }
        }
    }
}

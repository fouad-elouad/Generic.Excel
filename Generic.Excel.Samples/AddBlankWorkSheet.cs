using Generic.Excel.Export;
using Generic.Excel.Samples.Utilities;

namespace Generic.Excel.Samples
{
    public class AddBlankWorkSheet : SampleBase
    {
        public AddBlankWorkSheet(DataFactory dataFactory) : base(dataFactory)
        {
            _DataFactory = dataFactory;
        }

        public AddBlankWorkSheet() : base()
        {
        }

        public override void Run()
        {
            using (ExcelFile excelFile = ExcelFile.Create())
            {
                excelFile.AddBlankWorkSheet("BlankSheet");

                string filePath = FilePath(nameof(AddBlankWorkSheet) + _ExcelExtension);
                excelFile.SaveAs(filePath);
            }
        }
    }
}

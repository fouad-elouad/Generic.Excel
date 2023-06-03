using Generic.Excel.Export;
using Generic.Excel.Samples.Utilities;
using RandomSharp;

namespace Generic.Excel.Samples
{
    public class AddHorizontalSheetDictionary : SampleBase
    {
        public AddHorizontalSheetDictionary(DataFactory dataFactory) : base(dataFactory)
        {
            _DataFactory = dataFactory;
        }

        public AddHorizontalSheetDictionary() : base()
        {
        }

        public override void Run()
        {
            IRandomizer randomizer = new Randomizer();
            IDictionary<string, string> dico = new Dictionary<string, string>();
            for (int i = 0; i < 3; i++)
            {
                dico.Add(randomizer.String(4, 6), randomizer.String(8, 12));
            }

            using (ExcelFile excelFile = ExcelFile.Create())
            {
                excelFile.AddSheet(dico, "employees1", "key", "value", true);
                excelFile.AddSheet(dico, "employees2", horizontal: true);

                string filePath = FilePath(nameof(AddHorizontalSheetDictionary) + _ExcelExtension);
                excelFile.SaveAs(filePath);
            }
        }
    }
}

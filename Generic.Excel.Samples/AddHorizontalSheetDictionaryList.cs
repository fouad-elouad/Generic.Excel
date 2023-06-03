using Generic.Excel.Export;
using Generic.Excel.Samples.Utilities;

namespace Generic.Excel.Samples
{
    public class AddHorizontalSheetDictionaryList : SampleBase
    {
        public AddHorizontalSheetDictionaryList(DataFactory dataFactory) : base(dataFactory)
        {
            _DataFactory = dataFactory;
        }

        public AddHorizontalSheetDictionaryList() : base()
        {
        }

        public override void Run()
        {
            IDictionary<string, ICollection<string>> dico = new Dictionary<string, ICollection<string>>();
            for (int i = 0; i < 3; i++)
            {
                dico.Add("Company" + i, new List<string> { "employee" + i, "employee" + (i + 1), "employee" + (i + 2), "employee" + (i + 3) });
            }

            using (ExcelFile excelFile = ExcelFile.Create())
            {
                excelFile.AddSheet(dico, "employees1", "key", "value", horizontal: true);
                excelFile.AddSheet(dico, "employees2", horizontal: true);

                string filePath = FilePath(nameof(AddHorizontalSheetDictionaryList) + _ExcelExtension);
                excelFile.SaveAs(filePath);
            }
        }
    }
}

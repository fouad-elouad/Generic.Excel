using Generic.Excel.Export;
using Generic.Excel.Samples.Utilities;

namespace Generic.Excel.Samples
{
    public class AddSheetDatetimeDictionaryList : SampleBase
    {
        public AddSheetDatetimeDictionaryList(DataFactory dataFactory) : base(dataFactory)
        {
            _DataFactory = dataFactory;
        }

        public AddSheetDatetimeDictionaryList() : base()
        {
        }

        public override void Run()
        {
            IDictionary<int, ICollection<DateTime>> dico = new Dictionary<int, ICollection<DateTime>>();

            for (int i = 0; i < 3; i++)
            {
                dico.Add(2010 + i, new List<DateTime>
                { new DateTime(2010 + i,1,1), new DateTime(2010 + i,1,1).Date.AddDays(i + 1), new DateTime(2010 + i,1,1).Date.AddDays(i + 2) });
            }

            using (ExcelFile excelFile = ExcelFile.Create())
            {
                excelFile.AddSheet(dico, "employees1", "year", "control date");
                excelFile.AddSheet(dico, "employees2");

                string filePath = FilePath(nameof(AddSheetDatetimeDictionaryList) + _ExcelExtension);
                excelFile.SaveAs(filePath);
            }
        }
    }
}

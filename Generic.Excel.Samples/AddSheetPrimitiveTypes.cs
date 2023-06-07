using Generic.Excel.Export;
using Generic.Excel.Samples.Utilities;
using RandomSharp;
using RandomSharp.Impl;

namespace Generic.Excel.Samples
{
    public class AddSheetPrimitiveTypes : SampleBase
    {
        public AddSheetPrimitiveTypes(DataFactory dataFactory) : base(dataFactory)
        {
            _DataFactory = dataFactory;
        }

        public AddSheetPrimitiveTypes() : base()
        {
        }

        public override void Run()
        {
            IRandomizer randomizer = new DefaultRandomizer();
            IList<DateTime> dateTimes = new List<DateTime>();
            for (int i = 0; i < 10; i++)
            {
                dateTimes.Add(randomizer.Date(new DateTime(2010, 01, 01), new DateTime(2020, 01, 01)));
            }

            IList<string> strs = new List<string>();
            for (int i = 0; i < 10; i++)
            {
                strs.Add(randomizer.String(5, 10));
            }

            IList<int> integers = new List<int>();
            for (int i = 0; i < 10; i++)
            {
                integers.Add(randomizer.Int(0, 100));
            }

            using (ExcelFile excelFile = ExcelFile.Create())
            {
                excelFile.AddSheetList<DateTime>(dateTimes, nameof(dateTimes));
                excelFile.AddSheetList<DateTime>(dateTimes, nameof(dateTimes) + "with header", nameof(dateTimes));
                excelFile.AddSheet(strs, nameof(strs));
                excelFile.AddSheet(strs, nameof(strs) + "with header", nameof(strs));
                excelFile.AddSheetList<int>(integers, nameof(integers));
                excelFile.AddSheetList<int>(integers, nameof(integers) + "with header", nameof(integers));

                string filePath = FilePath(nameof(AddSheetPrimitiveTypes) + _ExcelExtension);
                excelFile.SaveAs(filePath);
            }
        }
    }
}

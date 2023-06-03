using Generic.Excel.Samples;

namespace Generic.Excel.UnitTest
{
    [TestClass]
    public class UnitTestGenericExcel
    {

        [TestMethod]
        public void TestAddSheet()
        {
            var sample = new AddSheet();
            sample.Run();
        }

        [TestMethod]
        public void TestAddSheetWithPropertiesToExport()
        {
            var sample = new AddSheetWithPropertiesToExport();
            sample.Run();
        }

        [TestMethod]
        public void TestAddSheetWithPropertiesDisplayName()
        {
            var sample = new AddSheetWithPropertiesDisplayName();
            sample.Run();
        }

        [TestMethod]
        public void TestAddHorizontalSheet()
        {
            var sample = new AddHorizontalSheet();
            sample.Run();
        }

        [TestMethod]
        public void TestAddHorizontalSheetWithPropertiesToExport()
        {
            var sample = new AddHorizontalSheetWithPropertiesToExport();
            sample.Run();
        }

        [TestMethod]
        public void TestAddHorizontalSheetWithPropertiesDisplayName()
        {
            var sample = new AddHorizontalSheetWithPropertiesDisplayName();
            sample.Run();
        }

        [TestMethod]
        public void TestAddBlankWorkSheet()
        {
            var sample = new AddBlankWorkSheet();
            sample.Run();
        }

        [TestMethod]
        public void TestAddSheetSingleObject()
        {
            var sample = new AddSheetSingleObject();
            sample.Run();
        }

        [TestMethod]
        public void TestAddSheetDictionary()
        {
            var sample = new AddSheetDictionary();
            sample.Run();
        }

        [TestMethod]
        public void TestAddHorizontalSheetDictionary()
        {
            var sample = new AddHorizontalSheetDictionary();
            sample.Run();
        }

        [TestMethod]
        public void TestAddSheetDictionaryList()
        {
            var sample = new AddSheetDictionaryList();
            sample.Run();
        }

        [TestMethod]
        public void TestAddHorizontalSheetDictionaryList()
        {
            var sample = new AddHorizontalSheetDictionaryList();
            sample.Run();
        }

        [TestMethod]
        public void TestAddSheetDatetimeDictionaryList()
        {
            var sample = new AddSheetDatetimeDictionaryList();
            sample.Run();
        }

        [TestMethod]
        public void TestAddSheetPrimitiveTypes()
        {
            var sample = new AddSheetPrimitiveTypes();
            sample.Run();
        }
    }
}

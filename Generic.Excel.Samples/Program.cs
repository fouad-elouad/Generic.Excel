// See https://aka.ms/new-console-template for more information
using Generic.Excel.Samples;

var sample1 = new AddSheet();
sample1.Run();

var sample2 = new AddSheetWithPropertiesToExport();
sample2.Run();

var sample3 = new AddSheetWithPropertiesDisplayName();
sample3.Run();

var sample4 = new AddHorizontalSheet();
sample4.Run();

var sample5 = new AddHorizontalSheetWithPropertiesToExport();
sample5.Run();

var sample6 = new AddHorizontalSheetWithPropertiesDisplayName();
sample6.Run();

var sample7 = new AddBlankWorkSheet();
sample7.Run();

var sample8 = new AddSheetSingleObject();
sample8.Run();

var sample9 = new AddSheetDictionary();
sample9.Run();

var sample10 = new AddHorizontalSheetDictionary();
sample10.Run();

var sample11 = new AddSheetDictionaryList();
sample11.Run();

var sample12 = new AddHorizontalSheetDictionaryList();
sample12.Run();

var sample13 = new AddSheetDatetimeDictionaryList();
sample13.Run();

var sample14 = new AddSheetPrimitiveTypes();
sample14.Run();

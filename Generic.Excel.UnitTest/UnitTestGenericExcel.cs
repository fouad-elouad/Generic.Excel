using Generic.Excel.Export;
using Generic.Excel.UnitTest.Models;
using Generic.Excel.UnitTest.Utilities;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using RandomSharp;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Generic.Excel.UnitTest
{
    [TestClass]
    public class UnitTestGenericExcel
    {
        IRandomizer randomizer;
        DataFactory dataFactory;

        [TestInitialize()]
        public void Initialize()
        {
            randomizer = new Randomizer();
            dataFactory = new DataFactory(randomizer);
        }

        [TestMethod]
        public void TestAddSheet()
        {
            IEnumerable<Employee> employees = dataFactory.Generate(10);

            using (ExcelFile excelFile = ExcelFile.Create())
            {
                excelFile.AddSheetList<Employee>(employees, nameof(employees));

                string filePath = FilePath(nameof(TestAddSheet) + ".xlsx");
                excelFile.SaveAs(filePath);
            }
        }

        [TestMethod]
        public void TestAddSheetWithPropertiesToExport()
        {
            IEnumerable<Employee> employees = dataFactory.Generate(10);

            using (ExcelFile excelFile = ExcelFile.Create())
            {
                var propertiesToExport = new List<string> { "EmployeeID", "IsSupervisor", "Department.FullName", "Salary" };
                excelFile.AddSheetList<Employee>(employees, nameof(employees), propertiesToExport);

                string filePath = FilePath(nameof(TestAddSheetWithPropertiesToExport) + ".xlsx");
                excelFile.SaveAs(filePath);
            }
        }

        [TestMethod]
        public void TestAddSheetWithPropertiesDisplayName()
        {
            IEnumerable<Employee> employees = dataFactory.Generate(10);

            using (ExcelFile excelFile = ExcelFile.Create())
            {
                var propertiesToExport = new List<string> { "EmployeeID", "IsSupervisor", "Department.FullName", "Salary" };
                var propertiesDisplayName = new List<string> { "display name 1", "display name 2", "display name 3", "display name 4" };
                excelFile.AddSheetList<Employee>(employees, nameof(employees), propertiesToExport, propertiesDisplayName);

                IDictionary<string, string> propertiesDictionary =
                    propertiesToExport.Zip(propertiesDisplayName, (k, v) => new { k, v }).ToDictionary(x => x.k, x => x.v);
                excelFile.AddSheetList<Employee>(employees, "Dictionary properties", propertiesDictionary);

                string filePath = FilePath(nameof(TestAddSheetWithPropertiesDisplayName) + ".xlsx");
                excelFile.SaveAs(filePath);
            }
        }

        [TestMethod]
        public void TestAddHorizontalSheet()
        {
            IEnumerable<Employee> employees = dataFactory.Generate(10);

            using (ExcelFile excelFile = ExcelFile.Create())
            {
                excelFile.AddSheetList<Employee>(employees, nameof(employees), horizontal: true);

                string filePath = FilePath(nameof(TestAddHorizontalSheet) + ".xlsx");
                excelFile.SaveAs(filePath);
            }
        }

        [TestMethod]
        public void TestAddHorizontalSheetWithPropertiesToExport()
        {
            IEnumerable<Employee> employees = dataFactory.Generate(10);

            using (ExcelFile excelFile = ExcelFile.Create())
            {
                var propertiesToExport = new List<string> { "EmployeeID", "IsSupervisor", "Department.FullName", "Salary" };
                excelFile.AddSheetList<Employee>(employees, nameof(employees), propertiesToExport, horizontal: true);

                string filePath = FilePath(nameof(TestAddHorizontalSheetWithPropertiesToExport) + ".xlsx");
                excelFile.SaveAs(filePath);
            }
        }

        [TestMethod]
        public void TestAddHorizontalSheetWithPropertiesDisplayName()
        {
            IEnumerable<Employee> employees = dataFactory.Generate(10);

            using (ExcelFile excelFile = ExcelFile.Create())
            {
                var propertiesToExport = new List<string> { "EmployeeID", "IsSupervisor", "Department.FullName", "Salary" };
                var propertiesDisplayName = new List<string> { "display name 1", "display name 2", "display name 3", "display name 4" };
                excelFile.AddSheetList<Employee>(employees, nameof(employees), propertiesToExport, propertiesDisplayName, true);

                string filePath = FilePath(nameof(TestAddHorizontalSheetWithPropertiesDisplayName) + ".xlsx");
                excelFile.SaveAs(filePath);
            }
        }

        [TestMethod]
        public void TestAddBlankWorkSheet()
        {
            using (ExcelFile excelFile = ExcelFile.Create())
            {
                excelFile.AddBlankWorkSheet("BlankSheet");

                string filePath = FilePath(nameof(TestAddBlankWorkSheet) + ".xlsx");
                excelFile.SaveAs(filePath);
            }
        }

        [TestMethod]
        public void TestAddSheetObject()
        {
            Employee employee = dataFactory.Generate(false);

            using (ExcelFile excelFile = ExcelFile.Create())
            {
                excelFile.AddSheetObj<Employee>(employee, "Default");
                var propertiesToExport = new List<string> { "EmployeeID", "IsSupervisor", "Department.FullName", "Salary" };
                excelFile.AddSheetObj<Employee>(employee, "propertiesToExport", propertiesToExport);
                var propertiesDisplayName = new List<string> { "display name 1", "display name 2", "display name 3", "display name 4" };
                excelFile.AddSheetObj<Employee>(employee, "propertiesDisplayName", propertiesToExport, propertiesDisplayName);
                excelFile.AddSheetObj<Employee>(employee, "Horizontal", horizontal: true);

                string filePath = FilePath(nameof(TestAddSheetObject) + ".xlsx");
                excelFile.SaveAs(filePath);
            }
        }

        [TestMethod]
        public void TestAddSheetDictionary()
        {
            IDictionary<string, string> dico = new Dictionary<string, string>();
            for (int i = 0; i < 3; i++)
            {
                dico.Add(randomizer.String(4, 6), randomizer.String(8, 12));
            }

            using (ExcelFile excelFile = ExcelFile.Create())
            {
                excelFile.AddSheet(dico, "employees1", "key", "value");
                excelFile.AddSheet(dico, "employees2");

                string filePath = FilePath(nameof(TestAddSheetDictionary) + ".xlsx");
                excelFile.SaveAs(filePath);
            }
        }

        [TestMethod]
        public void TestAddHorizontalSheetDictionary()
        {
            IDictionary<string, string> dico = new Dictionary<string, string>();
            for (int i = 0; i < 3; i++)
            {
                dico.Add(randomizer.String(4, 6), randomizer.String(8, 12));
            }

            using (ExcelFile excelFile = ExcelFile.Create())
            {
                excelFile.AddSheet(dico, "employees1", "key", "value", true);
                excelFile.AddSheet(dico, "employees2", horizontal: true);

                string filePath = FilePath(nameof(TestAddHorizontalSheetDictionary) + ".xlsx");
                excelFile.SaveAs(filePath);
            }
        }

        [TestMethod]
        public void TestAddSheetDictionaryList()
        {
            IDictionary<string, ICollection<string>> dico = new Dictionary<string, ICollection<string>>();
            for (int i = 0; i < 3; i++)
            {
                dico.Add("Company" + i, new List<string> { "employee" + i, "employee" + (i + 1), "employee" + (i + 2), "employee" + (i + 3) });
            }

            using (ExcelFile excelFile = ExcelFile.Create())
            {
                excelFile.AddSheet(dico, "employees1", "key", "value");
                excelFile.AddSheet(dico, "employees2");

                string filePath = FilePath(nameof(TestAddSheetDictionaryList) + ".xlsx");
                excelFile.SaveAs(filePath);
            }
        }

        [TestMethod]
        public void TestAddHorizontalSheetDictionaryList()
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

                string filePath = FilePath(nameof(TestAddHorizontalSheetDictionaryList) + ".xlsx");
                excelFile.SaveAs(filePath);
            }
        }

        [TestMethod]
        public void TestAddSheetDatetimeDictionaryList()
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

                string filePath = FilePath(nameof(TestAddSheetDatetimeDictionaryList) + ".xlsx");
                excelFile.SaveAs(filePath);
            }
        }

        [TestMethod]
        public void TestAddSheetPrimitiveTypes()
        {
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

                string filePath = FilePath(nameof(TestAddSheetPrimitiveTypes) + ".xlsx");
                excelFile.SaveAs(filePath);
            }
        }

        [TestMethod]
        public void TestWorkSheetsCount()
        {
            Employee employee = dataFactory.Generate(false);
            using (ExcelFile excelFile = ExcelFile.Create())
            {
                excelFile.AddSheetObj<Employee>(employee, nameof(employee));
                Assert.AreEqual(1, excelFile.WorkSheetsCount());
                excelFile.AddSheetObj<Employee>(employee, nameof(employee) + 2);
                Assert.AreEqual(2, excelFile.WorkSheetsCount());
                excelFile.AddBlankWorkSheet(nameof(employee) + 3);
                Assert.AreEqual(3, excelFile.WorkSheetsCount());
            }
        }

        private string FilePath(string fileName)
        {
            string directoryPath = AppDomain.CurrentDomain.BaseDirectory;
            return Path.Combine(directoryPath, fileName);
        }
    }
}

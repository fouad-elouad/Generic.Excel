using Generic.Excel.UnitTest.Models;
using RandomSharp;
using System;
using System.Collections.Generic;

namespace Generic.Excel.UnitTest.Utilities
{
    public class DataFactory
    {

        protected IRandomizer Randomizer { get; set; }

        public DataFactory(IRandomizer randomizer)
        {
            Randomizer = randomizer;
        }

        public Employee Generate(bool allowNull = true)
        {
            return new Employee
            {
                FirstName = Randomizer.RandomString(10),
                LastName = Randomizer.RandomString(10),
                BirthDate = Randomizer.RandomDate(new DateTime(1970, 01, 01), new DateTime(2000, 01, 01)),
                LastConnexion = Randomizer.RandomDateTime(new DateTime(2000, 01, 01), new DateTime(2020, 01, 01)),
                MaritalStatus = Randomizer.RandomEnum<MaritalStatus>(),
                EmployeeID = Randomizer.Random(100),
                IsSupervisor = allowNull ? Randomizer.RandomNullableBool() : Randomizer.RandomBool(),
                SocialSecurityNumber = Randomizer.Random(123456789, 1234567890),
                Points = Randomizer.Random(10M, 100M),
                Salary = Randomizer.Random(10000, 30000),
                Address = Randomizer.RandomString(30, 50),
                Skills = Randomizer.RandomString(100, 120),
                Department = allowNull ? Randomizer.RandomNullable<Department>(() => new Department // null or new Department
                {
                    ShortName = Randomizer.RandomString(7),
                    FullName = Randomizer.RandomString(20),
                    TeamSize = Randomizer.Random(11, 200),
                }) : new Department // null or new Department
                {
                    ShortName = Randomizer.RandomString(7),
                    FullName = Randomizer.RandomString(20),
                    TeamSize = Randomizer.Random(11, 200),
                },
            };
        }

        public IEnumerable<Employee> Generate(int count, bool allowNull = true)
        {
            IList<Employee> list = new List<Employee>(count);
            for (int i = 0; i < count; i++)
            {
                list.Add(Generate(allowNull));
            }
            return list;
        }
    }
}

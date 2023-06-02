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
                FirstName = Randomizer.String(10),
                LastName = Randomizer.String(10),
                BirthDate = Randomizer.Date(new DateTime(1970, 01, 01), new DateTime(2000, 01, 01)),
                LastConnexion = Randomizer.DateTime(new DateTime(2000, 01, 01), new DateTime(2020, 01, 01)),
                MaritalStatus = Randomizer.Enumeration<MaritalStatus>(),
                EmployeeID = Randomizer.Int(100),
                IsSupervisor = allowNull ? Randomizer.NullableBoolean() : Randomizer.Boolean(),
                SocialSecurityNumber = Randomizer.Int(123456789, 1234567890),
                Points = Randomizer.Decimal(10M, 100M),
                Salary = Randomizer.Int(10000, 30000),
                Address = Randomizer.String(30, 50, StringCharacterType.UppercaseAlphaNumeric),
                Skills = Randomizer.String(100, 120),
                Department = allowNull ? Randomizer.Nullable<Department>(() => new Department // null or new Department
                {
                    ShortName = Randomizer.String(7),
                    FullName = Randomizer.String(20),
                    TeamSize = Randomizer.Int(11, 200),
                }) : new Department // null or new Department
                {
                    ShortName = Randomizer.String(7),
                    FullName = Randomizer.String(20),
                    TeamSize = Randomizer.Int(11, 200),
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

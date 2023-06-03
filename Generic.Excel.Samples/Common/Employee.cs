using Generic.Excel.Export;
using System;

namespace Generic.Excel.Samples.Common
{
    public class Employee
    {
        [ExcelProperty(ignore: true)]
        public int EmployeeID { get; set; }

        [ExcelProperty("First Name", 2)]
        public string FirstName { get; set; }

        [ExcelProperty("Last Name", 1)]
        public string LastName { get; set; }

        public DateTime? BirthDate { get; set; }

        public DateTime? LastConnexion { get; set; }

        public string Address { get; set; }

        [ExcelProperty(ignore: true)]
        public string Skills { get; set; }

        [ExcelProperty("Social Security Number", order: 3)]
        public int? SocialSecurityNumber { get; set; }

        public decimal? Points { get; set; }

        public double? Salary { get; set; }

        [ExcelProperty(order: 4)]
        public bool? IsSupervisor { get; set; }

        public MaritalStatus? MaritalStatus { get; set; }

        [ExcelProperty("Department ShortName", 5, nameof(Common.Department.ShortName))]
        public Department Department { get; set; }
    }

    public enum MaritalStatus { Single, Married, Divorced }
}

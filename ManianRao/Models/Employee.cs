using System;


namespace TheChartedCompany.Models
{
    public class Employee
    {
        public string Name { get; set; }
        public double LeaveCredit { get; set; }
        public string Type { get; set; }
        public int Id { get; set; }
        public String DataOfJoining { get; set; }
        public double Salary { get; set; }
        public double CustomSalary { get; set; }
        public int Bonus { get; set; }
        public int Conveyance { get; set; }

        public string Status { get; set; }
        public string phone { get; set; }
        public string AccountNo { get; set; }
        public int MinimumStipend { get; set; }

        public int TakeCustomSalaryYN { get; set; }
        public double Leaves { get; set; }

        public int BasicSal { get; set; }

        public int Allowance { get; set; }
        public int HRA { get; set; }
        public int TDS { get; set; }

        public int PT { get; set; }

        public double LeavesTakenThisMonth { get; set; }

        public double NewLeavesRepo { get; set; }


    }
}
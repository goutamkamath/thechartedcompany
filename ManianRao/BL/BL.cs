using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using Npgsql;
using TheChartedCompany.Models;


namespace TheChartedCompany.BL
{
    public class BL
    {
        string connectionString = ConfigurationManager.AppSettings["ConnectionString"];

        public bool AddNewEmployee(Employee emp)
        {

            string strConnString = connectionString;
            APIResult result = new APIResult();

            NpgsqlConnection objConn = new NpgsqlConnection(strConnString);
            try
            {


                using (NpgsqlConnection connection = new NpgsqlConnection())
                {
                    connection.ConnectionString = strConnString;
                    connection.Open();
                    NpgsqlCommand cmd = new NpgsqlCommand();
                    cmd.Connection = connection;
                    cmd.CommandText = "insert into tbl_employee_goutamtest(name,type,dateofjoin,customsalary,takecustomsalaryyn) values(@name,@type,@dateofjoin,@customsalary,@takecustomsalaryyn)";
                    cmd.CommandType = CommandType.Text;

                    cmd.Parameters.Add(new NpgsqlParameter("@name", emp.Name));
                    cmd.Parameters.Add(new NpgsqlParameter("@type", emp.Type));
                    cmd.Parameters.Add(new NpgsqlParameter("@dateofjoin", emp.DataOfJoining));
                    //cmd.Parameters.Add(new NpgsqlParameter("@phone", emp.phone));
                    cmd.Parameters.Add(new NpgsqlParameter("@customsalary", emp.CustomSalary));
                    cmd.Parameters.Add(new NpgsqlParameter("@takecustomsalaryyn", emp.TakeCustomSalaryYN));

                    //cmd.Parameters.Add(new NpgsqlParameter("@minimumstipend", emp.MinimumStipend));
                    // cmd.Parameters.AddWithValue("@AccountNo", NpgsqlTypes.NpgsqlDbType.Text, (object)emp.AccountNo ?? DBNull.Value);



                    cmd.ExecuteNonQuery();
                    cmd.Dispose();
                    connection.Close();
                    NpgsqlConnection.ClearPool(connection);
                    result.Status = true;
                }
            }
            catch (Exception ex)
            {

                result.Status = false;
                throw ex;
            }
            finally
            {
                NpgsqlConnection.ClearPool(objConn);
                objConn.Close();
                objConn.Dispose();
            }




            return true;
        }



        
             public bool creditfixedleaves(double leavecount)
        {

            string strConnString = connectionString;
            APIResult result = new APIResult();

            NpgsqlConnection objConn = new NpgsqlConnection(strConnString);
            try
            {


                using (NpgsqlConnection connection = new NpgsqlConnection())
                {
                    connection.ConnectionString = strConnString;
                    connection.Open();
                    NpgsqlCommand cmd = new NpgsqlCommand();
                    cmd.Connection = connection;
                    cmd.CommandText = "update tbl_employee_goutamtest set leavesrepo=leavesrepo+"+ leavecount ;
                    cmd.CommandType = CommandType.Text;


                    cmd.ExecuteNonQuery();
                    cmd.Dispose();
                    connection.Close();
                    NpgsqlConnection.ClearPool(connection);
                    result.Status = true;
                }
            }
            catch (Exception ex)
            {

                result.Status = false;
                throw ex;
            }
            finally
            {
                NpgsqlConnection.ClearPool(objConn);
                objConn.Close();
                objConn.Dispose();
            }




            return result.Status;
        }

        public bool SyncLeaves()
        {

            string strConnString = connectionString;
            APIResult result = new APIResult();

            NpgsqlConnection objConn = new NpgsqlConnection(strConnString);
            try
            {


                using (NpgsqlConnection connection = new NpgsqlConnection())
                {
                    connection.ConnectionString = strConnString;
                    connection.Open();
                    NpgsqlCommand cmd = new NpgsqlCommand();
                    cmd.Connection = connection;
                    cmd.CommandText = "UPDATE tbl_employee_goutamtest a SET leavesrepo = b.newleavesrepo FROM tbl_employee_sal_goutamtest b JOIN tbl_employee_goutamtest c ON b.name = c.name  WHERE  a.name = c.name";
                    cmd.CommandType = CommandType.Text;


                    cmd.ExecuteNonQuery();
                    cmd.Dispose();
                    connection.Close();
                    NpgsqlConnection.ClearPool(connection);
                    result.Status = true;
                }
            }
            catch (Exception ex)
            {

                result.Status = false;
                throw ex;
            }
            finally
            {
                NpgsqlConnection.ClearPool(objConn);
                objConn.Close();
                objConn.Dispose();
            }




            return result.Status;
        }
        public bool EditEmployee(Employee emp)
        {

            string strConnString = connectionString;
            APIResult result = new APIResult();

            NpgsqlConnection objConn = new NpgsqlConnection(strConnString);
            try
            {


                using (NpgsqlConnection connection = new NpgsqlConnection())
                {
                    connection.ConnectionString = strConnString;
                    connection.Open();
                    NpgsqlCommand cmd = new NpgsqlCommand();
                    cmd.Connection = connection;
                    cmd.CommandText = "update tbl_employee_goutamtest set name=@name,type=@type,dateofjoin=@dateofjoin,accountno=@AccountNo,customsalary=@customsalary,takecustomsalaryyn=@takecustomsalaryyn,basicsal=@basicsal,hra=@hra,tds=@tds,allowance=@allowance where  id=@id";
                    cmd.CommandType = CommandType.Text;

                    cmd.Parameters.Add(new NpgsqlParameter("@name", emp.Name));
                    cmd.Parameters.Add(new NpgsqlParameter("@type", emp.Type));
                    cmd.Parameters.Add(new NpgsqlParameter("@dateofjoin", emp.DataOfJoining));
                    // cmd.Parameters.Add(new NpgsqlParameter("@phone", emp.phone));
                    cmd.Parameters.Add(new NpgsqlParameter("@id", emp.Id));
                    // cmd.Parameters.Add(new NpgsqlParameter("@minimumstipend", emp.MinimumStipend));
                    cmd.Parameters.Add(new NpgsqlParameter("@customsalary", emp.CustomSalary));
                    cmd.Parameters.Add(new NpgsqlParameter("@takecustomsalaryyn", emp.TakeCustomSalaryYN));
                    cmd.Parameters.Add(new NpgsqlParameter("@basicsal", emp.BasicSal));
                    cmd.Parameters.Add(new NpgsqlParameter("@hra", emp.HRA));
                    cmd.Parameters.Add(new NpgsqlParameter("@tds", emp.TDS));
                    cmd.Parameters.Add(new NpgsqlParameter("@allowance", emp.Allowance));


                    cmd.Parameters.AddWithValue("@AccountNo", NpgsqlTypes.NpgsqlDbType.Text, (object)emp.AccountNo ?? DBNull.Value);



                    cmd.ExecuteNonQuery();
                    cmd.Dispose();
                    connection.Close();
                    NpgsqlConnection.ClearPool(connection);
                    result.Status = true;
                }
            }
            catch (Exception ex)
            {

                result.Status = false;
                throw ex;
            }
            finally
            {
                NpgsqlConnection.ClearPool(objConn);
                objConn.Close();
                objConn.Dispose();
            }




            return true;
        }


        public bool DeleteHoliday(string ID)
        {
            string strConnString = connectionString;
            APIResult result = new APIResult();

            NpgsqlConnection objConn = new NpgsqlConnection(strConnString);
            try
            {


                using (NpgsqlConnection connection = new NpgsqlConnection())
                {
                    connection.ConnectionString = strConnString;
                    connection.Open();
                    NpgsqlCommand cmd = new NpgsqlCommand();
                    cmd.Connection = connection;
                    cmd.CommandText = "delete from tbl_holidays_goutamtest  where  id=@id";
                    cmd.CommandType = CommandType.Text;

                    cmd.Parameters.Add(new NpgsqlParameter("@id", Convert.ToInt32(ID)));


                    cmd.ExecuteNonQuery();
                    cmd.Dispose();
                    connection.Close();
                    NpgsqlConnection.ClearPool(connection);
                    result.Status = true;
                }
            }
            catch (Exception ex)
            {

                result.Status = false;
                throw ex;
            }
            finally
            {
                NpgsqlConnection.ClearPool(objConn);
                objConn.Close();
                objConn.Dispose();
            }




            return result.Status;




        }



        public bool DeleteEmployee(string id)
        {

            string strConnString = connectionString;
            APIResult result = new APIResult();

            NpgsqlConnection objConn = new NpgsqlConnection(strConnString);
            try
            {


                using (NpgsqlConnection connection = new NpgsqlConnection())
                {
                    connection.ConnectionString = strConnString;
                    connection.Open();
                    NpgsqlCommand cmd = new NpgsqlCommand();
                    cmd.Connection = connection;
                    cmd.CommandText = "delete from tbl_employee_goutamtest where  id=@id";
                    cmd.CommandType = CommandType.Text;


                    cmd.Parameters.Add(new NpgsqlParameter("@id", Convert.ToInt32(id)));


                    cmd.ExecuteNonQuery();
                    cmd.Dispose();
                    connection.Close();
                    NpgsqlConnection.ClearPool(connection);
                    result.Status = true;
                }
            }
            catch (Exception ex)
            {

                result.Status = false;
                throw ex;
            }
            finally
            {
                NpgsqlConnection.ClearPool(objConn);
                objConn.Close();
                objConn.Dispose();
            }




            return result.Status;
        }

        public bool EditHoliday(Holiday hol)
        {

            string strConnString = connectionString;
            APIResult result = new APIResult();

            NpgsqlConnection objConn = new NpgsqlConnection(strConnString);
            try
            {


                using (NpgsqlConnection connection = new NpgsqlConnection())
                {
                    connection.ConnectionString = strConnString;
                    connection.Open();
                    NpgsqlCommand cmd = new NpgsqlCommand();
                    cmd.Connection = connection;
                    cmd.CommandText = "update tbl_holidays_goutamtest set holidayfor=@holidayfor,holidayon=@holidayon where id=@id";
                    cmd.CommandType = CommandType.Text;

                    cmd.Parameters.Add(new NpgsqlParameter("@holidayfor", hol.HolidayFor));
                    cmd.Parameters.Add(new NpgsqlParameter("@id", hol.Id));
                    cmd.Parameters.Add(new NpgsqlParameter("@holidayon", hol.HolidayOn));



                    cmd.ExecuteNonQuery();
                    cmd.Dispose();
                    connection.Close();
                    NpgsqlConnection.ClearPool(connection);
                    result.Status = true;

                }
            }
            catch (Exception ex)
            {

                result.Status = false;
                throw ex;
            }
            finally
            {
                NpgsqlConnection.ClearPool(objConn);
                objConn.Close();
                objConn.Dispose();
            }




            return true;
        }


        public bool SaveLeaves(string name, double repo)
        {

            string strConnString = connectionString;
            APIResult result = new APIResult();

            NpgsqlConnection objConn = new NpgsqlConnection(strConnString);
            try
            {


                using (NpgsqlConnection connection = new NpgsqlConnection())
                {
                    connection.ConnectionString = strConnString;
                    connection.Open();
                    NpgsqlCommand cmd = new NpgsqlCommand();
                    cmd.Connection = connection;
                    cmd.CommandText = "update tbl_employee_goutamtest set leavesrepo=@leavesrepo where name=@name ";
                    cmd.CommandType = CommandType.Text;

                    cmd.Parameters.Add(new NpgsqlParameter("@name", name));
                    cmd.Parameters.Add(new NpgsqlParameter("@leavesrepo", repo));

                    cmd.ExecuteNonQuery();
                    cmd.Dispose();
                    connection.Close();
                    NpgsqlConnection.ClearPool(connection);
                    result.Status = true;
                }

            }
            catch (Exception ex)
            {

                result.Status = false;
                throw ex;
            }
            finally
            {
                NpgsqlConnection.ClearPool(objConn);
                objConn.Close();
                objConn.Dispose();
            }




            return true;
        }
        public bool EditPayroll(Employee emp)
        {

            string strConnString = connectionString;
            APIResult result = new APIResult();

            NpgsqlConnection objConn = new NpgsqlConnection(strConnString);
            try
            {


                using (NpgsqlConnection connection = new NpgsqlConnection())
                {
                    connection.ConnectionString = strConnString;
                    connection.Open();
                    NpgsqlCommand cmd = new NpgsqlCommand();
                    cmd.Connection = connection;
                    cmd.CommandText = "update tbl_employee_sal_goutamtest set salary=@salary,bonus=@bonus,conveyance=@conveyance where id=@id ";
                    cmd.CommandType = CommandType.Text;

                    cmd.Parameters.Add(new NpgsqlParameter("@bonus", emp.Bonus));
                    cmd.Parameters.Add(new NpgsqlParameter("@salary", emp.Salary));
                    cmd.Parameters.Add(new NpgsqlParameter("@conveyance", emp.Conveyance));
                    cmd.Parameters.Add(new NpgsqlParameter("@id", emp.Id));



                    cmd.ExecuteNonQuery();
                    cmd.Dispose();
                    connection.Close();
                    NpgsqlConnection.ClearPool(connection);
                    result.Status = true;
                }

            }
            catch (Exception ex)
            {

                result.Status = false;
                throw ex;
            }
            finally
            {
                NpgsqlConnection.ClearPool(objConn);
                objConn.Close();
                objConn.Dispose();
            }




            return true;
        }

        public bool AddNewHoliday(List<Holiday> hol)
        {

            string strConnString = connectionString;
            APIResult result = new APIResult();

            NpgsqlConnection objConn = new NpgsqlConnection(strConnString);
            try
            {

                foreach (Holiday item in hol)
                {
                    using (NpgsqlConnection connection = new NpgsqlConnection())
                    {
                        connection.ConnectionString = strConnString;
                        connection.Open();
                        NpgsqlCommand cmd = new NpgsqlCommand();
                        cmd.Connection = connection;
                        cmd.CommandText = "insert into tbl_holidays_goutamtest(holidayfor,holidayon) values(@holidayfor,@holidayon)";
                        cmd.CommandType = CommandType.Text;

                        cmd.Parameters.Add(new NpgsqlParameter("@holidayfor", item.HolidayFor));
                        cmd.Parameters.Add(new NpgsqlParameter("@holidayon", item.HolidayOn));

                        // cmd.Parameters.AddWithValue("@holidayon", NpgsqlTypes.NpgsqlDbType.Date, (object)item.HolidayOn ?? DBNull.Value);


                        cmd.ExecuteNonQuery();
                        cmd.Dispose();
                        connection.Close();
                        NpgsqlConnection.ClearPool(connection);
                        result.Status = true;
                    }
                }
            }
            catch (Exception ex)
            {

                result.Status = false;
                throw ex;
            }
            finally
            {
                NpgsqlConnection.ClearPool(objConn);
                objConn.Close();
                objConn.Dispose();
            }




            return true;
        }

        public bool MakeEntry(Employee emp, bool isemparticle)
        {

            string strConnString = connectionString;
            APIResult result = new APIResult();

            NpgsqlConnection objConn = new NpgsqlConnection(strConnString);
            try
            {


                using (NpgsqlConnection connection = new NpgsqlConnection())
                {
                    connection.ConnectionString = strConnString;
                    connection.Open();
                    NpgsqlCommand cmd = new NpgsqlCommand();
                    cmd.Connection = connection;
                    if (isemparticle)
                    {
                        cmd.CommandText = "update tbl_employee_sal_goutamtest set salary=@salary,status=@status,leavestakenthismnth=@leavestakenthismnth where name=@name ";
                    }
                    else
                    {
                        cmd.CommandText = "update tbl_employee_sal_goutamtest set salary=@salary,status=@status,pt=200,leavestakenthismnth=@leavestakenthismnth,leavecredit=@leavecredit where name=@name ";
                    }


                    cmd.CommandType = CommandType.Text;

                    cmd.Parameters.Add(new NpgsqlParameter("@name", emp.Name));
                    cmd.Parameters.Add(new NpgsqlParameter("@leavecredit", emp.LeaveCredit));
                    
                    cmd.Parameters.Add(new NpgsqlParameter("@salary", emp.Salary));
                    cmd.Parameters.Add(new NpgsqlParameter("@leavestakenthismnth", emp.LeavesTakenThisMonth));
                    cmd.Parameters.Add(new NpgsqlParameter("@status", emp.Status));



                    cmd.ExecuteNonQuery();
                    cmd.Dispose();
                    connection.Close();
                    NpgsqlConnection.ClearPool(connection);
                    result.Status = true;
                }

            }
            catch (Exception ex)
            {

                result.Status = false;
                throw ex;
            }
            finally
            {
                NpgsqlConnection.ClearPool(objConn);
                objConn.Close();
                objConn.Dispose();
            }




            return true;
        }

        public bool MakeDefaultPayrollEntry(int monthId)
        {

            string strConnString = connectionString;
            APIResult result = new APIResult();

            NpgsqlConnection objConn = new NpgsqlConnection(strConnString);
            try
            {


                using (NpgsqlConnection connection = new NpgsqlConnection())
                {
                    connection.ConnectionString = strConnString;
                    connection.Open();
                    NpgsqlCommand cmd = new NpgsqlCommand();
                    cmd.Connection = connection;

                    cmd.CommandText = "insert into tbl_employee_sal_goutamtest(name,leavesrepo) select name,leavesrepo from tbl_employee_goutamtest";

                    cmd.CommandType = CommandType.Text;



                    cmd.ExecuteNonQuery();
                    cmd.Dispose();
                    connection.Close();
                    NpgsqlConnection.ClearPool(connection);
                    result.Status = true;
                }

                List<Employee> articlelist = new List<Employee>();
                articlelist = GetOnlyEmployeeDetails("","N",true);
                foreach(Employee e in articlelist)
                {
                    using (NpgsqlConnection connection = new NpgsqlConnection())
                    {
                        connection.ConnectionString = strConnString;
                        connection.Open();
                        NpgsqlCommand cmd = new NpgsqlCommand();
                        cmd.Connection = connection;
                        int minsalary = GetSalaryByJoiningDate(e.DataOfJoining,true,0, 0,monthId);
                        
                        cmd.CommandText = "update tbl_employee_sal_goutamtest set salary="+minsalary+ " where name='"+e.Name+"' ";

                        cmd.CommandType = CommandType.Text;



                        cmd.ExecuteNonQuery();
                        cmd.Dispose();
                        connection.Close();
                        NpgsqlConnection.ClearPool(connection);
                        result.Status = true;
                    }


                }

            }
            catch (Exception ex)
            {

                result.Status = false;
                throw ex;
            }

            finally
            {
                NpgsqlConnection.ClearPool(objConn);
                objConn.Close();
                objConn.Dispose();
            }




            return true;
        }


        public List<Holiday> GetHolidayDetails(string ID = "", string subregParamYN = "N")
        {

            string strConnString = connectionString;
            string whereClause;
            if (ID == "")
            {
                whereClause = "";
            }
            else
            {
                if (subregParamYN == "Y")
                {
                    whereClause = " where id in(" + ID + ")";
                }
                else
                {
                    whereClause = " where id in(" + ID + ")";
                }

            }
            List<Holiday> holList = new List<Holiday>();
            try
            {
                NpgsqlConnection objConn = new NpgsqlConnection(strConnString);
                objConn.Open();
                string strSelectCmd = "select id,holidayfor,holidayon from tbl_holidays_goutamtest " + whereClause + " order by holidayon ";

                NpgsqlCommand command = new NpgsqlCommand(strSelectCmd, objConn);
                command.CommandType = CommandType.Text;

                NpgsqlDataReader dr = command.ExecuteReader();

                while (dr.Read())
                {

                    holList.Add(
                                 new Holiday
                                 {

                                     Id = Convert.ToInt32(dr[0]),

                                     HolidayFor = Convert.ToString(dr[1]),
                                     HolidayOn = Convert.ToString(dr[2])



                                 }
                                     ); ;
                }

                NpgsqlConnection.ClearPool(objConn);
                objConn.Close();
                objConn.Dispose();
            }
            catch (Exception e)
            {
                throw e;
            }
            return holList;
        }



        public List<Employee> GetEmployeeDetails(string ID = "", string subregParamYN = "N", string getdetbyname = "")
        {


            string strConnString = connectionString;
            string whereClause;
            if (ID == "" && getdetbyname == "")
            {
                whereClause = "";
            }
            else if (getdetbyname != "")
            {
                whereClause = " where name ='" + getdetbyname + "' ";
            }
            else
            {
                if (subregParamYN == "Y")
                {
                    whereClause = " where id in(" + ID + ")";
                }
                else
                {
                    whereClause = " where id in(" + ID + ")";
                }

            }
            List<Employee> empList = new List<Employee>();
            try
            {
                NpgsqlConnection objConn = new NpgsqlConnection(strConnString);
                objConn.Open();
                string strSelectCmd = "select id,name,type,dateofjoin,coalesce(salary,0),coalesce(bonus, 0) ,coalesce(conveyance, 0) ,coalesce(status, '') ,coalesce(customsalary, 0),coalesce(takecustomsalaryyn, 0),coalesce(basicsal, 0),coalesce(hra, 0),coalesce(tds, 0),coalesce(allowance, 0)  from tbl_employee_goutamtest " + whereClause + " order by name";

                NpgsqlCommand command = new NpgsqlCommand(strSelectCmd, objConn);
                command.CommandType = CommandType.Text;

                NpgsqlDataReader dr = command.ExecuteReader();

                while (dr.Read())
                {

                    empList.Add(
                                 new Employee
                                 {

                                     Id = Convert.ToInt32(dr[0]),

                                     Name = Convert.ToString(dr[1]),
                                     DataOfJoining = Convert.ToString(dr[3]),
                                     Type = Convert.ToString(dr[2]),
                                     Salary = Convert.ToInt32(dr[4]),
                                     Bonus = Convert.ToInt32(dr[5]),
                                     Conveyance = Convert.ToInt32(dr[6]),
                                     Status = Convert.ToString(dr[7]),
                                     // phone = Convert.ToString(dr[8]),
                                     CustomSalary = Convert.ToInt32(dr[8]),
                                     TakeCustomSalaryYN = Convert.ToBoolean(dr[9]) ? 1 : 0,
                                     BasicSal = Convert.ToInt32(dr[10]),
                                     HRA = Convert.ToInt32(dr[11]),
                                     TDS = Convert.ToInt32(dr[12]),
                                     Allowance = Convert.ToInt32(dr[13])

                                 }
                                     ); ;
                }

                NpgsqlConnection.ClearPool(objConn);
                objConn.Close();
                objConn.Dispose();
            }
            catch (Exception e)
            {
                throw e;
            }
            return empList;
        }


        public List<Employee> GetOnlyEmployeeDetails(string ID = "", string subregParamYN = "N", bool isEmpArticle = false)
        {

            string strConnString = connectionString;
            string whereClause;
            if (ID == "")
            {
                whereClause = "";
            }
            else
            {
                if (subregParamYN == "Y")
                {
                    whereClause = " and id in(" + ID + ")";
                }
                else
                {
                    whereClause = " and id in(" + ID + ")";
                }

            }
            List<Employee> empList = new List<Employee>();
            string strSelectCmd = "";
            try
            {
                NpgsqlConnection objConn = new NpgsqlConnection(strConnString);
                objConn.Open();
                if (!isEmpArticle)
                {
                    strSelectCmd = "select id,name,type,dateofjoin,coalesce(salary,0),coalesce(bonus, 0) ,coalesce(conveyance, 0) ,coalesce(status, '') ,coalesce(customsalary, 0),coalesce(takecustomsalaryyn, 0),coalesce(leavesrepo, 0)  from tbl_employee_goutamtest where type='Employee' " + whereClause + " order by name";
                }
                else
                {
                    strSelectCmd = "select id,name,type,dateofjoin,coalesce(salary,0),coalesce(bonus, 0) ,coalesce(conveyance, 0) ,coalesce(status, '') ,coalesce(customsalary, 0),coalesce(takecustomsalaryyn, 0),coalesce(leavesrepo, 0)  from tbl_employee_goutamtest where type='Article' " + whereClause + " order by name";
                }


                NpgsqlCommand command = new NpgsqlCommand(strSelectCmd, objConn);
                command.CommandType = CommandType.Text;

                NpgsqlDataReader dr = command.ExecuteReader();

                while (dr.Read())
                {

                    empList.Add(
                                 new Employee
                                 {

                                     Id = Convert.ToInt32(dr[0]),

                                     Name = Convert.ToString(dr[1]),
                                     DataOfJoining = Convert.ToString(dr[3]),
                                     Type = Convert.ToString(dr[2]),
                                     Salary = Convert.ToInt32(dr[4]),
                                     Bonus = Convert.ToInt32(dr[5]),
                                     Conveyance = Convert.ToInt32(dr[6]),
                                     Status = Convert.ToString(dr[7]),
                                     // phone = Convert.ToString(dr[8]),
                                     CustomSalary = Convert.ToInt32(dr[8]),
                                     TakeCustomSalaryYN = Convert.ToBoolean(dr[9]) ? 1 : 0,
                                     Leaves = Convert.ToDouble(dr[10])


                                 }
                                     ); ;
                }

                NpgsqlConnection.ClearPool(objConn);
                objConn.Close();
                objConn.Dispose();
            }
            catch (Exception e)
            {
                throw e;
            }
            return empList;
        }

        public int GetSalaryByJoiningDate(string doj, bool isminwage = false, int takecustomsalaryyn = 0, int customsalary = 0, int monthId = 0)
        {
            string strConnString = connectionString;
            MinumumWage mw = new MinumumWage();
            DateTime dojTime = new DateTime();


            try
            {
                NpgsqlConnection objConn = new NpgsqlConnection(strConnString);
                objConn.Open();
                string strSelectCmd = "select minimumwagearticleone,minimumwagearticletwo,minimumwagearticlethree,minimumwagearticleonesal,minimumwagearticletwosal,minimumwagearticlethreesal from tbl_org_goutamtest ";

                NpgsqlCommand command = new NpgsqlCommand(strSelectCmd, objConn);
                command.CommandType = CommandType.Text;

                NpgsqlDataReader dr = command.ExecuteReader();

                while (dr.Read())
                {

                    mw.ArticleFirstYear = Convert.ToString(dr[0]);

                    mw.ArticleSecondYear = Convert.ToString(dr[1]);
                    mw.ArticleThirdYear = Convert.ToString(dr[2]);
                    mw.ArticleSalFirstYear = Convert.ToString(dr[3]);

                    mw.ArticleSalSecYear = Convert.ToString(dr[4]);
                    mw.ArticleSalThirdYear = Convert.ToString(dr[5]);

                }
                if (takecustomsalaryyn == 1)
                {
                    mw.ArticleSalFirstYear = Convert.ToString(customsalary);

                    mw.ArticleSalSecYear = Convert.ToString(customsalary);
                    mw.ArticleSalThirdYear = Convert.ToString(customsalary);
                }


                NpgsqlConnection.ClearPool(objConn);
                objConn.Close();
                objConn.Dispose();

                dojTime = Convert.ToDateTime(doj);
                var firstDayOfMonth = new DateTime(DateTime.Now.Year, monthId != 0 ? monthId : DateTime.Now.Month, 1);
                var fifteenthOfMonth = firstDayOfMonth.AddDays(14);
                if (!isminwage)
                {
                    if (dojTime.AddYears(1) >= fifteenthOfMonth)
                    {
                        return Convert.ToInt32(mw.ArticleSalFirstYear);
                    }
                    else if ((dojTime.AddYears(2) >= fifteenthOfMonth) && (dojTime.AddYears(3) >= fifteenthOfMonth))
                    {
                        return Convert.ToInt32(mw.ArticleSalSecYear);
                    }
                    else if (dojTime.AddYears(3) >= fifteenthOfMonth)
                    {
                        return Convert.ToInt32(mw.ArticleSalThirdYear);
                    }
                }
                else
                {
                    if (dojTime.AddYears(1) >= fifteenthOfMonth)
                    {
                        return Convert.ToInt32(mw.ArticleFirstYear);
                    }
                    else if ((dojTime.AddYears(2) >= fifteenthOfMonth) && (dojTime.AddYears(3) >= fifteenthOfMonth))
                    {
                        return Convert.ToInt32(mw.ArticleSecondYear);
                    }
                    else if ((dojTime.AddYears(3) >= fifteenthOfMonth))
                    {
                        return Convert.ToInt32(mw.ArticleThirdYear);
                    }
                }
            }
            catch (Exception e)
            {
                throw e;
            }


            return 0;


        }

        public List<Employee> GetPayrollDetails(string ID = "", string subregParamYN = "N")
        {

            string strConnString = connectionString;
            string whereClause;
            if (ID == "")
            {
                whereClause = "";
            }
            else
            {
                if (subregParamYN == "Y")
                {
                    whereClause = " where id in(" + ID + ")";
                }
                else
                {
                    whereClause = " where id in(" + ID + ")";
                }

            }
            List<Employee> empList = new List<Employee>();
            try
            {
                NpgsqlConnection objConn = new NpgsqlConnection(strConnString);
                objConn.Open();
                string strSelectCmd = "select id,name,type,dateofjoin,coalesce(salary,0),coalesce(bonus, 0) ,coalesce(conveyance, 0) ,coalesce(status, ''),coalesce(pt, 0),coalesce(leavestakenthismnth, 0),coalesce(newleavesrepo, 0),coalesce(leavecredit, 0),coalesce(leavesrepo, 0)  from tbl_employee_sal_goutamtest " + whereClause + " order by name";

                NpgsqlCommand command = new NpgsqlCommand(strSelectCmd, objConn);
                command.CommandType = CommandType.Text;

                NpgsqlDataReader dr = command.ExecuteReader();

                while (dr.Read())
                {

                    empList.Add(
                                 new Employee
                                 {

                                     Id = Convert.ToInt32(dr[0]),

                                     Name = Convert.ToString(dr[1]),
                                     DataOfJoining = Convert.ToString(dr[3]),
                                     Type = Convert.ToString(dr[2]),
                                     Salary = Convert.ToInt32(dr[4]),
                                     Bonus = Convert.ToInt32(dr[5]),
                                     Conveyance = Convert.ToInt32(dr[6]),
                                     Status = Convert.ToString(dr[7]),
                                     PT = Convert.ToInt32(dr[8]),
                                     LeavesTakenThisMonth = Convert.ToDouble(dr[9]),
                                     NewLeavesRepo = Convert.ToDouble(dr[10]),
                                     LeaveCredit= Convert.ToDouble(dr[11]),
                                     Leaves = Convert.ToDouble(dr[12])


                                 }
                                     ); ;
                }

                NpgsqlConnection.ClearPool(objConn);
                objConn.Close();
                objConn.Dispose();
            }
            catch (Exception e)
            {
                throw e;
            }
            return empList;

        }

        public int getSalarybyName(string name, bool isminwage = false, int monthId = 0)
        {
            string strConnString = connectionString;
            string doj = "";
            int customsalary = 0;
            int takecustomsalaryyn = 0;
            try
            {
                NpgsqlConnection objConn = new NpgsqlConnection(strConnString);
                objConn.Open();
                string strSelectCmd = "select dateofjoin,customsalary,takecustomsalaryyn  from tbl_employee_goutamtest where name='" + name + "'";

                NpgsqlCommand command = new NpgsqlCommand(strSelectCmd, objConn);
                command.CommandType = CommandType.Text;

                NpgsqlDataReader dr = command.ExecuteReader();

                while (dr.Read())
                {

                    doj = Convert.ToString(dr[0]);
                    customsalary = Convert.ToInt32(dr[1]);
                    takecustomsalaryyn = Convert.ToInt32(dr[2]);

                }

                NpgsqlConnection.ClearPool(objConn);
                objConn.Close();
                objConn.Dispose();
            }
            catch (Exception e)
            {
                throw e;
            }

            return GetSalaryByJoiningDate(doj, isminwage, takecustomsalaryyn, customsalary, monthId);

        }


        public Employee calculateEmpSalOfTypeEmp(string name)
        {
            string strConnString = connectionString;

            Employee emp = new Employee();
            try
            {
                NpgsqlConnection objConn = new NpgsqlConnection(strConnString);
                objConn.Open();
                string strSelectCmd = "select coalesce(basicsal,0),coalesce(tds,0),coalesce(hra,0),coalesce(allowance,0) from tbl_employee_goutamtest where name='" + name + "'";

                NpgsqlCommand command = new NpgsqlCommand(strSelectCmd, objConn);
                command.CommandType = CommandType.Text;

                NpgsqlDataReader dr = command.ExecuteReader();

                while (dr.Read())
                {

                    emp.BasicSal = Convert.ToInt32(dr[0]);
                    emp.TDS = Convert.ToInt32(dr[1]);
                    emp.HRA = Convert.ToInt32(dr[2]);
                    emp.Allowance = Convert.ToInt32(dr[3]);

                }

                NpgsqlConnection.ClearPool(objConn);
                objConn.Close();
                objConn.Dispose();
            }
            catch (Exception e)
            {
                throw e;
            }



            return emp;


        }

        public List<Holiday> GetHolidayList(string ID = "", string subregParamYN = "N")
        {

            string strConnString = connectionString;
            string whereClause;
            if (ID == "")
            {
                whereClause = "";
            }
            else
            {
                if (subregParamYN == "Y")
                {
                    whereClause = " where id in(" + ID + ")";
                }
                else
                {
                    whereClause = " where id in(" + ID + ")";
                }

            }
            List<Holiday> holList = new List<Holiday>();
            try
            {
                NpgsqlConnection objConn = new NpgsqlConnection(strConnString);
                objConn.Open();
                string strSelectCmd = "select holidayon,holidayfor  from tbl_holidays_goutamtest " + whereClause + " order by holidayon ";

                NpgsqlCommand command = new NpgsqlCommand(strSelectCmd, objConn);
                command.CommandType = CommandType.Text;

                NpgsqlDataReader dr = command.ExecuteReader();

                while (dr.Read())
                {

                    holList.Add(
                                 new Holiday
                                 {

                                     HolidayFor = Convert.ToString(dr[1]),

                                     HolidayOn = Convert.ToString(dr[0])



                                 }
                                     ); ;
                }

                NpgsqlConnection.ClearPool(objConn);
                objConn.Close();
                objConn.Dispose();
            }
            catch (Exception e)
            {
                throw e;
            }
            return holList;
        }


        public bool DeleteHistory()
        {

            string strConnString = connectionString;
            APIResult result = new APIResult();

            NpgsqlConnection objConn = new NpgsqlConnection(strConnString);
            try
            {


                using (NpgsqlConnection connection = new NpgsqlConnection())
                {
                    connection.ConnectionString = strConnString;
                    connection.Open();
                    NpgsqlCommand cmd = new NpgsqlCommand();
                    cmd.Connection = connection;
                    cmd.CommandText = "truncate table tbl_employee_sal_goutamtest";
                    cmd.CommandType = CommandType.Text;


                    cmd.ExecuteNonQuery();
                    cmd.Dispose();
                    connection.Close();
                    NpgsqlConnection.ClearPool(connection);
                    result.Status = true;
                }

            }
            catch (Exception ex)
            {

                result.Status = false;
                throw ex;
            }
            finally
            {
                NpgsqlConnection.ClearPool(objConn);
                objConn.Close();
                objConn.Dispose();
            }




            return result.Status;
        }

        public MinumumWage GetFirmSettings()
        {
            string strConnString = connectionString;
            MinumumWage mw = new MinumumWage();

            try
            {
                NpgsqlConnection objConn = new NpgsqlConnection(strConnString);
                objConn.Open();
                string strSelectCmd = "select minimumwagearticleone,minimumwagearticletwo,minimumwagearticlethree,minimumwagearticleonesal,minimumwagearticletwosal,minimumwagearticlethreesal  from tbl_org_goutamtest where id=1";

                NpgsqlCommand command = new NpgsqlCommand(strSelectCmd, objConn);
                command.CommandType = CommandType.Text;

                NpgsqlDataReader dr = command.ExecuteReader();

                while (dr.Read())
                {
                    mw.ArticleFirstYear = Convert.ToString(dr[0]);
                    mw.ArticleSecondYear = Convert.ToString(dr[1]);
                    mw.ArticleThirdYear = Convert.ToString(dr[2]);
                    mw.ArticleSalFirstYear = Convert.ToString(dr[3]);
                    mw.ArticleSalSecYear = Convert.ToString(dr[4]);
                    mw.ArticleSalThirdYear = Convert.ToString(dr[5]);



                }

                NpgsqlConnection.ClearPool(objConn);
                objConn.Close();
                objConn.Dispose();
            }
            catch (Exception e)
            {
                throw e;
            }
            return mw;
        }

        public bool IsEmpArticle(string name)
        {
            string strConnString = connectionString;
            string type = "";
            try
            {
                NpgsqlConnection objConn = new NpgsqlConnection(strConnString);
                objConn.Open();

                string strSelectCmd = "select type   from tbl_employee_goutamtest where name='" + name + "'";

                NpgsqlCommand command = new NpgsqlCommand(strSelectCmd, objConn);
                command.CommandType = CommandType.Text;

                NpgsqlDataReader dr = command.ExecuteReader();

                while (dr.Read())
                {


                    type = Convert.ToString(dr[0]);




                }

                NpgsqlConnection.ClearPool(objConn);
                objConn.Close();
                objConn.Dispose();
            }
            catch (Exception e)
            {
                throw e;
            }


            return type.ToUpper() == "ARTICLE" ? true : false;
        }



        public bool GeneratePayroll(int monthId, bool manualpayrollYN = false, string manualfromdate = "", string manualtodate = "", int empid = 0)
        {

            Employee employee = new Employee();

            List<Holiday> holList = new List<Holiday>();

            List<string> holidayDateList = new List<string>();
            List<DateTime> holidayDateListDates = new List<DateTime>();
            List<string> monthDays = new List<string>();
            List<DateTime> workingdays = new List<DateTime>();
            List<DateTime> creditSalForDates = new List<DateTime>();
            DateTime toDate = new DateTime();
            DateTime fromDate = new DateTime();
            bool res = DeleteHistory();
            bool defresult = MakeDefaultPayrollEntry(monthId);
            Employee newjoineetracker = new Employee();


            if (!manualpayrollYN)
            {

                toDate = new DateTime(2022, monthId, 27);
                DateTime temp = new DateTime();
                temp = toDate.AddMonths(-1);
                fromDate = new DateTime(2022, temp.Month, 27);
            }
            else
            {


                toDate = Convert.ToDateTime(manualtodate);

                fromDate = Convert.ToDateTime(manualfromdate);
            }


            int numberOfDaysTotal = (toDate - fromDate).Days;


            var dates = Enumerable.Range(0, numberOfDaysTotal)
                        .Select(offset => fromDate.AddDays(offset))
                        .Where(date => date.DayOfWeek != DayOfWeek.Sunday);

            var saturdays = Enumerable.Range(0, numberOfDaysTotal)
                       .Select(offset => fromDate.AddDays(offset))
                       .Where(date => date.DayOfWeek == DayOfWeek.Saturday);

            holList = GetHolidayList();

            for (int i = 0; i < holList.Count; i++)
            {
                holidayDateList.Add(Convert.ToDateTime(holList[i].HolidayOn).ToShortDateString());
                holidayDateListDates.Add(Convert.ToDateTime(holList[i].HolidayOn));
            }



            for (int i = 0; i < dates.ToList().Count; i++)
            {
                monthDays.Add(Convert.ToDateTime(dates.ToList()[i]).ToShortDateString());
            }


            var aNotB = monthDays.Except(holidayDateList);

            aNotB.ToList().ForEach(x => workingdays.Add(Convert.ToDateTime(x)));

            if (!manualpayrollYN)
            {
                //specify the file name where its actually exist  
                List<Data> exceldataList = new List<Data>();
                string filepath = @"c:\Payroll\";


                foreach (string fileName in Directory.GetFiles(filepath))
                {
                    if (!fileName.Contains('$'))
                    {
                        Excel.Application oExcel = new Excel.Application();
                        Excel.Workbook WB = oExcel.Workbooks.Open(fileName);


                        // statement get the workbookname  
                        string ExcelWorkbookname = WB.Name;

                        // statement get the worksheet count  
                        int worksheetcount = WB.Worksheets.Count;

                        Excel.Worksheet wks = (Excel.Worksheet)WB.Worksheets[1];
                        int noofrowsinexcel = wks.UsedRange.Rows.Count - 1;
                        try
                        {
                            // statement get the firstworksheetname  

                            string firstworksheetname = wks.Name;
                            int cellwhichhasusername = 7;
                            int cellwhichhasdates = 1;
                            int cellwhichhasactivity = 3;


                            for (int i = 1; i < wks.UsedRange.Columns.Count; i++)
                            {

                                if (Convert.ToString(((Excel.Range)wks.Cells[1, i]).Value).Trim().ToUpper() == "USER")
                                {
                                    cellwhichhasusername = i;

                                }
                                else if (Convert.ToString(((Excel.Range)wks.Cells[1, i]).Value).Trim().ToUpper() == "DATE")
                                {
                                    cellwhichhasdates = i;
                                }
                                else if (Convert.ToString(((Excel.Range)wks.Cells[1, i]).Value).Trim().ToUpper() == "ACTIVITY")
                                {
                                    cellwhichhasactivity = i;
                                }

                            }



                            for (int i = 0; i < noofrowsinexcel; i++)
                            {
                                string name = "Error in name";
                                string activityname = "Error in activity name";
                                DateTime workingdates = new DateTime();

                                var datecellvalue = ((Excel.Range)wks.Cells[i + 2, cellwhichhasdates]).Value;
                                datecellvalue = Convert.ToString(datecellvalue);
                                if (datecellvalue != null && datecellvalue != "")
                                {
                                    if (datecellvalue.Contains("-"))
                                    {
                                        var arr = (datecellvalue.Trim()).Split('-');
                                        var datestring =( arr[1].Length == 1 ? "0" + arr[1] : arr[1] )+ "/" + (arr[0].Length == 1 ? "0" + arr[0] : arr[0] + "/" + arr[2]);
                                        workingdates = Convert.ToDateTime(datestring);
                                    }
                                    else
                                    {
                                        var arr = (datecellvalue.Trim()).Split('/');
                                        var datestring = (arr[1].Length == 1 ? "0" + arr[1] : arr[1]) + "/" + (arr[0].Length == 1 ? "0" + arr[0] : arr[0] + "/" + arr[2]);
                                        workingdates = Convert.ToDateTime(datestring);
                                    }

                                }

                                var namecellvalue = ((Excel.Range)wks.Cells[i + 2, cellwhichhasusername]).Value;
                                namecellvalue = Convert.ToString(namecellvalue);
                                if (namecellvalue != null && namecellvalue != "")
                                {
                                    name = namecellvalue;
                                }

                                var activitycellvalue = ((Excel.Range)wks.Cells[i + 2, cellwhichhasactivity]).Value;
                                activitycellvalue = Convert.ToString(activitycellvalue);
                                if (activitycellvalue != null && activitycellvalue != "")
                                {
                                    activityname = activitycellvalue.Trim();

                                    if (activityname.ToUpper() != "SUNDAY HOLIDAY" && activityname.ToUpper() != "LEAVE" && activityname.ToUpper() != "EXAMINATION LEAVE" && activityname.ToUpper() != "NATIONAL/OTHER HOLIDAY" && activityname.ToUpper() != "ICAI EXAMINATION LEAVE")
                                    {
                                        exceldataList.Add(
                                new Data
                                {

                                    Name = name,
                                    Activity = activityname,
                                    TimesheetDate = workingdates


                                }
                                    ); ;
                                    }

                                }



                            }

                            WB.Close();

                            oExcel.Quit();

                            Marshal.ReleaseComObject(wks);
                            Marshal.ReleaseComObject(WB);
                            Marshal.ReleaseComObject(oExcel);


                            //List<Tuple<String, DateTime>> mylist = new List<Tuple<String, DateTime>>();


                            Dictionary<string, List<(string, DateTime)>> DicData = new Dictionary<string, List<(string, DateTime)>>();
                            for (int i = 0; i < exceldataList.Count; i++)
                            {
                                var mylist = new List<(string activity, DateTime dates)>();
                                mylist.Add((exceldataList[i].Activity, exceldataList[i].TimesheetDate));
                                if (DicData.ContainsKey(exceldataList[i].Name))
                                {

                                    DicData[exceldataList[i].Name].Add(mylist.FirstOrDefault());
                                }
                                else
                                {

                                    DicData.Add(exceldataList[i].Name, mylist);
                                }


                            }


                            foreach (var dictionarydata in DicData)
                            {
                                List<DateTime> tempworkingdays = new List<DateTime>();
                                tempworkingdays = workingdays;
                                List<DateTime> dateList = new List<DateTime>();
                                List<DateTime> nomatterwhatworkeddateList = new List<DateTime>();
                                List<DateTime> excludedDateList = new List<DateTime>();
                                double numberOfHalfDays = 0;
                                foreach (var dateActivityForName in dictionarydata.Value)
                                {
                                    
                                    //var a = dateActivityForName.Item1;
                                    if (dateActivityForName.Item1.ToUpper() == "HALF DAY LEAVE" && !saturdays.ToList().Contains(dateActivityForName.Item2) )
                                    {
                                        excludedDateList.Add(dateActivityForName.Item2);
                                        numberOfHalfDays = numberOfHalfDays + 1;
                                    }
                                    else
                                    {
                                       
                                        dateList.Add(dateActivityForName.Item2);
                                    }


                                }
                                dateList = dateList.Distinct().ToList();
                                nomatterwhatworkeddateList = dateList;
                                dateList = dateList.Except(excludedDateList).ToList();

                                employee.Name = dictionarydata.Key;
                                bool isEmpArticle = IsEmpArticle(employee.Name);



                                var fdatelist = dateList.Except(holidayDateListDates);
                               

                                List<DateTime> FinaldateList = new List<DateTime>();

                                DateTime nextDate = fdatelist.ToList().Min(date => date);
                                while (nextDate <= fdatelist.ToList().Max(date => date))
                                {
                                    if (nextDate.DayOfWeek != DayOfWeek.Sunday && fdatelist.ToList().Contains(nextDate))
                                        FinaldateList.Add(nextDate);

                                    nextDate = nextDate.AddDays(1);
                                }


                                var datesWorkedOnHolidays = nomatterwhatworkeddateList.Except(workingdays);
                                
                                newjoineetracker = GetEmployeeDetails("", "N", employee.Name).FirstOrDefault();
                                if (newjoineetracker != null)
                                {
                                    bool checkifemployeejoinedinthismnth = false;
                                    checkifemployeejoinedinthismnth = (Convert.ToDateTime(newjoineetracker.DataOfJoining) >= fromDate && Convert.ToDateTime(newjoineetracker.DataOfJoining) < toDate);

                                    if (checkifemployeejoinedinthismnth)
                                    {
                                        int DaysTotal = (Convert.ToDateTime(newjoineetracker.DataOfJoining) - fromDate).Days;


                                        var tdates = Enumerable.Range(0, DaysTotal)
                                                    .Select(offset => fromDate.AddDays(offset))
                                                    .Where(date => date.DayOfWeek != DayOfWeek.Sunday);

                                        workingdays = workingdays.Except(tdates).ToList();
                                    }

                                    employee.LeavesTakenThisMonth = workingdays.Count -( FinaldateList.Count + (numberOfHalfDays / (double)2));
                                    employee.LeaveCredit = datesWorkedOnHolidays.ToList().Count;
                                   
                                    if (employee.Name.Contains("Venka"))
                                    {
                                        string a = "";
                                    }
                                    employee.Salary = masterSalaryCalculator(isEmpArticle, datesWorkedOnHolidays.ToList().Count, employee.Name, workingdays.Count, FinaldateList.Count + (numberOfHalfDays / (double)2), monthId, checkifemployeejoinedinthismnth);
                                    bool empexists = CheckIfEmpExists(employee.Name);
                                    employee.Status = empexists == true ? "Success" : "Error";
                                    MakeEntry(employee, isEmpArticle);
                                }
                                workingdays = tempworkingdays;
                            }



                        }
                        catch (Exception e)
                        {
                            WB.Close();

                            oExcel.Quit();

                            Marshal.ReleaseComObject(wks);
                            Marshal.ReleaseComObject(WB);
                            Marshal.ReleaseComObject(oExcel);

                            throw e;
                        }





                    }
                }
            }
            //else
            //{

            //    employee = (GetEmployeeDetails(Convert.ToString(empid))).FirstOrDefault();
            //    var fdatelist = dates.ToList().Except(holidayDateListDates);


            //    List<DateTime> FinaldateList = new List<DateTime>();

            //    DateTime nextDate = fdatelist.ToList().Min();
            //    while (nextDate <= fdatelist.ToList().Max())
            //    {
            //        if (nextDate.DayOfWeek != DayOfWeek.Sunday && fdatelist.ToList().Contains(nextDate))
            //            FinaldateList.Add(nextDate);

            //        nextDate = nextDate.AddDays(1);
            //    }


            //    var datesWorkedOnHolidays = dates.ToList().Except(FinaldateList);
            //    employee.Name = employee.Name;
            //    employee.Salary = dates.ToList().Count * calculateEmpSal(employee.Name, true, monthId) / (double)26;
            //    employee.Status = "Success";
            //    MakeEntry(employee,);
            //}

            return true;
        }






        public double masterSalaryCalculator(bool isEmpArticle, int datesWorkedOnHolidays, String empname, double workingdays, double FinaldateList, int monthId,bool joinedthismonth=false)
        {

            double salary = 0;
          
            
            int howmanydaysworkedonholiday = datesWorkedOnHolidays;
            double howmanydaysempabsent = 0;
            

            if (howmanydaysworkedonholiday > 0 && !isEmpArticle)
            {
                AddUptoEmpTotalLeaves(howmanydaysworkedonholiday, empname);
            }
            Employee empinstance = new Employee();
            if (!isEmpArticle)
            {
                empinstance = calculateEmpSalOfTypeEmp(empname);
            }


            if (workingdays > FinaldateList)
            {
                if (isEmpArticle)
                {
                    double sal = calculateEmpSal(empname, false, monthId);
                    if (!joinedthismonth)
                    {
                        salary = (26 - (workingdays - FinaldateList)) * (double)sal / 26 + (workingdays - FinaldateList) * calculateEmpSal(empname, true, monthId) / (double)26.00;
                    }
                    else
                    {
                        salary =   FinaldateList * (double)sal / 26 + (workingdays - FinaldateList) * calculateEmpSal(empname, true, monthId) / (double)26.00;
                    }
                   

                }
                else
                {
                    if ( (workingdays - FinaldateList) >= howmanydaysworkedonholiday)
                    {
                        howmanydaysempabsent = workingdays - FinaldateList - howmanydaysworkedonholiday;
                    }
                    else
                    {
                        howmanydaysempabsent = workingdays - FinaldateList;
                    }
                    
                    double goaheadwithdebit = DebitUptoEmpTotalLeaves(howmanydaysempabsent, empname);
                    //credint = (double)dateList.Count / (double)finalworkingdays;
                    if (goaheadwithdebit == 0)
                    {
                        if (!joinedthismonth)
                        {
                            salary = ((double)(empinstance.BasicSal + empinstance.HRA + empinstance.Allowance)) - empinstance.TDS;
                        }
                        else
                        {
                            salary = (((double)(empinstance.BasicSal + empinstance.HRA + empinstance.Allowance))* workingdays / 26) - empinstance.TDS; 
                        }
                    }
                    else
                    {

                        double sal = 0.00;
                        sal = (double)(empinstance.BasicSal + empinstance.HRA + empinstance.Allowance);
                        if (!joinedthismonth)
                        {
                            salary = ((26 - goaheadwithdebit) * sal / 26) - empinstance.TDS;
                        }
                        else
                        {
                            salary = ((workingdays - goaheadwithdebit) * sal / 26) - empinstance.TDS;
                        }

                    }
                }

            }
            else if (workingdays < FinaldateList)
            {
                if (isEmpArticle)
                {
                    salary = calculateEmpSal(empname, false, monthId);
                    //salary = calculateEmpSal(employee.Name, false, monthId) + (calculateEmpSal((employee.Name, false, monthId) * FinaldateList.Count / workingdays.Count) + (workingdays.Count - FinaldateList.Count) * calculateEmpSal(employee.Name, true, monthId);
                }
                else
                {
                    if (!joinedthismonth)
                    {
                        salary = ((double)(empinstance.BasicSal + empinstance.HRA + empinstance.Allowance)) - empinstance.TDS;
                    }
                    else
                    {
                        double sal = ((double)(empinstance.BasicSal + empinstance.HRA + empinstance.Allowance));
                        salary = (workingdays) * (double)sal / 26;
                    }
                }

            }
            else if (workingdays == FinaldateList)
            {
                if (isEmpArticle)
                {
                    if (!joinedthismonth)
                    {
                        salary = calculateEmpSal(empname, false, monthId);
                    }
                    else
                    {
                        double sal = calculateEmpSal(empname, false, monthId);
                        salary = (workingdays) * (double)sal / 26;
                    }
                }
                else
                {
                    if (!joinedthismonth)
                    {
                        salary = ((double)(empinstance.BasicSal + empinstance.HRA + empinstance.Allowance)) - empinstance.TDS;
                    }
                    else
                    {
                        double sal = ((double)(empinstance.BasicSal + empinstance.HRA + empinstance.Allowance));
                        salary = (workingdays) * (double)sal / 26;
                    }
                }

            }
            return salary;

        }
        public bool AddConveyance()
        {
            Excel.Application oExcel = new Excel.Application();

            Employee employee = new Employee();

            //specify the file name where its actually exist  
            string filepath = @"c:\Payroll\Conveyance";

            foreach (string fileName in Directory.GetFiles(filepath))
            {
                List<int> ConveyanceList = new List<int>();
                if (!fileName.Contains('$'))
                {

                    Excel.Workbook WB = oExcel.Workbooks.Open(fileName);


                    // statement get the workbookname  
                    string ExcelWorkbookname = WB.Name;

                    // statement get the worksheet count  
                    int worksheetcount = WB.Worksheets.Count;

                    Excel.Worksheet wks = (Excel.Worksheet)WB.Worksheets[1];
                    try
                    {
                        // statement get the firstworksheetname  

                        string firstworksheetname = wks.Name;
                        int cellwhichhasusername = 6;
                        int cellwhichhastotal = 1;

                        for (int i = 1; i < 11; i++)
                        {
                            if (Convert.ToString(((Excel.Range)wks.Cells[1, i]).Value) == "User")
                            {
                                cellwhichhasusername = i;
                            }

                        }


                        employee.Name = Convert.ToString(((Excel.Range)wks.Cells[2, cellwhichhasusername]).Value);

                        for (int i = 1; i < 11; i++)
                        {
                            if (Convert.ToString(((Excel.Range)wks.Cells[1, i]).Value) == "Total")
                            {
                                cellwhichhastotal = i;
                            }

                        }



                        int excelrows = wks.UsedRange.Rows.Count - 1;
                        Dictionary<string, List<int>> DicData = new Dictionary<string, List<int>>();
                        for (int i = 0; i < excelrows; i++)
                        {
                            var fcellvalue = ((Excel.Range)wks.Cells[i + 2, cellwhichhasusername]).Value;
                            string name = Convert.ToString(fcellvalue).Trim();

                            var fcelltotal = ((Excel.Range)wks.Cells[i + 2, cellwhichhastotal]).Value;
                            int celltotal = Convert.ToInt32(Convert.ToString(fcelltotal).Trim());

                            List<int> myList = new List<int>();
                            myList.Add(celltotal);

                            if (DicData.ContainsKey(name))
                            {

                                DicData[name].Add(myList.FirstOrDefault());
                            }
                            else
                            {

                                DicData.Add(name, myList);
                            }



                        }


                        foreach (var dictionarydata in DicData)
                        {
                            int total = 0;
                            foreach (var convce in dictionarydata.Value)
                            {
                                total = total + convce;


                            }

                            UpdateConveyance(total, dictionarydata.Key);
                        }



                    }
                    catch (Exception e)
                    {
                        WB.Close();

                        oExcel.Quit();

                        Marshal.ReleaseComObject(wks);
                        Marshal.ReleaseComObject(WB);
                        Marshal.ReleaseComObject(oExcel);

                        throw e;
                    }


                    WB.Close();

                    oExcel.Quit();

                    Marshal.ReleaseComObject(wks);
                    Marshal.ReleaseComObject(WB);
                    Marshal.ReleaseComObject(oExcel);


                }
            }

            return true;
        }

        public void UpdateConveyance(int ConveyanceTotal, string name)
        {
            string strConnString = connectionString;
            APIResult result = new APIResult();

            NpgsqlConnection objConn = new NpgsqlConnection(strConnString);
            try
            {


                using (NpgsqlConnection connection = new NpgsqlConnection())
                {
                    connection.ConnectionString = strConnString;
                    connection.Open();
                    NpgsqlCommand cmd = new NpgsqlCommand();
                    cmd.Connection = connection;
                    cmd.CommandText = "update tbl_employee_sal_goutamtest set conveyance=" + ConveyanceTotal + " where  name='" + name + "'";
                    cmd.CommandType = CommandType.Text;

                    //cmd.Parameters.Add(new NpgsqlParameter("@minimumwagearticleone", minwage.ArticleFirstYear));


                    cmd.ExecuteNonQuery();
                    cmd.Dispose();
                    connection.Close();
                    NpgsqlConnection.ClearPool(connection);
                    result.Status = true;
                }
            }
            catch (Exception ex)
            {

                result.Status = false;
                throw ex;
            }
            finally
            {
                NpgsqlConnection.ClearPool(objConn);
                objConn.Close();
                objConn.Dispose();
            }


        }


        public int calculateEmpSal(string name, bool isminwage = false, int monthId = 0)
        {
            return getSalarybyName(name, isminwage, monthId);
        }

        public void AddUptoEmpTotalLeaves(int workedonhol, string name)
        {
            string strConnString = connectionString;
            APIResult result = new APIResult();

            NpgsqlConnection objConn = new NpgsqlConnection(strConnString);
            try
            {


                using (NpgsqlConnection connection = new NpgsqlConnection())
                {
                    connection.ConnectionString = strConnString;
                    connection.Open();
                    NpgsqlCommand cmd = new NpgsqlCommand();
                    cmd.Connection = connection;
                    cmd.CommandText = "update tbl_employee_sal_goutamtest set  newleavesrepo=leavesrepo+" + workedonhol + " where  name='" + name + "'";
                    cmd.CommandType = CommandType.Text;

                    //cmd.Parameters.Add(new NpgsqlParameter("@minimumwagearticleone", minwage.ArticleFirstYear));


                    cmd.ExecuteNonQuery();
                    cmd.Dispose();
                    connection.Close();
                    NpgsqlConnection.ClearPool(connection);
                    result.Status = true;
                }
            }
            catch (Exception ex)
            {

                result.Status = false;
                throw ex;
            }
            finally
            {
                NpgsqlConnection.ClearPool(objConn);
                objConn.Close();
                objConn.Dispose();
            }

        }


        public double DebitUptoEmpTotalLeaves(double absentdays, string name)
        {
            //absentdays = absentdays + 1.25;


            string strConnString = connectionString;
            double leaves = 0.00;
            try
            {
                NpgsqlConnection objConn = new NpgsqlConnection(strConnString);
                objConn.Open();

                string strSelectCmd = "select leavesrepo   from tbl_employee_goutamtest where name='" + name + "'";

                NpgsqlCommand command = new NpgsqlCommand(strSelectCmd, objConn);
                command.CommandType = CommandType.Text;

                NpgsqlDataReader dr = command.ExecuteReader();

                while (dr.Read())
                {


                    leaves = Convert.ToDouble(dr[0]);




                }

                NpgsqlConnection.ClearPool(objConn);
                objConn.Close();
                objConn.Dispose();
            }
            catch (Exception e)
            {
                throw e;
            }
            if (leaves - absentdays > 0)
            {
                string strConnString2 = connectionString;
                APIResult result = new APIResult();

                NpgsqlConnection objConn = new NpgsqlConnection(strConnString2);
                try
                {


                    using (NpgsqlConnection connection = new NpgsqlConnection())
                    {
                        connection.ConnectionString = strConnString2;
                        connection.Open();
                        NpgsqlCommand cmd = new NpgsqlCommand();
                        cmd.Connection = connection;
                        cmd.CommandText = "update tbl_employee_sal_goutamtest set newleavesrepo=leavesrepo-" + absentdays + " where  name='" + name + "'";
                        cmd.CommandType = CommandType.Text;

                        //cmd.Parameters.Add(new NpgsqlParameter("@minimumwagearticleone", minwage.ArticleFirstYear));


                        cmd.ExecuteNonQuery();
                        cmd.Dispose();
                        connection.Close();
                        NpgsqlConnection.ClearPool(connection);
                        result.Status = true;
                    }
                }
                catch (Exception ex)
                {

                    result.Status = false;
                    throw ex;
                }
                finally
                {
                    NpgsqlConnection.ClearPool(objConn);
                    objConn.Close();
                    objConn.Dispose();
                }
                return 0;
            }
            else if (absentdays - leaves >= 0)
            {
                string strConnString2 = connectionString;
                APIResult result = new APIResult();

                NpgsqlConnection objConn = new NpgsqlConnection(strConnString2);
                try
                {


                    using (NpgsqlConnection connection = new NpgsqlConnection())
                    {
                        connection.ConnectionString = strConnString2;
                        connection.Open();
                        NpgsqlCommand cmd = new NpgsqlCommand();
                        cmd.Connection = connection;
                        cmd.CommandText = "update tbl_employee_sal_goutamtest set newleavesrepo=0 where name='" + name + "'";
                        cmd.CommandType = CommandType.Text;

                        //cmd.Parameters.Add(new NpgsqlParameter("@minimumwagearticleone", minwage.ArticleFirstYear));


                        cmd.ExecuteNonQuery();
                        cmd.Dispose();
                        connection.Close();
                        NpgsqlConnection.ClearPool(connection);
                        result.Status = true;
                    }
                }
                catch (Exception ex)
                {

                    result.Status = false;
                    throw ex;
                }
                finally
                {
                    NpgsqlConnection.ClearPool(objConn);
                    objConn.Close();
                    objConn.Dispose();
                }
                return absentdays - leaves;

            }
            return 0;

        }

        public bool CheckIfEmpExists(string name)
        {

            bool resut = false;
            string strConnString = connectionString;

            try
            {
                NpgsqlConnection objConn = new NpgsqlConnection(strConnString);
                objConn.Open();

                string strSelectCmd = "select 1   from tbl_employee_goutamtest where name='" + name + "'";

                NpgsqlCommand command = new NpgsqlCommand(strSelectCmd, objConn);
                command.CommandType = CommandType.Text;

                NpgsqlDataReader dr = command.ExecuteReader();

                while (dr.Read())
                {


                    resut = Convert.ToBoolean(dr[0]);




                }

                NpgsqlConnection.ClearPool(objConn);
                objConn.Close();
                objConn.Dispose();
            }
            catch (Exception e)
            {
                throw e;
            }


            return resut;

        }
        public IEnumerable<DateTime> EachDay(DateTime from, DateTime thru)
        {
            for (var day = from.Date; day.Date <= thru.Date; day = day.AddDays(1))
                yield return day;
        }

        public bool SaveSettings(MinumumWage minwage)
        {
            string strConnString = connectionString;
            APIResult result = new APIResult();

            NpgsqlConnection objConn = new NpgsqlConnection(strConnString);
            try
            {


                using (NpgsqlConnection connection = new NpgsqlConnection())
                {
                    connection.ConnectionString = strConnString;
                    connection.Open();
                    NpgsqlCommand cmd = new NpgsqlCommand();
                    cmd.Connection = connection;
                    cmd.CommandText = "update tbl_org_goutamtest set minimumwagearticleone=@minimumwagearticleone,minimumwagearticletwo=@minimumwagearticletwo,minimumwagearticlethree=@minimumwagearticlethree,minimumwagearticleoneSal=@minimumwagearticleoneSal,minimumwagearticletwoSal=@minimumwagearticletwoSal,minimumwagearticlethreeSal=@minimumwagearticlethreeSal where  id=1";
                    cmd.CommandType = CommandType.Text;

                    cmd.Parameters.Add(new NpgsqlParameter("@minimumwagearticleone", minwage.ArticleFirstYear));

                    cmd.Parameters.Add(new NpgsqlParameter("@minimumwagearticletwo", minwage.ArticleSecondYear));
                    cmd.Parameters.Add(new NpgsqlParameter("@minimumwagearticlethree", minwage.ArticleThirdYear));

                    cmd.Parameters.Add(new NpgsqlParameter("@minimumwagearticleoneSal", minwage.ArticleSalFirstYear));

                    cmd.Parameters.Add(new NpgsqlParameter("@minimumwagearticletwoSal", minwage.ArticleSalSecYear));
                    cmd.Parameters.Add(new NpgsqlParameter("@minimumwagearticlethreeSal", minwage.ArticleSalThirdYear));

                    cmd.ExecuteNonQuery();
                    cmd.Dispose();
                    connection.Close();
                    NpgsqlConnection.ClearPool(connection);
                    result.Status = true;
                }
            }
            catch (Exception ex)
            {

                result.Status = false;
                throw ex;
            }
            finally
            {
                NpgsqlConnection.ClearPool(objConn);
                objConn.Close();
                objConn.Dispose();
            }



            return true;
        }

      

    }
}
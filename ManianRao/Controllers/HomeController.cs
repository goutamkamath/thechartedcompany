using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;

using System.Web;
using System.Web.Mvc;
using TheChartedCompany.BL;
using TheChartedCompany.Models;


namespace ManianRao.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }
        public ActionResult Help()
        {

            return View();
        }

        public ActionResult Setting()
        {

            return View();
        }
        public ActionResult Holidays()
        {

            return View();
        }

        [HttpPost]
        public bool LoginUsers()
        {

            return true;

        }
        public ActionResult Dashboard()
        {

            return View();

        }
        public ActionResult Payroll()
        {
            BL blobj = new BL();
            List<Employee> emplist = new List<Employee>();
            emplist = blobj.GetOnlyEmployeeDetails("", "", true);
            Dictionary<int, string> empDic = new Dictionary<int, string>();
            empDic.Add(0, "Select Article");
            for (int i = 0; i < emplist.Count; i++)
            {
                empDic.Add(emplist[i].Id, emplist[i].Name);
            }


            ViewBag.ArticleList = empDic;


            return View();

        }
        public ActionResult AddEmployee()
        {

            return View();

        }


        [HttpGet]
        public int GetSalaryByJoiningDate(string doj)
        {
            BL blobj = new BL();
            int salary = 0;

            salary = blobj.GetSalaryByJoiningDate(doj);

            return salary;
        }

        [HttpPost]
        public bool AddNewEmployee(Employee emp)
        {
            BL blobj = new BL();
            bool result = false;
            result = blobj.AddNewEmployee(emp);

            return result;
        }

        [HttpPost]
        public bool creditfixedleaves(double leaves)
        {
            BL blobj = new BL();
            bool result = false;
            result = blobj.creditfixedleaves(leaves);

            return result;
        }

        [HttpPost]
        public bool SyncLeaves()
        {
            BL blobj = new BL();
            bool result = false;
            result = blobj.SyncLeaves();

            return result;
        }



        [HttpPost]
        public bool EditEmployee(Employee emp)
        {
            BL blobj = new BL();
            bool result = false;
            result = blobj.EditEmployee(emp);

            return result;
        }

        [HttpPost]
        public bool SaveLeaves(string name, double leavesrepo)
        {
            BL blobj = new BL();
            bool result = false;
            result = blobj.SaveLeaves(name, leavesrepo);

            return result;
        }


        [HttpPost]
        public bool AddHolidays(List<Holiday> hol)
        {
            BL blobj = new BL();
            bool result = false;

            result = blobj.AddNewHoliday(hol);

            return result;
        }
        [HttpPost]
        public bool EditHoliday(Holiday hol)
        {
            BL blobj = new BL();
            bool result = false;

            result = blobj.EditHoliday(hol);

            return result;
        }


        [HttpPost]
        public bool EditPayroll(Employee emp)
        {
            BL blobj = new BL();
            bool result = false;

            result = blobj.EditPayroll(emp);

            return result;
        }

        [HttpPost]
        public ActionResult UploadExcel(FileUpload membervalues)
        {
            //Use Namespace called :  System.IO  
            string FileName = Path.GetFileNameWithoutExtension(membervalues.ImageFile.FileName);

            //To Get File Extension  
            string FileExtension = Path.GetExtension(membervalues.ImageFile.FileName);

            //Add Current Date To Attached File Name  
            FileName = DateTime.Now.ToString("yyyyMMdd") + "-" + FileName.Trim() + FileExtension;

            //Get Upload path from Web.Config file AppSettings.  
            string UploadPath = ConfigurationManager.AppSettings["UserImagePath"].ToString();

            //Its Create complete path to store in server.  
            membervalues.ImagePath = UploadPath + FileName;

            //To copy and save file into server.  
            membervalues.ImageFile.SaveAs(membervalues.ImagePath);


            //To save Club Member Contact Form Detail to database table.  


            return View();
        }




        [HttpPost]
        public bool GeneratePayroll(int monthId)
        {
           
            BL blobj = new BL();

            return blobj.GeneratePayroll(monthId, false, "", "", 0);
        }
        [HttpPost]
        public bool AddConveyance()
        {
            BL blobj = new BL();

            return blobj.AddConveyance();
        }
        [HttpPost]
        public bool DeleteLogs()
        {
            BL blobj = new BL();

            return blobj.DeleteHistory();
        }




        [HttpGet]
        public JsonResult GetHolidayList(string ID = "", string subregParamYN = "N")
        {

            BL blobj = new BL();


            List<Holiday> holList = blobj.GetHolidayDetails(ID);
            return Json(holList, JsonRequestBehavior.AllowGet);

        }

        [HttpGet]
        public JsonResult GetEmployeeList(string ID = "", string subregParamYN = "N")
        {

            BL blobj = new BL();


            List<Employee> empList = blobj.GetEmployeeDetails(ID);
            return Json(empList, JsonRequestBehavior.AllowGet);

        }

        [HttpGet]
        public JsonResult GetOnlyEmployeeList(string ID = "", string subregParamYN = "N", bool isEmpArticle = false)
        {

            BL blobj = new BL();


            List<Employee> empList = blobj.GetOnlyEmployeeDetails(ID, "N", isEmpArticle);
            return Json(empList, JsonRequestBehavior.AllowGet);

        }




        [HttpGet]
        public JsonResult GetPayrollList(string ID = "", string subregParamYN = "N")
        {

            BL blobj = new BL();


            List<Employee> empList = blobj.GetPayrollDetails(ID);
            return Json(empList, JsonRequestBehavior.AllowGet);

        }

        [HttpPost]
        public JsonResult SaveSettings(MinumumWage mw)
        {

            BL blobj = new BL();
            bool res = false;

            res = blobj.SaveSettings(mw);
            return Json(res, JsonRequestBehavior.AllowGet);

        }
        [HttpPost]
        public bool DeleteEmployeeList(string ID)
        {

            BL blobj = new BL();
            bool res = false;

            res = blobj.DeleteEmployee(ID);
            return res;

        }

        [HttpPost]
        public bool DeleteHoliday(string ID)
        {

            BL blobj = new BL();
            bool res = false;

            res = blobj.DeleteHoliday(ID);
            return res;

        }

        public ActionResult LeaveManagement()
        {
            return View();
        }



        [HttpGet]
        public JsonResult GetFirmSettings()
        {

            BL blobj = new BL();

            MinumumWage mw = new MinumumWage();
            mw = blobj.GetFirmSettings();
            return Json(mw, JsonRequestBehavior.AllowGet);

        }



    }


}
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Web;
using System.Web.Mvc;

namespace TheChartedCompany.Models
{
    public class FileUpload
    {
        [DisplayName("Upload Timesheet")]
        public string ImagePath { get; set; }

        public HttpPostedFileBase ImageFile { get; set; }

        

        public List<SelectListItem> PayMonths { get; set; }

    }

}
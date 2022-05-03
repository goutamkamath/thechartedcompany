using System.Collections.Generic;

namespace TheChartedCompany.Models
{
    public class APIResult
    {
        public APIResult()
        {
            Messages = new List<string>();
        }

        public bool Status { get; set; }
        public List<string> Messages { get; set; }
        public dynamic Data { get; set; }
    }
}

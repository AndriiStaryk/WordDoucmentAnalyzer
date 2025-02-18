using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocumentAnalyzer;

struct DailyTask
{
    public string TaskName { get; set; }
    public DateTime StartDate { get; set; }
    public DateTime EndDate { get; set; }

    public DailyTask(string taskName = "", DateTime? startDate = null, DateTime? endDate = null)
    {
        TaskName = taskName;
        StartDate = startDate ?? DateTime.Today;
        EndDate = endDate ?? DateTime.Today;
    }

    public List<string> ToStringList()
    {
        return new List<string> { TaskName, StartDate.ToShortDateString(), EndDate.ToShortDateString() };
    }
}

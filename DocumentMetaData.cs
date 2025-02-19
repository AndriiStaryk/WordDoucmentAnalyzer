using DocumentFormat.OpenXml.Presentation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocumentAnalyzer;

internal struct DocumentMetaData
{
    public string NominativeCaseName { get; set; }
    public string GenitiveCaseName { get; set; }

    public Gender Gender { get; set; }

    public DateTime StartDate { get; set; }
    public DateTime EndDate { get; set; }

    public string PracticePlace { get; set; }
    public string Group {  get; set; }

    public string MentorsFromDepartment { get; set; }

    public string MentorsFromFaculty { get; set; }

    public string TaskDescription {  get; set; }

    public List<DailyTask> DailyTasks { get; set; }
   
    public string Characteristics { get; set; }
}

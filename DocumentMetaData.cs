using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocumentAnalyzer;

public enum Gender
{
    Male,
    Female
}
internal class DocumentMetaData
{
    public string NominativeCaseName { get; set; }
    public string GenitiveCaseName { get; set; }

    public Gender Gender { get; set; }

    public DateTime StartDate { get; set; }
    public DateTime EndDate { get; set; }

    public string PracticePlace { get; set; }
    public string Group {  get; set; }

    public string Mentors { get; set; }

    public string MainMentor { get; set; }
   
}

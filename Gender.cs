using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace DocumentAnalyzer;

public static class AttributesHelperExtension
{
    public static string ToDescription(this Gender gender, bool genitive = false)
    {
        var attribute = (DescriptionAttribute[])gender
            .GetType()
            .GetField(gender.ToString())
            .GetCustomAttributes(typeof(DescriptionAttribute), false);

        if (attribute.Length == 0) return gender.ToString();

        var values = attribute[0].Description.Split('|');
        return genitive ? values[1] : values[0];
    }
}

public enum Gender
{
    [Description("Студент|Студента")]
    Male,
    [Description("Студентка|Студентки")]
    Female
}
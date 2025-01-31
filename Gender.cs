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
    public static string ToDescription(this Enum value)
    {
        var da = (DescriptionAttribute[])(value.GetType().GetField(value.ToString())).GetCustomAttributes(typeof(DescriptionAttribute), false);
        return da.Length > 0 ? da[0].Description : value.ToString();
    }
}

public enum Gender
{
    [Description("Студент")]
    Male,
    [Description("Студентка")]
    Female
}
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DocumentAnalyzer;

public partial class MetaDataForm : Form
{
    DocxManager _docxManager = new DocxManager();
    public MetaDataForm()
    {
        InitializeComponent();
       
        foreach (Gender gender in Enum.GetValues(typeof(Gender)))
        {
            comboBox_Gender.Items.Add(gender.ToDescription());
        }

        comboBox_Gender.SelectedIndex = 0;
    }

    private void button_GenerateDoc_Click(object sender, EventArgs e)
    {
        var docMetaData = new DocumentMetaData
        {
            NominativeCaseName = textBox_NominativeCaseName.Text,
            GenitiveCaseName = textBox_GenitiveCaseName.Text,
            Gender = (Gender)comboBox_Gender.SelectedItem, 
            StartDate = dateTimePicker_StartDate.Value,
            EndDate = dateTimePicker_EndDate.Value,
            PracticePlace = textBox_PracticePlace.Text,
            Group = textBox_Group.Text,
            MentorsFromDepartment = textBox_MentorsDepartment.Text,
            MentorsFromFaculty = textBox_MentorsFaculty.Text
        };

        _docxManager.GenerateDocument(docMetaData);
    }
}

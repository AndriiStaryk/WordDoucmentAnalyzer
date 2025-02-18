using DocumentFormat.OpenXml.Office.CustomUI;
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
    private DocxManager _docxManager = new DocxManager();
    private List<DailyTask> _dailyTasks = new List<DailyTask>();
    public MetaDataForm()
    {
        InitializeComponent();

        foreach (Gender gender in Enum.GetValues(typeof(Gender)))
        {
            comboBox_Gender.Items.Add(gender.ToDescription());
        }

        comboBox_Gender.SelectedIndex = 0;


        DataGridViewTextBoxColumn nameColumn = new DataGridViewTextBoxColumn();
        nameColumn.HeaderText = "Опис діяльності";
        nameColumn.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
        nameColumn.FillWeight = 50;
        dataGridView_DailyTasksTable.Columns.Add(nameColumn);

        DataGridViewCalendarColumn dateStartColumn = new DataGridViewCalendarColumn();
        dateStartColumn.HeaderText = "Дата початку";
        dateStartColumn.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
        dateStartColumn.FillWeight = 20;
        dataGridView_DailyTasksTable.Columns.Add(dateStartColumn);

        DataGridViewCalendarColumn dateEndColumn = new DataGridViewCalendarColumn();
        dateEndColumn.HeaderText = "Дата закінчення";
        dateEndColumn.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
        dateEndColumn.FillWeight = 20;
        dataGridView_DailyTasksTable.Columns.Add(dateEndColumn);


        dataGridView_DailyTasksTable.RowsAdded += DataGridView_DailyTasksTable_RowsAdded;
        dataGridView_DailyTasksTable.CellValueChanged += DataGridView_DailyTasksTable_CellValueChanged;
        dataGridView_DailyTasksTable.RowsRemoved += DataGridView_DailyTasksTable_RowsRemoved;

        //Test Data
        textBox_NominativeCaseName.Text = "Старик Андрій Сергійович";
        textBox_GenitiveCaseName.Text = "Старика Андрія Сергійовича";
        textBox_Group.Text = "ТТП-32";
        textBox_PracticePlace.Text = "факультет комп'ютерних наук та кібернетики";
        richTextBox_MentorsDepartment.Text = "Доц.Зубенко В.В., ас. Свистунов А.О., ас. Шишацький А.В., ас. Галавай О.М., ас. Пушкаренко Ю.В.";
        richTextBox_MentorsFaculty.Text = "зав. декана Омельчук Людмила Леонідівна";
        dateTimePicker_StartDate.Value = new DateTime(2025, 1, 20);
        dateTimePicker_EndDate.Value = new DateTime(2025, 2, 28);
    }

    private void button_GenerateDoc_Click(object sender, EventArgs e)
    {
        string selectedString = comboBox_Gender.SelectedItem.ToString();

        Gender selectedGender = Enum.GetValues(typeof(Gender))
                                    .Cast<Gender>()
                                    .FirstOrDefault(g => g.ToDescription() == selectedString);

        var docMetaData = new DocumentMetaData
        {
            NominativeCaseName = textBox_NominativeCaseName.Text,
            GenitiveCaseName = textBox_GenitiveCaseName.Text,
            Gender = selectedGender,
            StartDate = dateTimePicker_StartDate.Value,
            EndDate = dateTimePicker_EndDate.Value,
            PracticePlace = textBox_PracticePlace.Text,
            Group = textBox_Group.Text,
            MentorsFromDepartment = richTextBox_MentorsDepartment.Text,
            MentorsFromFaculty = richTextBox_MentorsFaculty.Text,
            TaskDescription = richTextBox_TaskDescription.Text,
            DailyTasks = _dailyTasks
        };

        _docxManager.GenerateDocument(docMetaData);
    }


    private void DataGridView_DailyTasksTable_CellValueChanged(object sender, DataGridViewCellEventArgs e)
    {
        if (e.RowIndex >= 0 && e.RowIndex < _dailyTasks.Count)
        {
            var task = _dailyTasks[e.RowIndex];

            if (e.ColumnIndex == 0) 
            {
                task.TaskName = dataGridView_DailyTasksTable.Rows[e.RowIndex].Cells[e.ColumnIndex].Value?.ToString();
            }
            else if (e.ColumnIndex == 1)
            {
                task.StartDate = Convert.ToDateTime(dataGridView_DailyTasksTable.Rows[e.RowIndex].Cells[e.ColumnIndex].Value);
            }
            else if (e.ColumnIndex == 2)
            {
                task.EndDate = Convert.ToDateTime(dataGridView_DailyTasksTable.Rows[e.RowIndex].Cells[e.ColumnIndex].Value);
            }

            _dailyTasks[e.RowIndex] = task;
        }
    }

    private void DataGridView_DailyTasksTable_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
    {
        for (int i = 0; i < e.RowCount; i++)
        {
            _dailyTasks.Add(new DailyTask(""));
        }
    }

    private void DataGridView_DailyTasksTable_RowsRemoved(object sender, DataGridViewRowsRemovedEventArgs e)
    {
        for (int i = 0; i < e.RowCount; i++)
        {
            if (e.RowIndex < _dailyTasks.Count)
            {
                _dailyTasks.RemoveAt(e.RowIndex);
            }
        }
    }

}

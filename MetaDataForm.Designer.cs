namespace DocumentAnalyzer
{
    partial class MetaDataForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            textBox_NominativeCaseName = new TextBox();
            textBox_GenitiveCaseName = new TextBox();
            comboBox_Gender = new ComboBox();
            dateTimePicker_StartDate = new DateTimePicker();
            dateTimePicker_EndDate = new DateTimePicker();
            label1 = new Label();
            label2 = new Label();
            label3 = new Label();
            label4 = new Label();
            textBox_PracticePlace = new TextBox();
            textBox_Group = new TextBox();
            label5 = new Label();
            label6 = new Label();
            label7 = new Label();
            label8 = new Label();
            button_GenerateDoc = new Button();
            richTextBox_MentorsDepartment = new RichTextBox();
            richTextBox_MentorsFaculty = new RichTextBox();
            panel1 = new Panel();
            label11 = new Label();
            richTextBox_studentsСharacteristic = new RichTextBox();
            label10 = new Label();
            dataGridView_DailyTasksTable = new DataGridView();
            richTextBox_TaskDescription = new RichTextBox();
            label9 = new Label();
            panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)dataGridView_DailyTasksTable).BeginInit();
            SuspendLayout();
            // 
            // textBox_NominativeCaseName
            // 
            textBox_NominativeCaseName.Location = new Point(226, 29);
            textBox_NominativeCaseName.Name = "textBox_NominativeCaseName";
            textBox_NominativeCaseName.Size = new Size(646, 47);
            textBox_NominativeCaseName.TabIndex = 0;
            // 
            // textBox_GenitiveCaseName
            // 
            textBox_GenitiveCaseName.Location = new Point(226, 141);
            textBox_GenitiveCaseName.Name = "textBox_GenitiveCaseName";
            textBox_GenitiveCaseName.Size = new Size(646, 47);
            textBox_GenitiveCaseName.TabIndex = 1;
            // 
            // comboBox_Gender
            // 
            comboBox_Gender.FormattingEnabled = true;
            comboBox_Gender.Location = new Point(25, 243);
            comboBox_Gender.Name = "comboBox_Gender";
            comboBox_Gender.Size = new Size(302, 49);
            comboBox_Gender.TabIndex = 2;
            // 
            // dateTimePicker_StartDate
            // 
            dateTimePicker_StartDate.Location = new Point(65, 355);
            dateTimePicker_StartDate.Name = "dateTimePicker_StartDate";
            dateTimePicker_StartDate.Size = new Size(500, 47);
            dateTimePicker_StartDate.TabIndex = 3;
            // 
            // dateTimePicker_EndDate
            // 
            dateTimePicker_EndDate.Location = new Point(682, 356);
            dateTimePicker_EndDate.Name = "dateTimePicker_EndDate";
            dateTimePicker_EndDate.Size = new Size(500, 47);
            dateTimePicker_EndDate.TabIndex = 4;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new Point(25, 29);
            label1.Name = "label1";
            label1.Size = new Size(171, 41);
            label1.TabIndex = 5;
            label1.Text = "П.І.Б. (Н. В.)";
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Location = new Point(25, 147);
            label2.Name = "label2";
            label2.Size = new Size(167, 41);
            label2.TabIndex = 6;
            label2.Text = "П.І.Б. (Р. В.)";
            // 
            // label3
            // 
            label3.AutoSize = true;
            label3.Location = new Point(25, 355);
            label3.Name = "label3";
            label3.Size = new Size(34, 41);
            label3.TabIndex = 7;
            label3.Text = "З";
            // 
            // label4
            // 
            label4.AutoSize = true;
            label4.Location = new Point(605, 361);
            label4.Name = "label4";
            label4.Size = new Size(53, 41);
            label4.TabIndex = 8;
            label4.Text = "по";
            // 
            // textBox_PracticePlace
            // 
            textBox_PracticePlace.Location = new Point(319, 460);
            textBox_PracticePlace.Name = "textBox_PracticePlace";
            textBox_PracticePlace.Size = new Size(863, 47);
            textBox_PracticePlace.TabIndex = 9;
            // 
            // textBox_Group
            // 
            textBox_Group.Location = new Point(622, 243);
            textBox_Group.Name = "textBox_Group";
            textBox_Group.Size = new Size(250, 47);
            textBox_Group.TabIndex = 10;
            // 
            // label5
            // 
            label5.AutoSize = true;
            label5.Location = new Point(25, 463);
            label5.Name = "label5";
            label5.Size = new Size(274, 41);
            label5.TabIndex = 13;
            label5.Text = "Місце проведення";
            // 
            // label6
            // 
            label6.AutoSize = true;
            label6.Location = new Point(508, 246);
            label6.Name = "label6";
            label6.Size = new Size(97, 41);
            label6.TabIndex = 14;
            label6.Text = "Група";
            // 
            // label7
            // 
            label7.AutoSize = true;
            label7.Location = new Point(25, 618);
            label7.Name = "label7";
            label7.Size = new Size(633, 41);
            label7.TabIndex = 15;
            label7.Text = "Керівники практики (кафедра/підприємство)";
            // 
            // label8
            // 
            label8.AutoSize = true;
            label8.Location = new Point(25, 846);
            label8.Name = "label8";
            label8.Size = new Size(455, 41);
            label8.TabIndex = 16;
            label8.Text = "Керівники практики (факультет)";
            // 
            // button_GenerateDoc
            // 
            button_GenerateDoc.Location = new Point(450, 1978);
            button_GenerateDoc.Name = "button_GenerateDoc";
            button_GenerateDoc.Size = new Size(486, 58);
            button_GenerateDoc.TabIndex = 17;
            button_GenerateDoc.Text = "Згенерувати щоденник";
            button_GenerateDoc.UseVisualStyleBackColor = true;
            button_GenerateDoc.Click += button_GenerateDoc_Click;
            // 
            // richTextBox_MentorsDepartment
            // 
            richTextBox_MentorsDepartment.Location = new Point(682, 618);
            richTextBox_MentorsDepartment.Name = "richTextBox_MentorsDepartment";
            richTextBox_MentorsDepartment.Size = new Size(612, 178);
            richTextBox_MentorsDepartment.TabIndex = 18;
            richTextBox_MentorsDepartment.Text = "";
            // 
            // richTextBox_MentorsFaculty
            // 
            richTextBox_MentorsFaculty.Location = new Point(682, 846);
            richTextBox_MentorsFaculty.Name = "richTextBox_MentorsFaculty";
            richTextBox_MentorsFaculty.Size = new Size(612, 174);
            richTextBox_MentorsFaculty.TabIndex = 19;
            richTextBox_MentorsFaculty.Text = "";
            // 
            // panel1
            // 
            panel1.AutoScroll = true;
            panel1.Controls.Add(label11);
            panel1.Controls.Add(richTextBox_studentsСharacteristic);
            panel1.Controls.Add(label10);
            panel1.Controls.Add(dataGridView_DailyTasksTable);
            panel1.Controls.Add(richTextBox_TaskDescription);
            panel1.Controls.Add(label9);
            panel1.Controls.Add(textBox_NominativeCaseName);
            panel1.Controls.Add(button_GenerateDoc);
            panel1.Controls.Add(richTextBox_MentorsFaculty);
            panel1.Controls.Add(textBox_GenitiveCaseName);
            panel1.Controls.Add(richTextBox_MentorsDepartment);
            panel1.Controls.Add(comboBox_Gender);
            panel1.Controls.Add(dateTimePicker_StartDate);
            panel1.Controls.Add(label8);
            panel1.Controls.Add(dateTimePicker_EndDate);
            panel1.Controls.Add(label7);
            panel1.Controls.Add(label1);
            panel1.Controls.Add(label6);
            panel1.Controls.Add(label2);
            panel1.Controls.Add(label5);
            panel1.Controls.Add(label3);
            panel1.Controls.Add(textBox_Group);
            panel1.Controls.Add(label4);
            panel1.Controls.Add(textBox_PracticePlace);
            panel1.Dock = DockStyle.Fill;
            panel1.Location = new Point(0, 0);
            panel1.Name = "panel1";
            panel1.Size = new Size(1489, 941);
            panel1.TabIndex = 20;
            // 
            // label11
            // 
            label11.AutoSize = true;
            label11.Location = new Point(32, 1346);
            label11.Name = "label11";
            label11.Size = new Size(235, 41);
            label11.TabIndex = 25;
            label11.Text = "Характеристика";
            // 
            // richTextBox_studentsСharacteristic
            // 
            richTextBox_studentsСharacteristic.Location = new Point(682, 1346);
            richTextBox_studentsСharacteristic.Name = "richTextBox_studentsСharacteristic";
            richTextBox_studentsСharacteristic.Size = new Size(612, 240);
            richTextBox_studentsСharacteristic.TabIndex = 24;
            richTextBox_studentsСharacteristic.Text = "";
            // 
            // label10
            // 
            label10.AutoSize = true;
            label10.Location = new Point(645, 2058);
            label10.Name = "label10";
            label10.Size = new Size(82, 41);
            label10.TabIndex = 23;
            label10.Text = "2025";
            label10.Click += label10_Click;
            // 
            // dataGridView_DailyTasksTable
            // 
            dataGridView_DailyTasksTable.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridView_DailyTasksTable.Location = new Point(32, 1608);
            dataGridView_DailyTasksTable.Name = "dataGridView_DailyTasksTable";
            dataGridView_DailyTasksTable.RowHeadersWidth = 102;
            dataGridView_DailyTasksTable.Size = new Size(1262, 319);
            dataGridView_DailyTasksTable.TabIndex = 22;
            // 
            // richTextBox_TaskDescription
            // 
            richTextBox_TaskDescription.Location = new Point(682, 1046);
            richTextBox_TaskDescription.Name = "richTextBox_TaskDescription";
            richTextBox_TaskDescription.Size = new Size(612, 266);
            richTextBox_TaskDescription.TabIndex = 21;
            richTextBox_TaskDescription.Text = "";
            // 
            // label9
            // 
            label9.AutoSize = true;
            label9.Location = new Point(25, 1049);
            label9.Name = "label9";
            label9.Size = new Size(317, 41);
            label9.TabIndex = 20;
            label9.Text = "Завдання на практику";
            // 
            // MetaDataForm
            // 
            AutoScaleDimensions = new SizeF(17F, 41F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(1489, 941);
            Controls.Add(panel1);
            Name = "MetaDataForm";
            Text = "MetaDataForm";
            panel1.ResumeLayout(false);
            panel1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)dataGridView_DailyTasksTable).EndInit();
            ResumeLayout(false);
        }

        #endregion

        private TextBox textBox_NominativeCaseName;
        private TextBox textBox_GenitiveCaseName;
        private ComboBox comboBox_Gender;
        private DateTimePicker dateTimePicker_StartDate;
        private DateTimePicker dateTimePicker_EndDate;
        private Label label1;
        private Label label2;
        private Label label3;
        private Label label4;
        private TextBox textBox_PracticePlace;
        private TextBox textBox_Group;
        private Label label5;
        private Label label6;
        private Label label7;
        private Label label8;
        private Button button_GenerateDoc;
        private RichTextBox richTextBox_MentorsDepartment;
        private RichTextBox richTextBox_MentorsFaculty;
        private Panel panel1;
        private RichTextBox richTextBox_TaskDescription;
        private Label label9;
        private DataGridView dataGridView_DailyTasksTable;
        private Label label10;
        private Label label11;
        private RichTextBox richTextBox_studentsСharacteristic;
    }
}
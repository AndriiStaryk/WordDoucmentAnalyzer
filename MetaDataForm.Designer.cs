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
            textBox_MentorsDepartment = new TextBox();
            textBox_MentorsFaculty = new TextBox();
            label5 = new Label();
            label6 = new Label();
            label7 = new Label();
            label8 = new Label();
            button_GenerateDoc = new Button();
            SuspendLayout();
            // 
            // textBox_NominativeCaseName
            // 
            textBox_NominativeCaseName.Location = new Point(264, 98);
            textBox_NominativeCaseName.Name = "textBox_NominativeCaseName";
            textBox_NominativeCaseName.Size = new Size(250, 47);
            textBox_NominativeCaseName.TabIndex = 0;
            // 
            // textBox_GenitiveCaseName
            // 
            textBox_GenitiveCaseName.Location = new Point(264, 210);
            textBox_GenitiveCaseName.Name = "textBox_GenitiveCaseName";
            textBox_GenitiveCaseName.Size = new Size(250, 47);
            textBox_GenitiveCaseName.TabIndex = 1;
            // 
            // comboBox_Gender
            // 
            comboBox_Gender.FormattingEnabled = true;
            comboBox_Gender.Location = new Point(184, 305);
            comboBox_Gender.Name = "comboBox_Gender";
            comboBox_Gender.Size = new Size(302, 49);
            comboBox_Gender.TabIndex = 2;
            // 
            // dateTimePicker_StartDate
            // 
            dateTimePicker_StartDate.Location = new Point(184, 424);
            dateTimePicker_StartDate.Name = "dateTimePicker_StartDate";
            dateTimePicker_StartDate.Size = new Size(500, 47);
            dateTimePicker_StartDate.TabIndex = 3;
            // 
            // dateTimePicker_EndDate
            // 
            dateTimePicker_EndDate.Location = new Point(820, 434);
            dateTimePicker_EndDate.Name = "dateTimePicker_EndDate";
            dateTimePicker_EndDate.Size = new Size(500, 47);
            dateTimePicker_EndDate.TabIndex = 4;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new Point(63, 98);
            label1.Name = "label1";
            label1.Size = new Size(171, 41);
            label1.TabIndex = 5;
            label1.Text = "П.І.Б. (Н. В.)";
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Location = new Point(63, 216);
            label2.Name = "label2";
            label2.Size = new Size(167, 41);
            label2.TabIndex = 6;
            label2.Text = "П.І.Б. (Р. В.)";
            // 
            // label3
            // 
            label3.AutoSize = true;
            label3.Location = new Point(63, 424);
            label3.Name = "label3";
            label3.Size = new Size(34, 41);
            label3.TabIndex = 7;
            label3.Text = "З";
            // 
            // label4
            // 
            label4.AutoSize = true;
            label4.Location = new Point(729, 424);
            label4.Name = "label4";
            label4.Size = new Size(53, 41);
            label4.TabIndex = 8;
            label4.Text = "по";
            // 
            // textBox_PracticePlace
            // 
            textBox_PracticePlace.Location = new Point(372, 529);
            textBox_PracticePlace.Name = "textBox_PracticePlace";
            textBox_PracticePlace.Size = new Size(250, 47);
            textBox_PracticePlace.TabIndex = 9;
            // 
            // textBox_Group
            // 
            textBox_Group.Location = new Point(184, 647);
            textBox_Group.Name = "textBox_Group";
            textBox_Group.Size = new Size(250, 47);
            textBox_Group.TabIndex = 10;
            // 
            // textBox_MentorsDepartment
            // 
            textBox_MentorsDepartment.Location = new Point(775, 760);
            textBox_MentorsDepartment.Name = "textBox_MentorsDepartment";
            textBox_MentorsDepartment.Size = new Size(250, 47);
            textBox_MentorsDepartment.TabIndex = 11;
            // 
            // textBox_MentorsFaculty
            // 
            textBox_MentorsFaculty.Location = new Point(578, 865);
            textBox_MentorsFaculty.Name = "textBox_MentorsFaculty";
            textBox_MentorsFaculty.Size = new Size(250, 47);
            textBox_MentorsFaculty.TabIndex = 12;
            // 
            // label5
            // 
            label5.AutoSize = true;
            label5.Location = new Point(63, 532);
            label5.Name = "label5";
            label5.Size = new Size(274, 41);
            label5.TabIndex = 13;
            label5.Text = "Місце проведення";
            // 
            // label6
            // 
            label6.AutoSize = true;
            label6.Location = new Point(63, 647);
            label6.Name = "label6";
            label6.Size = new Size(97, 41);
            label6.TabIndex = 14;
            label6.Text = "Група";
            // 
            // label7
            // 
            label7.AutoSize = true;
            label7.Location = new Point(63, 760);
            label7.Name = "label7";
            label7.Size = new Size(633, 41);
            label7.TabIndex = 15;
            label7.Text = "Керівники практики (кафедра/підприємство)";
            // 
            // label8
            // 
            label8.AutoSize = true;
            label8.Location = new Point(63, 865);
            label8.Name = "label8";
            label8.Size = new Size(455, 41);
            label8.TabIndex = 16;
            label8.Text = "Керівники практики (факультет)";
            // 
            // button_GenerateDoc
            // 
            button_GenerateDoc.Location = new Point(72, 977);
            button_GenerateDoc.Name = "button_GenerateDoc";
            button_GenerateDoc.Size = new Size(486, 58);
            button_GenerateDoc.TabIndex = 17;
            button_GenerateDoc.Text = "Згенерувати щоденник";
            button_GenerateDoc.UseVisualStyleBackColor = true;
            button_GenerateDoc.Click += button_GenerateDoc_Click;
            // 
            // MetaDataForm
            // 
            AutoScaleDimensions = new SizeF(17F, 41F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(1735, 1103);
            Controls.Add(button_GenerateDoc);
            Controls.Add(label8);
            Controls.Add(label7);
            Controls.Add(label6);
            Controls.Add(label5);
            Controls.Add(textBox_MentorsFaculty);
            Controls.Add(textBox_MentorsDepartment);
            Controls.Add(textBox_Group);
            Controls.Add(textBox_PracticePlace);
            Controls.Add(label4);
            Controls.Add(label3);
            Controls.Add(label2);
            Controls.Add(label1);
            Controls.Add(dateTimePicker_EndDate);
            Controls.Add(dateTimePicker_StartDate);
            Controls.Add(comboBox_Gender);
            Controls.Add(textBox_GenitiveCaseName);
            Controls.Add(textBox_NominativeCaseName);
            Name = "MetaDataForm";
            Text = "MetaDataForm";
            ResumeLayout(false);
            PerformLayout();
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
        private TextBox textBox_MentorsDepartment;
        private TextBox textBox_MentorsFaculty;
        private Label label5;
        private Label label6;
        private Label label7;
        private Label label8;
        private Button button_GenerateDoc;
    }
}
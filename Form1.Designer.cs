namespace DocumentAnalyzer
{
    partial class Form1
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
            btn_OpenFile = new Button();
            richTextBox_DOCXPreview = new RichTextBox();
            numericUpDown_TopMargin = new NumericUpDown();
            numericUpDown_LeftMargin = new NumericUpDown();
            numericUpDown_BottomMargin = new NumericUpDown();
            numericUpDown_RightMargin = new NumericUpDown();
            label_Top = new Label();
            label_Bottom = new Label();
            label_Left = new Label();
            label_Right = new Label();
            comboBox_Fonts = new ComboBox();
            label_Font = new Label();
            numericUpDown_FontSize = new NumericUpDown();
            label_Size = new Label();
            checkBox_Pictures = new CheckBox();
            checkBox_Tables = new CheckBox();
            checkBox_Margins = new CheckBox();
            checkBox_Font = new CheckBox();
            ((System.ComponentModel.ISupportInitialize)numericUpDown_TopMargin).BeginInit();
            ((System.ComponentModel.ISupportInitialize)numericUpDown_LeftMargin).BeginInit();
            ((System.ComponentModel.ISupportInitialize)numericUpDown_BottomMargin).BeginInit();
            ((System.ComponentModel.ISupportInitialize)numericUpDown_RightMargin).BeginInit();
            ((System.ComponentModel.ISupportInitialize)numericUpDown_FontSize).BeginInit();
            SuspendLayout();
            // 
            // btn_OpenFile
            // 
            btn_OpenFile.Location = new Point(732, 63);
            btn_OpenFile.Name = "btn_OpenFile";
            btn_OpenFile.Size = new Size(188, 58);
            btn_OpenFile.TabIndex = 0;
            btn_OpenFile.Text = "Select file";
            btn_OpenFile.UseVisualStyleBackColor = true;
            btn_OpenFile.Click += btn_OpenFile_Click;
            // 
            // richTextBox_DOCXPreview
            // 
            richTextBox_DOCXPreview.Location = new Point(732, 151);
            richTextBox_DOCXPreview.Name = "richTextBox_DOCXPreview";
            richTextBox_DOCXPreview.Size = new Size(1497, 1020);
            richTextBox_DOCXPreview.TabIndex = 1;
            richTextBox_DOCXPreview.Text = "";
            // 
            // numericUpDown_TopMargin
            // 
            numericUpDown_TopMargin.DecimalPlaces = 2;
            numericUpDown_TopMargin.Increment = new decimal(new int[] { 25, 0, 0, 131072 });
            numericUpDown_TopMargin.Location = new Point(37, 238);
            numericUpDown_TopMargin.Maximum = new decimal(new int[] { 3, 0, 0, 0 });
            numericUpDown_TopMargin.Name = "numericUpDown_TopMargin";
            numericUpDown_TopMargin.Size = new Size(300, 47);
            numericUpDown_TopMargin.TabIndex = 2;
            numericUpDown_TopMargin.ValueChanged += numericUpDown_TopMargin_ValueChanged;
            // 
            // numericUpDown_LeftMargin
            // 
            numericUpDown_LeftMargin.DecimalPlaces = 2;
            numericUpDown_LeftMargin.Increment = new decimal(new int[] { 25, 0, 0, 131072 });
            numericUpDown_LeftMargin.Location = new Point(37, 353);
            numericUpDown_LeftMargin.Maximum = new decimal(new int[] { 3, 0, 0, 0 });
            numericUpDown_LeftMargin.Name = "numericUpDown_LeftMargin";
            numericUpDown_LeftMargin.Size = new Size(300, 47);
            numericUpDown_LeftMargin.TabIndex = 3;
            // 
            // numericUpDown_BottomMargin
            // 
            numericUpDown_BottomMargin.DecimalPlaces = 2;
            numericUpDown_BottomMargin.Increment = new decimal(new int[] { 25, 0, 0, 131072 });
            numericUpDown_BottomMargin.Location = new Point(386, 238);
            numericUpDown_BottomMargin.Maximum = new decimal(new int[] { 3, 0, 0, 0 });
            numericUpDown_BottomMargin.Name = "numericUpDown_BottomMargin";
            numericUpDown_BottomMargin.Size = new Size(300, 47);
            numericUpDown_BottomMargin.TabIndex = 4;
            // 
            // numericUpDown_RightMargin
            // 
            numericUpDown_RightMargin.DecimalPlaces = 2;
            numericUpDown_RightMargin.Increment = new decimal(new int[] { 25, 0, 0, 131072 });
            numericUpDown_RightMargin.Location = new Point(386, 353);
            numericUpDown_RightMargin.Maximum = new decimal(new int[] { 3, 0, 0, 0 });
            numericUpDown_RightMargin.Name = "numericUpDown_RightMargin";
            numericUpDown_RightMargin.Size = new Size(300, 47);
            numericUpDown_RightMargin.TabIndex = 5;
            // 
            // label_Top
            // 
            label_Top.AutoSize = true;
            label_Top.Location = new Point(34, 184);
            label_Top.Name = "label_Top";
            label_Top.Size = new Size(67, 41);
            label_Top.TabIndex = 6;
            label_Top.Text = "Top";
            // 
            // label_Bottom
            // 
            label_Bottom.AutoSize = true;
            label_Bottom.Location = new Point(386, 184);
            label_Bottom.Name = "label_Bottom";
            label_Bottom.Size = new Size(117, 41);
            label_Bottom.TabIndex = 7;
            label_Bottom.Text = "Bottom";
            // 
            // label_Left
            // 
            label_Left.AutoSize = true;
            label_Left.Location = new Point(34, 309);
            label_Left.Name = "label_Left";
            label_Left.Size = new Size(67, 41);
            label_Left.TabIndex = 8;
            label_Left.Text = "Left";
            // 
            // label_Right
            // 
            label_Right.AutoSize = true;
            label_Right.Location = new Point(386, 309);
            label_Right.Name = "label_Right";
            label_Right.Size = new Size(88, 41);
            label_Right.TabIndex = 9;
            label_Right.Text = "Right";
            // 
            // comboBox_Fonts
            // 
            comboBox_Fonts.FormattingEnabled = true;
            comboBox_Fonts.Items.AddRange(new object[] { "Times New Roman", "Calibri", "Custom" });
            comboBox_Fonts.Location = new Point(37, 533);
            comboBox_Fonts.Name = "comboBox_Fonts";
            comboBox_Fonts.Size = new Size(437, 49);
            comboBox_Fonts.TabIndex = 10;
            // 
            // label_Font
            // 
            label_Font.AutoSize = true;
            label_Font.Location = new Point(37, 479);
            label_Font.Name = "label_Font";
            label_Font.Size = new Size(78, 41);
            label_Font.TabIndex = 11;
            label_Font.Text = "Font";
            label_Font.Click += label5_Click;
            // 
            // numericUpDown_FontSize
            // 
            numericUpDown_FontSize.DecimalPlaces = 1;
            numericUpDown_FontSize.Increment = new decimal(new int[] { 5, 0, 0, 65536 });
            numericUpDown_FontSize.Location = new Point(505, 535);
            numericUpDown_FontSize.Maximum = new decimal(new int[] { 45, 0, 0, 0 });
            numericUpDown_FontSize.Minimum = new decimal(new int[] { 8, 0, 0, 0 });
            numericUpDown_FontSize.Name = "numericUpDown_FontSize";
            numericUpDown_FontSize.Size = new Size(181, 47);
            numericUpDown_FontSize.TabIndex = 12;
            numericUpDown_FontSize.Value = new decimal(new int[] { 8, 0, 0, 0 });
            // 
            // label_Size
            // 
            label_Size.AutoSize = true;
            label_Size.Location = new Point(505, 479);
            label_Size.Name = "label_Size";
            label_Size.Size = new Size(71, 41);
            label_Size.TabIndex = 13;
            label_Size.Text = "Size";
            // 
            // checkBox_Pictures
            // 
            checkBox_Pictures.AutoSize = true;
            checkBox_Pictures.Location = new Point(37, 616);
            checkBox_Pictures.Name = "checkBox_Pictures";
            checkBox_Pictures.Size = new Size(244, 45);
            checkBox_Pictures.TabIndex = 16;
            checkBox_Pictures.Text = "Pictures check";
            checkBox_Pictures.UseVisualStyleBackColor = true;
            // 
            // checkBox_Tables
            // 
            checkBox_Tables.AutoSize = true;
            checkBox_Tables.Location = new Point(37, 680);
            checkBox_Tables.Name = "checkBox_Tables";
            checkBox_Tables.Size = new Size(222, 45);
            checkBox_Tables.TabIndex = 17;
            checkBox_Tables.Text = "Tables check";
            checkBox_Tables.UseVisualStyleBackColor = true;
            // 
            // checkBox_Margins
            // 
            checkBox_Margins.AutoSize = true;
            checkBox_Margins.Checked = true;
            checkBox_Margins.CheckState = CheckState.Checked;
            checkBox_Margins.Location = new Point(37, 117);
            checkBox_Margins.Name = "checkBox_Margins";
            checkBox_Margins.Size = new Size(163, 45);
            checkBox_Margins.TabIndex = 18;
            checkBox_Margins.Text = "Margins";
            checkBox_Margins.UseVisualStyleBackColor = true;
            checkBox_Margins.CheckedChanged += checkBox_Margins_CheckedChanged;
            // 
            // checkBox_Font
            // 
            checkBox_Font.AutoSize = true;
            checkBox_Font.Checked = true;
            checkBox_Font.CheckState = CheckState.Checked;
            checkBox_Font.Location = new Point(37, 431);
            checkBox_Font.Name = "checkBox_Font";
            checkBox_Font.Size = new Size(116, 45);
            checkBox_Font.TabIndex = 19;
            checkBox_Font.Text = "Font";
            checkBox_Font.UseVisualStyleBackColor = true;
            checkBox_Font.CheckedChanged += checkBox_Font_CheckedChanged;
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(17F, 41F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(2275, 1241);
            Controls.Add(checkBox_Font);
            Controls.Add(checkBox_Margins);
            Controls.Add(checkBox_Tables);
            Controls.Add(checkBox_Pictures);
            Controls.Add(label_Size);
            Controls.Add(numericUpDown_FontSize);
            Controls.Add(label_Font);
            Controls.Add(comboBox_Fonts);
            Controls.Add(label_Right);
            Controls.Add(label_Left);
            Controls.Add(label_Bottom);
            Controls.Add(label_Top);
            Controls.Add(numericUpDown_RightMargin);
            Controls.Add(numericUpDown_BottomMargin);
            Controls.Add(numericUpDown_LeftMargin);
            Controls.Add(numericUpDown_TopMargin);
            Controls.Add(richTextBox_DOCXPreview);
            Controls.Add(btn_OpenFile);
            DoubleBuffered = true;
            Name = "Form1";
            Text = "Form1";
            ((System.ComponentModel.ISupportInitialize)numericUpDown_TopMargin).EndInit();
            ((System.ComponentModel.ISupportInitialize)numericUpDown_LeftMargin).EndInit();
            ((System.ComponentModel.ISupportInitialize)numericUpDown_BottomMargin).EndInit();
            ((System.ComponentModel.ISupportInitialize)numericUpDown_RightMargin).EndInit();
            ((System.ComponentModel.ISupportInitialize)numericUpDown_FontSize).EndInit();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Button btn_OpenFile;
        private RichTextBox richTextBox_DOCXPreview;
        private NumericUpDown numericUpDown_TopMargin;
        private NumericUpDown numericUpDown_LeftMargin;
        private NumericUpDown numericUpDown_BottomMargin;
        private NumericUpDown numericUpDown_RightMargin;
        private Label label_Top;
        private Label label_Bottom;
        private Label label_Left;
        private Label label_Right;
        private ComboBox comboBox_Fonts;
        private Label label_Font;
        private NumericUpDown numericUpDown_FontSize;
        private Label label_Size;
        private Label label7;
        private Label label8;
        private CheckBox checkBox_Pictures;
        private CheckBox checkBox_Tables;
        private CheckBox checkBox_Margins;
        private CheckBox checkBox_Font;
    }
}
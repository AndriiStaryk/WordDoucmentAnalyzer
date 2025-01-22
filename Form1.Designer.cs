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
            SuspendLayout();
            // 
            // btn_OpenFile
            // 
            btn_OpenFile.Location = new Point(223, 327);
            btn_OpenFile.Name = "btn_OpenFile";
            btn_OpenFile.Size = new Size(188, 58);
            btn_OpenFile.TabIndex = 0;
            btn_OpenFile.Text = "Select file";
            btn_OpenFile.UseVisualStyleBackColor = true;
            btn_OpenFile.Click += btn_OpenFile_Click;
            // 
            // richTextBox_DOCXPreview
            // 
            richTextBox_DOCXPreview.Location = new Point(500, 148);
            richTextBox_DOCXPreview.Name = "richTextBox_DOCXPreview";
            richTextBox_DOCXPreview.Size = new Size(1497, 1020);
            richTextBox_DOCXPreview.TabIndex = 1;
            richTextBox_DOCXPreview.Text = "";
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(17F, 41F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(2065, 1241);
            Controls.Add(richTextBox_DOCXPreview);
            Controls.Add(btn_OpenFile);
            DoubleBuffered = true;
            Name = "Form1";
            Text = "Form1";
            ResumeLayout(false);
        }

        #endregion

        private Button btn_OpenFile;
        private RichTextBox richTextBox_DOCXPreview;
    }
}
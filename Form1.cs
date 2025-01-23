namespace DocumentAnalyzer;

public partial class Form1 : Form
{
    public Form1()
    {
        InitializeComponent();
    }

    private void btn_OpenFile_Click(object sender, EventArgs e)
    {
        OpenFileDialog openFileDialog = new OpenFileDialog
        {
            Filter = "Word Documents|*.docx",
            Title = "Select a Word Document"
        };

        if (openFileDialog.ShowDialog() == DialogResult.OK)
        {
            string filePath = openFileDialog.FileName;
            DocxParser.ParseWordDocument(filePath.ToLower(), richTextBox_DOCXPreview);
        }
    }

    private void numericUpDown_TopMargin_ValueChanged(object sender, EventArgs e)
    {

    }

    private void label5_Click(object sender, EventArgs e)
    {

    }

    private void checkBox_Margins_CheckedChanged(object sender, EventArgs e)
    {
        ToggleVisibility(checkBox_Margins.Checked,
                        label_Top, label_Bottom, label_Left, label_Right,
                        numericUpDown_TopMargin, numericUpDown_BottomMargin,
                        numericUpDown_LeftMargin, numericUpDown_RightMargin);

    }

    private void checkBox_Font_CheckedChanged(object sender, EventArgs e)
    {
        ToggleVisibility(checkBox_Font.Checked,
                        label_Font, label_Size,
                        comboBox_Fonts,
                        numericUpDown_FontSize);
    }

    private void ToggleVisibility(bool isVisible, params Control[] controls)
    {
        foreach (var control in controls) { control.Visible = isVisible; }
    }

}

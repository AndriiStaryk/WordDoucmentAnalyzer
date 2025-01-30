namespace DocumentAnalyzer;

public partial class Form1 : Form
{
    private DocxParser _docxParser = new DocxParser();

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
            _docxParser.ParseWordDocument(filePath.ToLower(), richTextBox_DOCXPreview);
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
        ToggleAvailability(checkBox_Margins.Checked,
                        label_Top, label_Bottom, label_Left, label_Right,
                        numericUpDown_TopMargin, numericUpDown_BottomMargin,
                        numericUpDown_LeftMargin, numericUpDown_RightMargin);

    }

    private void checkBox_Font_CheckedChanged(object sender, EventArgs e)
    {
        ToggleAvailability(checkBox_Font.Checked,
                        label_Font, label_Size,
                        comboBox_Fonts,
                        numericUpDown_FontSize);
    }

    private void ToggleAvailability(bool isAvailable, params Control[] controls)
    {
        foreach (var control in controls) { control.Enabled = isAvailable; }
    }

    private void checkBox_Interval_CheckedChanged(object sender, EventArgs e)
    {
        ToggleAvailability(checkBox_Interval.Checked,
                           label_Interval, numericUpDown_Interval);
    }

    private void button_GenerateDoc_Click(object sender, EventArgs e)
    {
        _docxParser.GenerateDocument(new DocumentMetaData());
    }
}

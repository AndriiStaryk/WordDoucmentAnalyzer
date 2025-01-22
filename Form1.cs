namespace DocumentAnalyzer;

public partial class Form1 : Form
{
    private DocxParser docxParser = new DocxParser();
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
            docxParser.ParseWordDocument(filePath.ToLower(), richTextBox_DOCXPreview);
        }
    }
}

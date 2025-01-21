using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

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
            ParseWordDocument(filePath);
        }
    }

    private void ParseWordDocument(string filePath)
    {
        try
        {
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(filePath, false))
            {
                var paragraphs = wordDoc.MainDocumentPart.Document.Body.Elements<Paragraph>()
                                        .Select(p => p.InnerText)
                                        .ToList();

                listView.Items.Clear();
                foreach (var para in paragraphs)
                {
                    listView.Items.Add(para); 
                }

                // Extract font and style details
                var stylesPart = wordDoc.MainDocumentPart.StyleDefinitionsPart;
                if (stylesPart != null)
                {
                    foreach (var style in stylesPart.Styles.Elements<Style>())
                    {
                        richTextBox_DOCXPreview.Text += $"Style ID: {style.StyleId}, Name: {style.StyleName?.Val}{Environment.NewLine}";
                    }
                }

                // Get margins from section properties
                var sectionProps = wordDoc.MainDocumentPart.Document.Body.Elements<SectionProperties>().FirstOrDefault();
                if (sectionProps != null)
                {
                    var pageMargins = sectionProps.GetFirstChild<PageMargin>();
                    if (pageMargins != null)
                    {
                        richTextBox_DOCXPreview.Text += $"Margins: Top {pageMargins.Top}, Bottom {pageMargins.Bottom}, Left {pageMargins.Left}, Right {pageMargins.Right}{Environment.NewLine}";
                    }
                }
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error: {ex.Message}");
        }
    }
}

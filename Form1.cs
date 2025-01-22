using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
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
using Run = DocumentFormat.OpenXml.Wordprocessing.Run;
using RunProperties = DocumentFormat.OpenXml.Wordprocessing.RunProperties;

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
                richTextBox_DOCXPreview.Text += "Paragraph Details:\n\n";

                var mainPart = wordDoc.MainDocumentPart;
                if (mainPart == null)
                {
                    Console.WriteLine("Main document part not found.");
                    return;
                }

                var mainPartBody = mainPart.Document.Body;

                if (mainPartBody == null)
                {
                    Console.WriteLine("Body part not found.");
                    return;
                }

                // Extract margins
                extractMargins(mainPartBody);
               
                // Extract styles Fonts Names and Sizes
                extractStyles(mainPart);

                //Extract paragraphs aka regular Word Text
                parseParagraphs(mainPartBody);

            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error: {ex.Message}");
        }
    }

    private void extractMargins(Body body)
    {
        var sectionProps = body.Elements<SectionProperties>().FirstOrDefault();
        if (sectionProps != null)
        {
            var pageMargins = sectionProps.GetFirstChild<PageMargin>();
            if (pageMargins != null)
            {
                richTextBox_DOCXPreview.Text += "Margins:\n\n";
                richTextBox_DOCXPreview.Text += $"  Top: {pageMargins.Top} twips\n";
                richTextBox_DOCXPreview.Text += $"  Bottom: {pageMargins.Bottom} twips\n";
                richTextBox_DOCXPreview.Text += $"  Left: {pageMargins.Left} twips\n";
                richTextBox_DOCXPreview.Text += $"  Right: {pageMargins.Right} twips\n";
            }
        }
    }

    private void extractStyles(MainDocumentPart mdp)
    {
        var stylesPart = mdp.StyleDefinitionsPart;
        if (stylesPart != null)
        {
            richTextBox_DOCXPreview.Text += "Styles:\n\n";

            var stylesProps = stylesPart.Styles;
            if (stylesProps != null)
            {
                var styles = stylesProps.Elements<Style>();
                foreach (var style in styles)
                {
                    richTextBox_DOCXPreview.Text += $"Style ID: {style.StyleId}, Name: {style.StyleName?.Val}{Environment.NewLine}";

                    var runProps = style.StyleRunProperties;
                    if (runProps != null)
                    {
                        var runFonts = runProps.GetFirstChild<RunFonts>();
                        if (runFonts != null)
                        {
                            richTextBox_DOCXPreview.Text += $"  Default Font: {runFonts.Ascii ?? "Not Set"}{Environment.NewLine}";

                        }

                        var fontSize = runProps.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.FontSize>();
                        if (fontSize != null)
                        {
                            richTextBox_DOCXPreview.Text += $"  Default Font Size: {fontSize.Val} (in half-points){Environment.NewLine}";
                        }
                    }
                }

                richTextBox_DOCXPreview.Text += "\n";
            }
        }
    }
    private void parseParagraphs(Body body)
    {
        var paragraphs = body.Elements<DocumentFormat.OpenXml.Wordprocessing.Paragraph>();
        foreach (var paragraph in paragraphs)
        {
            // Paragraph text
            richTextBox_DOCXPreview.Text += $"Text: {paragraph.InnerText}\n";

            // Paragraph properties (spacing)
            var paragraphProps = paragraph.ParagraphProperties;
            if (paragraphProps != null)
            {
                var spacing = paragraphProps.SpacingBetweenLines;
                if (spacing != null)
                {
                    richTextBox_DOCXPreview.Text += $"  Line Spacing: {spacing.Line ?? "Default"} (in 20ths of a point)\n";
                }
            }

            richTextBox_DOCXPreview.Text += "\n";
        }
    }

}

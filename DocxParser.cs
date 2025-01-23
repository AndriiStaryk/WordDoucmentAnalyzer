using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocumentAnalyzer;

internal class DocxParser
{
    private AnalyzeItem analyzeItem;

    public DocxParser(AnalyzeItem item)
    {
        analyzeItem = new AnalyzeItem();
    }

    public void CompareItems(AnalyzeItem itemToCompareWith)
    {
        //TODO
    }

    public void ParseWordDocument(string filePath, RichTextBox richTextBox)
    {
        try
        {
            using WordprocessingDocument wordDoc = WordprocessingDocument.Open(filePath, false);

            richTextBox.Text += "Paragraph Details:\n\n";

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

            ExtractMargins(mainPartBody);
            ExtractStyles(mainPart);
            ParseParagraphs(mainPartBody, richTextBox);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error: {ex.Message}");
        }
    }

    private void ExtractMargins(Body body)
    {
        var sectionProps = body.Elements<SectionProperties>().FirstOrDefault();
        if (sectionProps != null)
        {
            var pageMargins = sectionProps.GetFirstChild<PageMargin>();
            if (pageMargins != null)
            {
                analyzeItem.Margin = new Margin(pageMargins.Top ?? 0,
                                                pageMargins.Bottom ?? 0,
                                                pageMargins.Left ?? 0,
                                                pageMargins.Right ?? 0);
            }
        }
    }

    private void ExtractStyles(MainDocumentPart mdp)
    {
        var stylesPart = mdp.StyleDefinitionsPart;
        if (stylesPart != null)
        {
            var stylesProps = stylesPart.Styles;
            if (stylesProps != null)
            {
                var styles = stylesProps.Elements<Style>();
                foreach (var style in styles)
                {
                    var runProps = style.StyleRunProperties;
                    if (runProps != null)
                    {
                        FontInfo fontInfo = new FontInfo();
                        var runFonts = runProps.GetFirstChild<RunFonts>();
                        if (runFonts != null)
                        { 
                            fontInfo.Name = $"{runFonts.Ascii ?? "Not Set"}";
                        }

                        var fontSize = runProps.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.FontSize>();
                        if (fontSize != null)
                        {
                            if (fontSize.Val != null)
                            {
                                fontInfo.Size = (double)(Convert.ToInt32(fontSize.Val)) / 2 ;
                            }
                        }

                        analyzeItem.FontInfo = fontInfo;
                    }
                }

            }
        }
    }
    private static void ParseParagraphs(Body body, RichTextBox richTextBox)
    {
        var paragraphs = body.Elements<DocumentFormat.OpenXml.Wordprocessing.Paragraph>();
        foreach (var paragraph in paragraphs)
        {
            richTextBox.Text += $"Text: {paragraph.InnerText}\n";

            var paragraphProps = paragraph.ParagraphProperties;
            if (paragraphProps != null)
            {
                var spacing = paragraphProps.SpacingBetweenLines;
                if (spacing != null)
                {
                    richTextBox.Text += $"  Line Spacing: {spacing.Line ?? "Default"} (in 20ths of a point)\n";
                }
            }

            richTextBox.Text += "\n";
        }
    }
}

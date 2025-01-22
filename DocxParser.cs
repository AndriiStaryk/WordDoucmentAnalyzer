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
    public void ParseWordDocument(string filePath, RichTextBox richTextBox)
    {
        try
        {
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(filePath, false))
            {
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

                // Extract margins
                extractMargins(mainPartBody, richTextBox);

                // Extract styles Fonts Names and Sizes
                extractStyles(mainPart, richTextBox);

                //Extract paragraphs aka regular Word Text
                parseParagraphs(mainPartBody, richTextBox);

            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error: {ex.Message}");
        }
    }

    private void extractMargins(Body body, RichTextBox richTextBox)
    {
        var sectionProps = body.Elements<SectionProperties>().FirstOrDefault();
        if (sectionProps != null)
        {
            var pageMargins = sectionProps.GetFirstChild<PageMargin>();
            if (pageMargins != null)
            {
                richTextBox.Text += "Margins:\n\n";
                richTextBox.Text += $"  Top: {pageMargins.Top} twips\n";
                richTextBox.Text += $"  Bottom: {pageMargins.Bottom} twips\n";
                richTextBox.Text += $"  Left: {pageMargins.Left} twips\n";
                richTextBox.Text += $"  Right: {pageMargins.Right} twips\n";
            }
        }
    }

    private void extractStyles(MainDocumentPart mdp, RichTextBox richTextBox)
    {
        var stylesPart = mdp.StyleDefinitionsPart;
        if (stylesPart != null)
        {
            richTextBox.Text += "Styles:\n\n";

            var stylesProps = stylesPart.Styles;
            if (stylesProps != null)
            {
                var styles = stylesProps.Elements<Style>();
                foreach (var style in styles)
                {
                    richTextBox.Text += $"Style ID: {style.StyleId}, Name: {style.StyleName?.Val}{Environment.NewLine}";

                    var runProps = style.StyleRunProperties;
                    if (runProps != null)
                    {
                        var runFonts = runProps.GetFirstChild<RunFonts>();
                        if (runFonts != null)
                        {
                            richTextBox.Text += $"  Default Font: {runFonts.Ascii ?? "Not Set"}{Environment.NewLine}";
                        }

                        var fontSize = runProps.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.FontSize>();
                        if (fontSize != null)
                        {
                            richTextBox.Text += $"  Default Font Size: {fontSize.Val} (in half-points){Environment.NewLine}";
                        }
                    }
                }

                richTextBox.Text += "\n";
            }
        }
    }
    private void parseParagraphs(Body body, RichTextBox richTextBox)
    {
        var paragraphs = body.Elements<DocumentFormat.OpenXml.Wordprocessing.Paragraph>();
        foreach (var paragraph in paragraphs)
        {
            // Paragraph text
            richTextBox.Text += $"Text: {paragraph.InnerText}\n";

            // Paragraph properties (spacing)
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

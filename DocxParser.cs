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
            ParseParagraphs(mainPartBody, wordDoc, richTextBox);
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
                analyzeItem.Margin = new Margin((pageMargins.Top) ?? 0,
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
    private void ParseParagraphs(Body body, WordprocessingDocument wordDoc, RichTextBox richTextBox)
    {
        //For debuggging purposes
        richTextBox.Text += analyzeItem.ToString();
        var paragraphs = body.Elements<DocumentFormat.OpenXml.Wordprocessing.Paragraph>();

        foreach (var paragraph in paragraphs)
        {
            richTextBox.Text += $"Text: {paragraph.InnerText}\n";

            string lastFontName = "";
            double lastFontSize = 0;

            var runs = paragraph.Elements<DocumentFormat.OpenXml.Wordprocessing.Run>();
            foreach (var run in runs)
            {
                var runProperties = run.RunProperties;
                string fontName = "Default";
                double fontSizeValue = 11.0; 

                if (runProperties != null)
                {
                    var runFonts = runProperties.GetFirstChild<RunFonts>();
                    if (runFonts != null)
                    {
                        fontName = runFonts.Ascii ?? runFonts.HighAnsi ?? "Default";
                    }

                    var fontSize = runProperties.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.FontSize>();
                    if (fontSize?.Val != null)
                    {
                        fontSizeValue = Convert.ToInt32(fontSize.Val) / 2.0;
                    }
                }
                else
                {
                    var paraStyleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val;
                    if (paraStyleId != null)
                    {
                        fontName = GetFontFromStyle(wordDoc, paraStyleId);
                        fontSizeValue = GetFontSizeFromStyle(wordDoc, paraStyleId);
                    }
                }

                if (fontName != lastFontName || fontSizeValue != lastFontSize)
                {
                    richTextBox.Text += $"  Font: {fontName}, Size: {fontSizeValue}pt\n";
                    lastFontName = fontName;
                    lastFontSize = fontSizeValue;
                }
            }
            richTextBox.Text += "\n";
        }
    }

    private string GetFontFromStyle(WordprocessingDocument wordDoc, string styleId)
    {
        var stylesPart = wordDoc.MainDocumentPart.StyleDefinitionsPart;
        if (stylesPart != null)
        {
            var style = stylesPart.Styles.Elements<Style>().FirstOrDefault(s => s.StyleId == styleId);
            if (style != null)
            {
                var runProps = style.StyleRunProperties;
                var runFonts = runProps?.GetFirstChild<RunFonts>();
                return runFonts?.Ascii ?? runFonts?.HighAnsi ?? "Default";
            }
        }
        return "Default";
    }

    private double GetFontSizeFromStyle(WordprocessingDocument wordDoc, string styleId)
    {
        var stylesPart = wordDoc.MainDocumentPart.StyleDefinitionsPart;
        if (stylesPart != null)
        {
            var style = stylesPart.Styles.Elements<Style>().FirstOrDefault(s => s.StyleId == styleId);
            if (style != null)
            {
                var runProps = style.StyleRunProperties;
                var fontSize = runProps?.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.FontSize>();
                if (fontSize?.Val != null)
                {
                    return Convert.ToInt32(fontSize.Val) / 2.0;
                }
            }
        }
        return 11.0; 
    }

}

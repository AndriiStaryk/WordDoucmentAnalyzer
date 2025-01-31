using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DocumentAnalyzer;

internal class DocxManager
{
    private AnalyzeItem _analyzeItem = new AnalyzeItem();

    string copyFilePath = Path.Combine(Application.StartupPath, "modified_template.docx");

    public void CompareItems(AnalyzeItem itemToCompareWith)
    {

    }

    public void GenerateDocument(DocumentMetaData data)
    {
        CreateCopyOfTemplate();
        
        var replacements = new Dictionary<string, string>
        {
            { "{{NominativeCaseName}}", data.NominativeCaseName },
            { "{{GenitiveCaseName}}", data.GenitiveCaseName },
            { "{{Gender}}", data.Gender.ToDescription() },
            { "{{StartDate}}", data.StartDate.ToShortDateString() },
            { "{{EndDate}}", data.EndDate.ToShortDateString() },
            { "{{PracticePlace}}", data.PracticePlace },
            { "{{Group}}", data.Group },
            { "{{MentorsFromDepartment}}", data.MentorsFromDepartment },
            { "{{MentorsFromFaculty}}", data.MentorsFromFaculty }
        };

        ReplacePlaceholders(copyFilePath, replacements);
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
                _analyzeItem.Margin = new Margin((pageMargins.Top) ?? 0,
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
                                fontInfo.Size = (double)(Convert.ToInt32(fontSize.Val)) / 2;
                            }
                        }

                        _analyzeItem.FontInfo = fontInfo;
                    }
                }

            }
        }
    }
    private void ParseParagraphs(Body body, WordprocessingDocument wordDoc, RichTextBox richTextBox)
    {
        //For debuggging purposes
        //richTextBox.Text += analyzeItem.ToString();
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


    public void CreateCopyOfTemplate()
    {
        string originalFilePath = Path.Combine(Application.StartupPath, "Resources", "diaryFixed.docx");
        File.Copy(originalFilePath, copyFilePath, true);
    }


    public void ReplacePlaceholders(string filePath, Dictionary<string, string> replacements)
    {

        using WordprocessingDocument doc = WordprocessingDocument.Open(filePath, true);
        
        MainDocumentPart mainPart = doc.MainDocumentPart;
        if (mainPart == null)
        {
            throw new InvalidOperationException("Main document part not found.");
        }

        
        var paragraphs = mainPart.Document.Body.Elements<DocumentFormat.OpenXml.Wordprocessing.Paragraph>();

        
        foreach (var paragraph in paragraphs)
        {           
            string paragraphText = paragraph.InnerText;

            foreach (var replacement in replacements)
            {
                if (paragraphText.Contains(replacement.Key))
                {
                    ReplaceTextInParagraph(paragraph, replacement.Key, replacement.Value);
                }
            }
        }

        mainPart.Document.Save();
    }

    private void ReplaceTextInParagraph(Paragraph paragraph, string placeholder, string replacementText)
    {
        string paragraphText = string.Join("", paragraph.Elements<Run>().Select(run => run.InnerText));

        if (paragraphText.Contains(placeholder))
        {
            string[] parts = paragraphText.Split(new[] { placeholder }, StringSplitOptions.None);

            foreach (Run run in paragraph.Elements<Run>())
            {
                Text text = run.Elements<Text>().FirstOrDefault();
                if (text != null)
                {
                    text.Text = ""; 
                }
            }

            int runIndex = 0;
            var runs = paragraph.Elements<Run>().ToList();

            for (int i = 0; i < parts.Length; i++)
            {
                if (!string.IsNullOrEmpty(parts[i]))
                {
                    if (runIndex < runs.Count)
                    {
                        Text text = runs[runIndex].Elements<Text>().FirstOrDefault();
                        if (text != null)
                        {
                            text.Text = parts[i];
                        }
                        runIndex++;
                    }
                    else
                    {
                        paragraph.AppendChild(new Run(new Text(parts[i])));
                    }
                }

                if (i < parts.Length - 1)
                {
                    if (runIndex < runs.Count)
                    {
                        Text text = runs[runIndex].Elements<Text>().FirstOrDefault();
                        if (text != null)
                        {
                            text.Text = replacementText;
                        }
                        runIndex++;
                    }
                    else
                    {
                        paragraph.AppendChild(new Run(new Text(replacementText)));
                    }
                }
            }
        }
    }
}
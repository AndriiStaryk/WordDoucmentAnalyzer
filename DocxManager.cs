﻿using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace DocumentAnalyzer;

internal class DocxManager
{
    public void GenerateDocument(DocumentMetaData data)
    {
        FileManager.CreateCopyOfTemplate();
        
        var replacements = new Dictionary<string, string>
        {
            { "{{NominativeCaseName}}", data.NominativeCaseName },
            { "{{GenderNominativeCase}}", data.Gender.ToDescription() },
            { "{{GenderNominativeCaseLC}}", data.Gender.ToDescription().ToLower() },
            { "{{GenitiveCaseName}}", data.GenitiveCaseName },
            { "{{GenderGenitiveCase}}", data.Gender.ToDescription(true) },
            { "{{GenderGenitiveCaseLC}}", data.Gender.ToDescription(true).ToLower() },
            { "{{StartDate}}", data.StartDate.ToShortDateString() },
            { "{{EndDate}}", data.EndDate.ToShortDateString() },
            { "{{PracticePlace}}", data.PracticePlace },
            { "{{Group}}", data.Group },
            { "{{MentorsFromDepartment}}", data.MentorsFromDepartment },
            { "{{MentorsFromFaculty}}", data.MentorsFromFaculty },
        };

        ReplacePlaceholders(replacements);

        List<string> taskDescriptionLines = SplitTextIntoLines(data.TaskDescription);
        Table taskDescriptionTable = TableGenerator.CreateSimpleTableBasedOnLines(taskDescriptionLines);
        ReplacePlaceholderWithTable("{{TaskDescriptionTable}}", taskDescriptionTable);

        List<string> characteristicsLines = SplitTextIntoLines(data.Characteristics);
        Table characteristicsTable = TableGenerator.CreateSimpleTableBasedOnLines(characteristicsLines);
        ReplacePlaceholderWithTable("{{CharacteristicTable}}", characteristicsTable);

        

        List<List<string>> dailyTasksDescription = data.DailyTasks
                                                        .Select(task => task.ToStringList())
                                                        .ToList();
        Table dailyTasksTable = TableGenerator.CreateDailyTasksDescriptionTable(dailyTasksDescription);
        TableGenerator.MergeCells(dailyTasksTable, new List<(int, int)> { (0, 0), (1, 0) }, TableGenerator.MergeDirection.Vertical);
        TableGenerator.MergeCells(dailyTasksTable, new List<(int, int)> { (0, 1), (1, 1) }, TableGenerator.MergeDirection.Vertical);
        TableGenerator.MergeCells(dailyTasksTable, new List<(int, int)> { (0, 4), (1, 4) }, TableGenerator.MergeDirection.Vertical);
        TableGenerator.MergeCells(dailyTasksTable, new List<(int, int)> { (0, 2), (0, 3) }, TableGenerator.MergeDirection.Horizontal);
        
        ReplacePlaceholderWithTable("{{DailyTasksTable}}", dailyTasksTable);

        FileManager.OpenDocx(FileManager.copyFilePath);
    }

    public void ReplacePlaceholders(Dictionary<string, string> replacements)
    {
        using WordprocessingDocument doc = WordprocessingDocument.Open(FileManager.copyFilePath, true);

        MainDocumentPart mainPart = doc.MainDocumentPart;
        if (mainPart == null)
        {
            throw new InvalidOperationException("Main document part not found.");
        }

        var paragraphs = mainPart.Document.Body.Descendants<Paragraph>();

        foreach (var paragraph in paragraphs)
        {
            SearchReplacementsAndReplace(paragraph, replacements);
        }
        
        mainPart.Document.Save();
    }

    private void SearchReplacementsAndReplace(Paragraph paragraph, Dictionary<string, string> replacements)
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

    private void ReplaceTextInParagraph(Paragraph paragraph, string placeholder, string replacementText)
    {
        string paragraphText = string.Join("", paragraph.Elements<Run>().Select(run => run.InnerText));

        if (paragraphText.Contains(placeholder))
        {
            int placeholderStart = paragraphText.IndexOf(placeholder);
            int placeholderEnd = placeholderStart + placeholder.Length;

            // Iterate through the runs to find the runs that contain the placeholder
            int currentPosition = 0;
            List<Run> runsToUpdate = new List<Run>();
            List<string> runTexts = new List<string>();

            foreach (Run run in paragraph.Elements<Run>())
            {
                Text text = run.Elements<Text>().FirstOrDefault();
                if (text != null)
                {
                    int runStart = currentPosition;
                    int runEnd = runStart + text.Text.Length;

                    if (runStart < placeholderEnd && runEnd > placeholderStart)
                    {
                        runsToUpdate.Add(run);
                        runTexts.Add(text.Text);
                    }

                    currentPosition += text.Text.Length;
                }
            }

            // Reconstruct the runs with the placeholder replaced
            if (runsToUpdate.Count > 0)
            {
                string combinedText = string.Join("", runTexts);
                int placeholderIndex = combinedText.IndexOf(placeholder);

                if (placeholderIndex >= 0)
                {
                    string before = combinedText.Substring(0, placeholderIndex);
                    string after = combinedText.Substring(placeholderIndex + placeholder.Length);

                    runsToUpdate[0].GetFirstChild<Text>().Text = before + replacementText;
                    for (int i = 1; i < runsToUpdate.Count; i++)
                    {
                        runsToUpdate[i].GetFirstChild<Text>().Text = "";
                    }

                    if (!string.IsNullOrEmpty(after))
                    {
                        runsToUpdate[runsToUpdate.Count - 1].GetFirstChild<Text>().Text += after;
                    }
                }
            }
        }
    }

    private void ReplacePlaceholderWithTable(string placeholder, Table table)
    {
        using WordprocessingDocument doc = WordprocessingDocument.Open(FileManager.copyFilePath, true);

        MainDocumentPart mainPart = doc.MainDocumentPart;
        if (mainPart == null)
        {
            throw new InvalidOperationException("Main document part not found.");
        }

        var body = mainPart.Document.Body;
        
        var paragraph = body.Descendants<Paragraph>().FirstOrDefault(p => p.InnerText.Contains(placeholder));

        if (paragraph != null)
        {
            body.InsertAfter(table, paragraph);
            paragraph.Remove();
        }
        
        mainPart.Document.Save();
    }
   
    private List<string> SplitTextIntoLines(string text, int maxCharactersPerRow = 68)
    {
        List<string> lines = new List<string>();
        int startIndex = 0;

        while (startIndex < text.Length)
        {
            int length = Math.Min(maxCharactersPerRow, text.Length - startIndex);

            if (startIndex + length < text.Length && text[startIndex + length] != ' ')
            {
                int lastSpace = text.LastIndexOf(' ', startIndex + length);
                if (lastSpace > startIndex)
                {
                    length = lastSpace - startIndex;
                }
            }

            lines.Add(text.Substring(startIndex, length).Trim());
            startIndex += length;
        }

        return lines;
    }



    private AnalyzeItem _analyzeItem = new AnalyzeItem();

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
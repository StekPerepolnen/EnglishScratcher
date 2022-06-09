// See https://aka.ms/new-console-template for more information
using AngleSharp.Html.Parser;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Net;
using System.Text.RegularExpressions;

namespace EnglishScratcher
{
    public record ParsedWordCard
    {
        public string Term { get; set; } = String.Empty;
        public string PartOfSpeech { get; set; } = String.Empty;
        public string Title { get; set; } = String.Empty;
        public string Level { get; set; } = String.Empty;
        public string Definition { get; set; } = String.Empty;
        public string Link { get; set; } = String.Empty;
        public List<string> Quotes { get; set; } = new List<string>();
    }

    public class Program
    {
        public static List<string> forbiddenPartsOfSpeech = new List<string> { "determiner", "conjunction" };

        public static void Main()
        {
            var wordCards = new List<ParsedWordCard>();
            for (int pageIndex = 1; pageIndex < 6756; pageIndex++)
            {
                string html = string.Empty;
                string url = @"https://www.englishprofile.org/american-english/words/usdetail/" + pageIndex;

                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);

                using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
                using (Stream stream = response.GetResponseStream())
                using (StreamReader reader = new StreamReader(stream))
                {
                    html = reader.ReadToEnd();
                }

                var parser = new HtmlParser();
                var document = parser.ParseDocument(html);

                string? term = document.QuerySelector(".headword")?.InnerHtml;
                string? partOfSpeech = document.QuerySelector(".pos")?.InnerHtml;

                if (forbiddenPartsOfSpeech.Contains(partOfSpeech))
                    continue;
                foreach (var body in document.QuerySelectorAll(".info.sense"))
                {
                    string? title;
                    string? definition;
                    string? level;
                    List<string> quotes = new List<string>();
                    title = body.QuerySelector(".sense_title")?.InnerHtml;
                    level = body.QuerySelector(".label")?.InnerHtml;
                    definition = body.QuerySelector(".definition")?.InnerHtml;
                    foreach (var element in body.QuerySelectorAll(".blockquote"))
                    {
                        var clearElement = Regex.Replace(element.InnerHtml, "<.*?>", String.Empty);
                        quotes.Add(clearElement);
                    }

                    Console.WriteLine(pageIndex + " " + title);
                    if (title != null && definition != null && level != null)
                    {
                        wordCards.Add(new ParsedWordCard()
                        {
                            Term = term,
                            Title = title,
                            Definition = definition,
                            Level = level,
                            PartOfSpeech = partOfSpeech,
                            Link = url,
                            Quotes = quotes
                        });
                    }
                }

            }

            //
            for (int i = 0; i < wordCards.Count; i++)
            {
                var wordCard = wordCards[i];
                for (int j = 0; j < wordCard.Quotes.Count; j++)
                {
                    wordCard.Quotes[j] = wordCard.Quotes[j].Replace(wordCard.Term, "[...]");
                }
            }

            // If you use EPPlus in a noncommercial context
            // according to the Polyform Noncommercial license:
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var package = new ExcelPackage())
            {
                ExcelWorksheet sheet = package.Workbook.Worksheets.Add("Dictionary");
                package.Workbook.Worksheets.Add("Deck");

                var namedStyle = sheet.Workbook.Styles.CreateNamedStyle("HyperLink");
                namedStyle.Style.Font.UnderLine = true;
                namedStyle.Style.Font.Color.SetColor(System.Drawing.Color.Blue);

                sheet.Cells[1, 1].Value = "Words";
                sheet.Cells[1, 1].Style.Font.Bold = true;
                sheet.Cells[1, 1].Style.Font.Color.SetColor(System.Drawing.Color.White);
                sheet.Cells[1, 1].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                sheet.Cells[1, 1].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DarkSlateGray);
                sheet.Cells[1, 2].Value = "Cards";
                sheet.Cells[1, 2].Style.Font.Bold = true;
                sheet.Cells[1, 2].Style.Font.Color.SetColor(System.Drawing.Color.White);
                sheet.Cells[1, 2].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                sheet.Cells[1, 2].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DarkSlateGray);
                sheet.Cells[1, 2].Style.Border.Top.Style = ExcelBorderStyle.Medium;
                sheet.Cells[1, 2].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                sheet.Cells[1, 2].Style.Border.Left.Style = ExcelBorderStyle.Medium;
                sheet.Cells[1, 2].Style.Border.Right.Style = ExcelBorderStyle.Medium;

                for (int i = 0; i < wordCards.Count; i++)
                {
                    var wordCard = wordCards[i];
                    var row = i + 2;
                    //var definition = $"({wordCard.PartOfSpeech}) {wordCard.Definition}";

                    sheet.Cells[row, 1].Value = $"{wordCard.Title} [{wordCard.Level}]";
                    sheet.Cells[row, 1].Hyperlink = new ExcelHyperLink(wordCard.Link);
                    //sheet.Cells[row, 1].StyleName = namedStyle.Name;

                    var cell = sheet.Cells[row, 2];
                    cell.IsRichText = true;
                    var richText = cell.RichText.Add($"({wordCard.PartOfSpeech})");
                    richText.Italic = true;
                    richText.Color = System.Drawing.Color.DarkOrange;
                    richText = cell.RichText.Add($" {wordCard.Definition}");
                    richText.Italic = false;
                    richText.Color = System.Drawing.Color.Black;
                    if (wordCard.Quotes.Count > 0)
                    {
                        foreach (var quote in wordCard.Quotes)
                        {
                            richText = cell.RichText.Add("\n" + quote);
                            richText.Italic = true;
                            richText.Color = System.Drawing.Color.LightSkyBlue;
                        }
                    }
                    cell.AddComment(wordCard.Term);
                }

                // Setting & getting values
                sheet.Columns[1].Width = 40;
                sheet.Columns[1].Style.VerticalAlignment = ExcelVerticalAlignment.Top;

                sheet.Columns[2].Width = 80;
                sheet.Columns[2].Style.WrapText = true;
                sheet.Columns[2].Style.VerticalAlignment = ExcelVerticalAlignment.Top;

                package.SaveAs(new FileInfo(@"EnglishDecks.xlsx"));
            }
        }
    }
}
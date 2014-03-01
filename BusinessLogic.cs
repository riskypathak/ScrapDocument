using DocumentFormat.OpenXml.Packaging;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Linq;

namespace ScrapDocument
{
    public static class BusinessLogic
    {
        private static XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        private static XName r = w + "r";
        private static XName ins = w + "ins";

        private static XmlWriter writer = XmlWriter.Create(@"ArticleInfo.xml");

        public static XmlDocument xmlDocument = null;
        private static int previousEndPage = 0;

        internal static Article ExtractArticleData(string fileName, string templateType)
        {
            Article article = new Article() { Filename = Path.GetFileName(fileName) };

            int pageCount = 0;

            foreach (string templateItem in (new TemplateXml()).ReadTemplate(templateType))
            {
                using (WordprocessingDocument doc = WordprocessingDocument.Open(fileName, false))
                {
                    var element = doc.MainDocumentPart
                        .GetXDocument()
                        .Descendants(w + "sdt")
                        .Where(e => ((string)e.Elements(w + "sdtPr")
                                              .Elements(w + "alias")
                                              .Attributes(w + "val")
                                              .FirstOrDefault()).ToLower() == templateItem)
                        .Select(b => GetTextFromContentControl(b));

                    foreach (string value in element)
                    {
                        PopulateArticle(ref article, templateItem, SanitizeXmlString(value));
                    }

                    pageCount = Convert.ToInt32(doc.ExtendedFilePropertiesPart.Properties.Pages.InnerText.ToString());
                }
            }

            article.FirstPage = previousEndPage + 1;
            previousEndPage = article.LastPage = article.FirstPage + pageCount - 1;
            return article;
        }

        private static string GetTextFromContentControl(XElement contentControlNode)
        {
            return contentControlNode.Descendants(w + "p")
                .Select
                (
                    p => p.Elements()
                          .Where(z => z.Name == r || z.Name == ins)
                          .Descendants(w + "t")
                          .StringConcatenate(element => (string)element) + Environment.NewLine
                ).StringConcatenate();
        }

        internal static string StringConcatenate<T>(this IEnumerable<T> source, Func<T, string> func)
        {
            StringBuilder sb = new StringBuilder();
            foreach (T item in source)
                sb.Append(func(item));
            return sb.ToString();
        }

        internal static string StringConcatenate(this IEnumerable<string> source)
        {
            StringBuilder sb = new StringBuilder();
            foreach (string item in source)
            {
                sb.Append(item);
            }
            return sb.ToString();
        }

        private static void PopulateArticle(ref Article article, string templateItem, string value)
        {
            switch (templateItem)
            {
                case "title": article.Title = value; break;
                case "english_title": article.EnglishTitle = value; break;
                case "contributors":
                    {
                        article.Contributors = GetContributors(value);
                    }
                    break;
                case "keywords": article.Keywords = value.Split(new char[1] { ';' }, StringSplitOptions.RemoveEmptyEntries).ToList(); break;
                case "english_keywords": article.EnglishKeywords = value.Split(';').ToList(); break;
                case "abstract": article.Abstract = value; break;
                case "english_abstract": article.Title = value; break;
                case "citation_list":
                    {
                        article.Citations = GetCitations(value);
                    } break;
            }
        }

        internal static XDocument GetXDocument(this OpenXmlPart part)
        {
            XDocument xdoc = part.Annotation<XDocument>();
            if (xdoc != null)
                return xdoc;
            using (StreamReader sr = new StreamReader(part.GetStream()))
            using (XmlReader xr = XmlReader.Create(sr))
                xdoc = XDocument.Load(xr);
            part.AddAnnotation(xdoc);
            return xdoc;
        }

        private static IList<string> GetCitations(string value)
        {
            IList<string> citationList = new List<string>();

            if (!string.IsNullOrWhiteSpace(value))
            {
                string[] citations = value.Split(new string[1] { "\r" }, StringSplitOptions.RemoveEmptyEntries);

                citationList = citations.ToList();
            }
            return citationList;
        }

        private static IList<Contributor> GetContributors(string value)
        {
            IList<Contributor> contributorsList = new List<Contributor>();

            if (!string.IsNullOrWhiteSpace(value))
            {
                string[] contributorNames = value.Trim().Split(new char[1] { ';' }, StringSplitOptions.RemoveEmptyEntries);

                foreach (string contributorName in contributorNames)
                {
                    Contributor contributor = new Contributor();

                    var names = contributorName.Split(',');

                    if (names.Length.Equals(2))
                    {
                        contributor.Surname = names[0].Trim();
                        contributor.GivenName = names[1].Trim();
                    }
                    else
                    {
                        contributor.GivenName = names[0].Trim();
                    }

                    contributorsList.Add(contributor);
                }
            }

            return contributorsList;
        }

        internal static string GenerateXML(IList<Article> ArticlesData)
        {
            XmlDocument xmlJournalDocument = new XmlDocument();

            using (writer = xmlJournalDocument.CreateNavigator().AppendChild())
            {
                writer.WriteStartElement("journal");

                int articleIndex = 0;

                foreach (Article article in ArticlesData)
                {
                    writer.WriteStartElement("journal_article");
                    writer.WriteElementString("filename", article.Filename);

                    writer.WriteElementString("title", article.Title.JournalText());
                    writer.WriteElementString("english_title", article.EnglishTitle.JournalText());
                    //writer.WriteElementString("spanish_title", article.Title.JournalText());

                    writer.WriteStartElement("contributors");

                    int contributorIndex = 0;
                    foreach (Contributor contributor in article.Contributors)
                    {

                        writer.WriteStartElement("author");
                        writer.WriteAttributeString("key", string.Format("cnbot{0}-{1}", articleIndex, contributorIndex));

                        if (!string.IsNullOrEmpty(contributor.Surname))
                        {
                            writer.WriteString(string.Format("{0}, {1}", contributor.Surname, contributor.GivenName));
                        }

                        writer.WriteEndElement();

                        contributorIndex++;
                    }
                    writer.WriteEndElement();

                    writer.WriteElementString("abstracts", article.Abstract.JournalText());
                    //writer.WriteElementString("english_abstracts", article.JournalText());
                    //writer.WriteElementString("spanish_abstracts", article.Title.JournalText());

                    writer.WriteStartElement("keywords");

                    int keywordIndex = 0;
                    foreach (string keyword in article.Keywords)
                    {

                        writer.WriteStartElement("keyword");
                        writer.WriteAttributeString("key", string.Format("kywrd{0}-{1}", articleIndex, keywordIndex));
                        writer.WriteString(keyword.JournalText());

                        writer.WriteEndElement();

                        keywordIndex++;
                    }
                    writer.WriteEndElement();

                    writer.WriteStartElement("pages");
                    writer.WriteElementString("first_page", article.FirstPage.ToString());
                    writer.WriteElementString("last_page", article.LastPage.ToString());

                    writer.WriteEndElement();

                    if (article.Citations != null)
                    {
                        writer.WriteStartElement("citation_list");

                        int citationIndex = 0;

                        foreach (string citation in article.Citations)
                        {

                            writer.WriteStartElement("citation");
                            writer.WriteAttributeString("key", string.Format("cttn{0}-{1}", articleIndex, citationIndex));
                            writer.WriteString(citation.JournalText());

                            writer.WriteEndElement();

                            citationIndex++;
                        }
                    }

                    writer.WriteEndElement();
                    writer.WriteEndElement();
                    articleIndex++;
                }


                writer.WriteEndElement();
            }

            return FormatXml(xmlJournalDocument);
        }

        public static string JournalText(this string value)
        {
            if (!string.IsNullOrEmpty(value))
            {
                return value;
            }
            else
            {
                return string.Empty;
            }
        }

        private static string FormatXml(XmlDocument xmlDocument)
        {
            StringBuilder stringBuilder = new StringBuilder();
            StringWriter stringWriter = new StringWriter(stringBuilder);
            XmlTextWriter xmlTextWriter = null;
            try
            {
                xmlTextWriter = new XmlTextWriter(stringWriter);
                xmlTextWriter.Formatting = Formatting.Indented;
                xmlDocument.WriteTo(xmlTextWriter);
            }
            finally
            {
                if (xmlTextWriter != null)
                    xmlTextWriter.Close();
            }

            return stringBuilder.ToString();
        }

        private static string SanitizeXmlString(string xml)
        {
            if (xml == null)
            {
                throw new ArgumentNullException("xml");
            }

            StringBuilder buffer = new StringBuilder(xml.Length);

            foreach (char c in xml)
            {
                if (IsLegalXmlChar(c))
                {
                    buffer.Append(c);
                }
            }

            return buffer.ToString();
        }

        private static bool IsLegalXmlChar(int character)
        {
            return
            (
                 character == 0x9 /* == '\t' == 9   */          ||
                 character == 0xA /* == '\n' == 10  */          ||
                 character == 0xD /* == '\r' == 13  */          ||
                (character >= 0x20 && character <= 0xD7FF) ||
                (character >= 0xE000 && character <= 0xFFFD) ||
                (character >= 0x10000 && character <= 0x10FFFF)
            );
        }

        internal static void WritetoExcel(IList<Article> ArticlesData, string folderName)
        {
            try
            {
                var excelApp = new Microsoft.Office.Interop.Excel.Application();
                var workbook = excelApp.Workbooks.Add(1);



                WriteArticlesSheet(ArticlesData, workbook);

                object misValue = System.Reflection.Missing.Value;
                workbook.SaveAs(folderName + "/Excel.xls", Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                workbook.Close(true, misValue, misValue);
                excelApp.Quit();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private static void WriteArticlesSheet(IList<Article> ArticlesData, Microsoft.Office.Interop.Excel.Workbook workbook)
        {

            Microsoft.Office.Interop.Excel.Worksheet wsArticlesInfo = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Sheets.Add();
            wsArticlesInfo.Name = "Articles";

            Microsoft.Office.Interop.Excel.Range rangeat1 = wsArticlesInfo.Cells[1, 1] as Range;
            rangeat1.Value2 = "File Name";

            Microsoft.Office.Interop.Excel.Range rangeat2 = wsArticlesInfo.Cells[1, 2] as Range;
            rangeat2.Value2 = "Title";

            Microsoft.Office.Interop.Excel.Range rangeat3 = wsArticlesInfo.Cells[1, 3] as Range;
            rangeat3.Value2 = "English Title";

            Microsoft.Office.Interop.Excel.Range rangeat4 = wsArticlesInfo.Cells[1, 4] as Range;
            rangeat4.Value2 = "Spanish Title";

            Microsoft.Office.Interop.Excel.Range rangeat5 = wsArticlesInfo.Cells[1, 5] as Range;
            rangeat5.Value2 = "Contributors";

            Microsoft.Office.Interop.Excel.Range rangeat6 = wsArticlesInfo.Cells[1, 6] as Range;
            rangeat6.Value2 = "Keywords";

            Microsoft.Office.Interop.Excel.Range rangeat7 = wsArticlesInfo.Cells[1, 7] as Range;
            rangeat7.Value2 = "English Keywords";

            Microsoft.Office.Interop.Excel.Range rangeat8 = wsArticlesInfo.Cells[1, 8] as Range;
            rangeat8.Value2 = "Spanish Keywords";

            Microsoft.Office.Interop.Excel.Range rangeat9 = wsArticlesInfo.Cells[1, 9] as Range;
            rangeat9.Value2 = "First Page";

            Microsoft.Office.Interop.Excel.Range rangeat10 = wsArticlesInfo.Cells[1, 10] as Range;
            rangeat10.Value2 = "Last Page";

            Microsoft.Office.Interop.Excel.Range rangeat11 = wsArticlesInfo.Cells[1, 11] as Range;
            rangeat11.Value2 = "Citations";

            int index = 2;

            foreach (Article article in ArticlesData)
            {
                Microsoft.Office.Interop.Excel.Range range1 = wsArticlesInfo.Cells[index, 1] as Range;
                range1.Value2 = article.Filename;

                Microsoft.Office.Interop.Excel.Range range2 = wsArticlesInfo.Cells[index, 2] as Range;
                range2.Value2 = article.Title;

                Microsoft.Office.Interop.Excel.Range range3 = wsArticlesInfo.Cells[index, 3] as Range;
                range3.Value2 = article.EnglishTitle;

                Microsoft.Office.Interop.Excel.Range range4 = wsArticlesInfo.Cells[index, 4] as Range;
                range4.Value2 = article.SpanishTitle;

                Microsoft.Office.Interop.Excel.Range range5 = wsArticlesInfo.Cells[index, 5] as Range;
                article.Contributors.ToList().ForEach(c => range5.Value2 += string.Format("{0} {1}\r", c.Surname, c.GivenName));

                Microsoft.Office.Interop.Excel.Range range6 = wsArticlesInfo.Cells[index, 6] as Range;
                range6.Value2 = string.Join(";", article.Keywords);

                Microsoft.Office.Interop.Excel.Range range7 = wsArticlesInfo.Cells[index, 7] as Range;
                range7.Value2 = string.Join(";", article.EnglishKeywords);

                Microsoft.Office.Interop.Excel.Range range8 = wsArticlesInfo.Cells[index, 8] as Range;
                range8.Value2 = string.Join(";", article.SpanishKeywords);

                Microsoft.Office.Interop.Excel.Range range9 = wsArticlesInfo.Cells[index, 9] as Range;
                range9.Value2 = article.FirstPage;

                Microsoft.Office.Interop.Excel.Range range10 = wsArticlesInfo.Cells[index, 10] as Range;
                range10.Value2 = article.LastPage;

                Microsoft.Office.Interop.Excel.Range range11 = wsArticlesInfo.Cells[index, 11] as Range;
                range11.Value2 = string.Join("\r", article.Citations);

                index++;
            }
        }
    }
}

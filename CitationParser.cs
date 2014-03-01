using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace ScrapDocument
{
    static class CitationParser
    {
        public static Citation ParseCitation(string input)
        {
            Citation citation = new Citation();
            input = TrimStartingNumers(input);

            GetDOI(input, ref citation);

            if (string.IsNullOrEmpty(citation.Doi))
            {
                try
                {
                    string[] dotArray = input.Split(new char[1] { '.' }, StringSplitOptions.RemoveEmptyEntries);

                    GetAuthorName(dotArray, ref citation);
                    GetYear(dotArray, ref citation);
                    GetTitle(dotArray, ref citation);
                    GetPages(dotArray, ref citation);

                    citation.UnstructuredData = input;
                }
                catch (Exception ex)
                {
                    citation.UnstructuredData = input;
                }
            }

            return citation;
        }

        private static string TrimStartingNumers(string input)
        {
            string output = string.Empty;
            Regex expression = new Regex("[\\[][0-9]{1,2}[\\]]");
            Match match = expression.Match(input);

            if (match.Success)
            {
                output = input.Replace(match.Value, string.Empty).Trim();
            }
            else
            {
                output = input;
            }

            return output;
        }

        private static void GetPages(string[] dotArray, ref Citation citation)
        {
            foreach (string element in dotArray)
            {
                if (element.Contains("pages"))
                {
                    citation.FirstPage = Regex.Replace(element, @"[^\d]", "");
                    break;
                }
                else
                {
                    citation.FirstPage = GetMatch(element, "([0-9]?\\d{1,})(?:[ ]{0,1}-[ ]{0,1}([0-9]?\\d{1,}))");

                    if (!string.IsNullOrEmpty(citation.FirstPage))
                    {
                        citation.FirstPage = citation.FirstPage.Substring(0, citation.FirstPage.IndexOf('-')).Trim();
                        break;
                    }
                }
            }

            if (string.IsNullOrEmpty(citation.FirstPage))
            {
                foreach (string element in dotArray)
                {
                    int page;
                    bool result = Int32.TryParse(element.Trim(' ', ',', ':'), out page);

                    if (result)
                    {
                        citation.FirstPage = page.ToString();
                        break;
                    }
                }
            }
        }

        private static void GetTitle(string[] dotArray, ref Citation citation)
        {
            int titleIndex = GetTitleIndex(dotArray, 1);


            if (!citation.Type.Equals(CitationType.Conference))
            {
                if (titleIndex >= dotArray.Length)
                {
                    citation.VolumeTitle = dotArray[titleIndex - 1].Trim();
                    dotArray[titleIndex-1] = string.Empty;
                }
                else
                {
                    citation.VolumeTitle = dotArray[titleIndex].Trim();
                    dotArray[titleIndex] = string.Empty;
                }
            }
        }

        private static int GetTitleIndex(string[] dotArray, int titleIndex)
        {
            int newIndex = titleIndex;

            if (titleIndex < dotArray.Length && titleIndex <= dotArray.Length / 2)
            {
                if (dotArray[titleIndex].Contains(";") || dotArray[titleIndex].Contains(",") || (!string.IsNullOrEmpty(dotArray[titleIndex]) && dotArray[titleIndex].Length <= 2))
                {
                    newIndex = titleIndex + 1;
                    GetTitleIndex(dotArray, newIndex);
                }
            }

            return newIndex;
        }

        private static void GetYear(string[] dotArray, ref Citation citation)
        {
            for (int index = dotArray.Length - 1; index >= 0; index--)
            {
                // First find year number as combination
                citation.CYear = GetMatch(dotArray[index], "((18|19|20)\\d{2})(?:[ ]{0,1}-[ ]{0,1}((18|19|20)\\d{2}))");

                if (string.IsNullOrEmpty(citation.CYear))
                {
                    //Find four digit year no.
                    citation.CYear = GetMatch(dotArray[index], "(18|19|20)\\d{2}");

                    if (!string.IsNullOrEmpty(citation.CYear))
                    {
                        //dotArray[index] = string.Empty;
                        break;
                    }
                }
                else
                {
                    //dotArray[index] = string.Empty;
                    break;
                }
            }
        }

        private static void GetAuthorName(string[] dotArray, ref Citation citation)
        {
            int indexofName = dotArray[0].IndexOf(',');

            if (indexofName >= 0)
            {
                citation.AuthorName = dotArray[0].Substring(0, indexofName);
            }
            else
            {
                citation.AuthorName = dotArray[0];
            }

            // If surname name is incorrect
            if (citation.AuthorName.Length <= 1)
            {
                citation.AuthorName = string.Empty;
            }

            dotArray[0] = string.Empty;
        }

        private static void GetDOI(string input, ref Citation citation)
        {
            citation.Doi = GetMatch(input, "\\b(10[.][0-9]{4,}(?:[.][0-9]+)*/(?:(?![\"&\'<>])\\S)+)\\b");
        }

        public static string GetMatch(string input, string pattern)
        {
            Regex expression = new Regex(pattern);
            Match match = expression.Match(input);

            if (match.Success)
            {
                return match.Value;
            }
            else
            {
                return null;
            }
        }
    }

}

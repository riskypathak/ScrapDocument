using System.Collections.Generic;

namespace ScrapDocument
{
    public class Journal
    {
        public IList<Article> Articles = null;
    }

    public class Article
    {

        public IList<Contributor> Contributors = new List<Contributor>();
        public string Abstract = string.Empty;

        public IList<string> Citations = new List<string>();
        public string Filename = string.Empty;

        public string Title = string.Empty;
        public IList<string> Keywords = new List<string>();

        public string EnglishTitle;
        public IList<string> EnglishKeywords = new List<string>();

        public string SpanishTitle;
        public IList<string> SpanishKeywords = new List<string>();

        public int FirstPage;
        public int LastPage;
    }

    public class Contributor
    {
        public string GivenName = string.Empty;
        public string Surname = string.Empty;
        public string ContributorAdresses = string.Empty;
        public ContributorInfo ContributorInfo = new ContributorInfo();
    }

    public class ContributorInfo
    {
        public string PhoneNo = string.Empty; //-- ilist
        public string EmailID = string.Empty;
    }
}

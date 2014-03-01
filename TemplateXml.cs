using System.Collections.Generic;
using System.Xml;
using System.Xml.Linq;
using System.Linq;

namespace ScrapDocument
{
    public class TemplateXml
    {
        public IList<string> ReadTemplate(string templateType)
        {
            IList<string> templateElements = new List<string>();

            XDocument xDoc = XDocument.Load("TemplateSchema.xml");

            var nodes = xDoc.Descendants("Template").Where(e => e.Attribute("name").Value.Equals(templateType)).Descendants();

            foreach (XElement element in nodes)
            {
                templateElements.Add(element.Value);
            }
            return templateElements;
        }

        public IList<string> GetAllTemplateTypes()
        {
            IList<string> templateElements = null;

            XDocument xDoc = XDocument.Load("TemplateSchema.xml");

            templateElements = xDoc.Descendants("Template").Select(e => e.Attribute("name").Value).ToList();

            return templateElements;


        }

    }
}

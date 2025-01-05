using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using mxd.Dwml;
using System.Xml;

namespace WordDocumentTableParserProject
{
    public class WordFormatter : IQuestionFormatter
    {
        private readonly XmlDocument _document = new();
        public string Format(TableCell cell, QuestionPart questionPart)
        {
            string x =  Format(cell);
            return x;
        }

        public string Format(TableCell cell)
        {
            StringBuilder result = new StringBuilder();

            // Iterate through paragraphs in the cell
            foreach (var paragraph in cell.Elements<Paragraph>())
            {
                // Recursively find all OfficeMath elements within the paragraph
                var mathElements = paragraph.Descendants<DocumentFormat.OpenXml.Math.OfficeMath>();
                foreach (var element in mathElements)
                {
                    // Convert OfficeMath to LaTeX
                    _document.LoadXml(element.OuterXml);
                    string latex = MLConverter.Convert(_document.DocumentElement);
                    result.Append(latex);
                }

                // Add non-math text from the paragraph
                var text = string.Join("", paragraph.Descendants<Run>().Select(r => r.InnerText));
                result.Append(text);

                //result.AppendLine(); // Add line breaks between paragraphs
            }

            return result.ToString();
        }



    }
}

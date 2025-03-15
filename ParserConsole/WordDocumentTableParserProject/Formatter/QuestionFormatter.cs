using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using mxd.Dwml;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using WordDocumentTableParserProject.Models;

namespace WordDocumentTableParserProject.Formatter
{
    public class QuestionFormatter : IQuestionFormatter
    {
        private readonly XmlDocument _document = new();
        private readonly string pattern = @"(_\{[^}]+\})_\{([^}]+)\}";
        private readonly string replacement = "_{$1$2}";

        public Question Format(RawQuestion rawQuestion)
        {
            var question = new Question
            {
                QuestionText = FormatQuestionText(rawQuestion.QuestionText),
                QuestionChoices = FormatQuestionChoices(rawQuestion.QuestionChoices),
                Answer = ExtractAnswer(rawQuestion.QuestionAnswer)
            };

            return question;
        }

        private List<List<QuestionSentence>> FormatQuestionChoices(OpenXmlElement questionChoices)
        {
            List<List<QuestionSentence>> questionSentences = new();
            foreach (Paragraph p in questionChoices.Elements<Paragraph>())
            {
                List<QuestionSentence> result = new();
                FormatElement(p, result);
                questionSentences.Add(result);
            }
            return questionSentences;
        }

        private List<QuestionSentence> FormatQuestionText(OpenXmlElement questionText)
        {
            var result = new List<QuestionSentence>();
            foreach (Paragraph paragraph in questionText.Elements<Paragraph>())
            {
                FormatElement(paragraph, result);
            }
            return result;
        }

        private void FormatElement(OpenXmlElement outerElement, List<QuestionSentence> result)
        {
                foreach (OpenXmlElement elem in outerElement.Elements())
                {
                    if (elem is Run run)
                    {
                        //should add use case for images in the future here
                        //we are assuming for now that each run element contains only text
                        result.Add(new QuestionSentence()
                        {
                            Text = string.Concat(run.Elements<Text>().Select(t => t.Text)),
                            QuestionSentenceType = QuestionSentenceType.SimpleText
                        });
                    }
                    else if (elem is DocumentFormat.OpenXml.Math.OfficeMath mathElement) 
                    {
                        result.Add(FormatMath(mathElement));
                    }
                    else if (elem is DocumentFormat.OpenXml.Math.Paragraph mathParagraph)
                    {
                        result.Add(FormatMathParagraph(mathParagraph));
                    }
                }
        }


        // Format mathematical equations into LaTeX
        private QuestionSentence FormatMath(DocumentFormat.OpenXml.Math.OfficeMath oMathElement)
        {
            string latex = LoadMathElement(oMathElement);
            return new QuestionSentence
            {
                Text = $"\\({latex}\\)",
                QuestionSentenceType = QuestionSentenceType.InlineEquation
            };
        }
        /// <summary>
        /// Formatting MathParagraph
        /// </summary>
        /// <param name="oMathElement"></param>
        /// <returns></returns>
        private QuestionSentence FormatMathParagraph(DocumentFormat.OpenXml.Math.Paragraph oMathElement)
        {
            StringBuilder builder = new();
            foreach (DocumentFormat.OpenXml.Math.OfficeMath mathElem in oMathElement.Elements<DocumentFormat.OpenXml.Math.OfficeMath>())
            {
                string latex = LoadMathElement(mathElem);
                builder.Append(latex);
            }
            return new QuestionSentence
            {
                Text = $"\\({builder}\\)",
                QuestionSentenceType = QuestionSentenceType.ParagraphEquation
            };
        }

        private string LoadMathElement(OpenXmlElement oMathElement)
        {
            _document.LoadXml(oMathElement.OuterXml);
            string latex = MLConverter.Convert(_document.DocumentElement);
            latex = Regex.Replace(latex, pattern, replacement);
            return latex;
        }

        // Extract the answer from the answer element
        private string ExtractAnswer(OpenXmlElement answerElement)
        {
            return answerElement.InnerText; // Assuming the answer is plain text
        }

        // TODO: Implement formatter for images or other complex elements
        private QuestionSentence FormatImage(Drawing drawingElement)
        {
            // Extract image data (e.g., base64 or URL) and return as a QuestionSentence
            return new QuestionSentence
            {
                Text = "ImagePlaceholder", // Replace with actual image data
                QuestionSentenceType = QuestionSentenceType.ImageUrl,
                AltText = "Image description" // Replace with actual alt text
            };
        }
    }
}
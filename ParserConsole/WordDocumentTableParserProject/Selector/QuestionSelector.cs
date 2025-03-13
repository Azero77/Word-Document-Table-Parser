using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WordDocumentTableParserProject.Models;
using WordDocumentTableParserProject.QuestionParsers;
using WordDocumentTableParserProject.Formatter;

namespace WordDocumentTableParserProject.Selector
{
    public class QuestionSelector : IQuestionSelector
    {
        private readonly IQuestionParser _parser;
        private readonly Formatter.IQuestionFormatter _formatter;
        public QuestionSelector(IQuestionParser parser,
                                Formatter.IQuestionFormatter formatter)
        {
            _parser = parser;
            _formatter = formatter;
        }

        public IEnumerable<Question> Process(WordprocessingDocument document)
        {
            var rawQuestions = _parser.ProcessDocument(document);
            foreach (RawQuestion rawQuestion in rawQuestions)
            {
                yield return _formatter.Format(rawQuestion);
            }
        }
    }
}

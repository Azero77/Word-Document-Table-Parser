using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WordDocumentTableParserProject.Models;

namespace WordDocumentTableParserProject.QuestionParsers
{
    public interface IQuestionParser : IDisposable
    {
        IEnumerable<RawQuestion> ProcessDocument(WordprocessingDocument document);
    }
}

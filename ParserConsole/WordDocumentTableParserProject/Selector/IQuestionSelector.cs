using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordDocumentTableParserProject.Selector
{
    public interface IQuestionSelector
    {
        IEnumerable<Question> Process(WordprocessingDocument document);
    }
}

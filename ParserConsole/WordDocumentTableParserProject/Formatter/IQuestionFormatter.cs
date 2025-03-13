using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WordDocumentTableParserProject.Models;

namespace WordDocumentTableParserProject.Formatter
{
    public interface IQuestionFormatter
    {
        Question Format(RawQuestion rawQuestion);
    }
}

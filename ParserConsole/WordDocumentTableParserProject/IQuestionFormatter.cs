using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordDocumentTableParserProject
{
    /// <summary>
    /// Formatter for every sentence in the word document to format specific on some arguments
    /// </summary>
    public interface IQuestionFormatter
    {
        string Format(TableCell cell,QuestionPart questionPart);
        string Format(TableCell cell);
    }

    public enum QuestionPart
    {
        Number,
        Text,
        Choices,
        Answer
    }

}

using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordDocumentTableParserProject
{
    public class JsonFormatter : IQuestionFormatter
    {

        public string Format(TableCell cell, QuestionPart questionPart)
        {
            return Format(cell);
        }

        public string Format(TableCell cell)
        {
            return cell.InnerText.Trim();
        }
    }
}

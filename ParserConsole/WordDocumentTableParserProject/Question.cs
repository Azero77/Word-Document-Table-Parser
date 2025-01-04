using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordDocumentTableParserProject
{
    public class Question
    {
        public List<string> Choices { get; set; } = new List<string>();
        public required string QuestionText { get; set; }
        public int QuestionAnswerIndex { get; set; }

    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static WordDocumentTableParserProject.Ai.QuizAIModel;

namespace WordDocumentTableParserProject.Ai
{
    internal interface IAIModel
    {
        public Task<AIModelResult> GetResponseAsync(string prompt);
    }
}

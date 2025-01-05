using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;

namespace WordDocumentTableParserProject
{
    //Adapter For writer and questionParser
    public class WordDocumentParser : IFileParser
    {
        private readonly string _documentPath;
        private readonly IWriter _writer;
        private QuestionParser? _questionParser;

        private WordDocumentParser(string documentPath, IWriter writer)
        {
            _documentPath = documentPath;
            _writer = writer;
        }

        public static WordDocumentParser LoadWordDocumentParser(string documentPath,IWriter writer)
        {
            WordDocumentParser wordDocumentParser = new(documentPath,writer);
            wordDocumentParser.LoadQuestionParser();
            return wordDocumentParser;
        }

        private void LoadQuestionParser()
        {
            WordprocessingDocument document = WordprocessingDocument.Open(_documentPath, false);
            _questionParser = new(document);
            
        }

        public async Task ParseAsync()
        {
            //Every QUESTION parsed will be writting to a file by the writer
            var questions = _questionParser?.GetQuestions()!;
            foreach (var question in questions)
            {
                await _writer.WriteAsync(question);
            }

            await _writer.DisposeAsync();
        }
    }
}

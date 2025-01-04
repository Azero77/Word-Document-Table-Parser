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
    public class WordDocumentParser : IFileParser
    {
        private readonly string _documentPath;
        private readonly IWriter _writer;
        private QuestionParser? _questionParser;

        public WordDocumentParser(string documentPath, IWriter writer)
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
            var list = _questionParser?.ParseQuestions(_documentPath);
            await Console.Out.WriteLineAsync(JsonSerializer.Serialize(list));
        }
    }
}

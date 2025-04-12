using DocumentFormat.OpenXml.Packaging;
using Microsoft.Extensions.Configuration;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WordDocumentTableParserProject.Ai;
using WordDocumentTableParserProject.Models;

namespace WordDocumentTableParserProject.QuestionParsers
{
    /// <summary>
    /// This parser work by reading the document and writing into .json file with the schema similiar to Raw question 
    /// </summary>
    public class AIQuestionParser : IQuestionParser
    {
        private readonly IAIModel _model;
        private readonly JsonSerializer _serializer = new() { Formatting = Formatting.Indented };

        public AIQuestionParser(IAIModel model)
        {
            _model = model;
        }

        public void Dispose()
        {
        }

        public async IAsyncEnumerable<RawQuestion> ProcessDocument(WordprocessingDocument document)
        {
            string path = await Convert(document);
            //reading json objects that represents a raw question
            StreamReader text_reader = new(path);
            JsonTextReader jsonReader = new(text_reader) { SupportMultipleContent = false};
            //Deserializing the object
            if ((await jsonReader.ReadAsync()) && jsonReader.TokenType == JsonToken.StartArray)
            {
                if (jsonReader.TokenType == JsonToken.EndArray)
                {
                    DeleteFile(path);
                    Dispose();
                    yield break;
                }

                RawQuestion? rawQuestion = _serializer.Deserialize<RawQuestion>(jsonReader);
                if (rawQuestion is null)
                    yield break;
                yield return rawQuestion;
            }
            yield break;
        }

        private void DeleteFile(string path)
        {
        }

        protected virtual async Task<string> Convert(WordprocessingDocument document)
        {
            //Reading the file
            string prompt = await CreatePrompt(document);
            var stream = CreateFile(out string path);
            await _model.WriteStreamingResponseAsync(stream, prompt);
            return path;
        }

        private static async Task<string> CreatePrompt(WordprocessingDocument document)
        {
            string document_file = await Read(document);
            string questionPrompt = GetPrompt();
            return questionPrompt + document_file;
        }

        private FileStream CreateFile(out string path)
        {
            string fileName = DateTime.Now.ToString() + "-"+ Guid.NewGuid().ToString() + ".json";
            path = Directory.GetCurrentDirectory() + fileName;
            return File.Create(path);
        }

        private static Task<string> Read(WordprocessingDocument document)
        {
            var reader = new StreamReader(document?.MainDocumentPart?.GetStream() ?? throw new InvalidDataException());
            return reader.ReadToEndAsync();
        }
        private static string GetPrompt()
        {
            return @"I want you to read the .docx which contains multiple choice questions, i want you to extract the questions found
                    and fill in a json format that represtents a list of questions object where each value is the open xml element of it (it can be a table cell or a run element) that look like this :
            ----
             [
                {
                    QuestionText : [Open xml for the question]
                    QuestionChoices : [ choice , choice, chioce] # each choice is an open xml for the question choice
                    Answer : [open xml for the answer]
                },
                {
                    Another Question
                }
            ]
            ----
            some instructions to keep in mind while parsing the .docx file:
            - i want each open xml to have the full information about the question part talked about,

            here is the file:
            
            ";
            
        }
    }
}

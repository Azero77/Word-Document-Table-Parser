using DocumentFormat.OpenXml.Office2013.Excel;
using mxd.Dwml;
using Newtonsoft.Json;
using System.Xml;
using System.Xml.Linq;
using WordDocumentTableParserProject;
using WordDocumentTableParserProject.Formatter;
using WordDocumentTableParserProject.QuestionParsers;
using WordDocumentTableParserProject.Selector;
using WordDocumentTableParserProject.WordFileParser;
using WordDocumentTableParserProject.Writer;

namespace ParserConsole
{
    internal class Program
    {
        static async Task Main(string[] args)
        {
            /*if (args.Length != 2)
            {
                Console.WriteLine("must specify document route and json file route");
                return;
            }*/
            string word_document_path = @"E:\Chemistry\Bank\EquillibrumBank.docx";
            string json_file_path ="test.json";

            WordDocumentParser parser = new(word_document_path,new JsonFileWriter(json_file_path),new QuestionSelector(
                new AnasQuestionParser(),
                new QuestionFormatter()
                ));
            Console.WriteLine("Loading...");
            await parser.ParseAsync();
            Console.WriteLine("Fininshed Converting...");
            Console.Read();
        }
    }
}

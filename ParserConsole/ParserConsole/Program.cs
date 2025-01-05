using mxd.Dwml;
using System.Xml;
using System.Xml.Linq;
using WordDocumentTableParserProject;

namespace ParserConsole
{
    internal class Program
    {
        static async Task Main(string[] args)
        {
            if (args.Length != 2)
            {
                Console.WriteLine("must specify document route and json file route");
                return;
            }
            string word_document_path = args[0];
            string json_file_path = args[1];
            WordDocumentParser parser = WordDocumentParser.LoadWordDocumentParser(word_document_path, new JsonFileWriter(json_file_path));
            Console.WriteLine("Loading...");
            await parser.ParseAsync();
            Console.Clear();
            Console.WriteLine("Fininshed Converting...");
            Console.Read();
        }
    }
}

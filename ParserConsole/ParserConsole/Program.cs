using mxd.Dwml;
using System.Xml;
using System.Xml.Linq;
using WordDocumentTableParserProject;

namespace ParserConsole
{
    internal class Program
    {
        static void Main(string[] args)
        {
            WordDocumentParser parser = WordDocumentParser.LoadWordDocumentParser(@"E:\Chemistry\Bank\EquillibrumBank.docx", new JsonFileWriter("test.json"));
            parser.ParseAsync().Wait();
            Console.WriteLine("Converstion Worked Successfully");
            Console.Read();
        }
    }
}

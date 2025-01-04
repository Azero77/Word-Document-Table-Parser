using System.Xml.Linq;
using WordDocumentTableParserProject;

namespace ParserConsole
{
    internal class Program
    {
        static void Main(string[] args)
        {
            /*WordDocumentParser parser = WordDocumentParser.LoadWordDocumentParser(@"E:\Chemistry\Bank\EquillibrumBank.docx",new JsonWriter("test.json"));
            parser.ParseAsync().Wait();*/

            XElement element = new XElement("LastName", "Zaggar");
            element.Add(new XAttribute("id", 123));
            element.Add(new XComment("Test Comment"));
            XElement client = new XElement("Client");
            client.Add(element);
            client.Save(File.Open("test.xml",FileMode.Open));

            Console.Read();
        }
    }
}

using DocumentFormat.OpenXml.Office2013.Excel;
using Microsoft.Extensions.Configuration;
using mxd.Dwml;
using Newtonsoft.Json;
using System.Net.Http.Headers;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using WordDocumentTableParserProject;
using WordDocumentTableParserProject.Ai;
using WordDocumentTableParserProject.Formatter;
using WordDocumentTableParserProject.QuestionParsers;
using WordDocumentTableParserProject.Selector;
using WordDocumentTableParserProject.WordFileParser;
using WordDocumentTableParserProject.Writer;
using static WordDocumentTableParserProject.Ai.QuizAIModel;

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
            /* string word_document_path = @"E:\Chemistry\Bank\EquillibrumBank.docx";
             string json_file_path ="test.json";

             WordDocumentParser parser = new(word_document_path,new JsonFileWriter(json_file_path),new QuestionSelector(
                 new AnasQuestionParser(),
                 new QuestionFormatter()
                 ));
             Console.WriteLine("Loading...");
             await parser.ParseAsync();
             Console.WriteLine("Fininshed Converting...");
             Console.Read();*/
            Console.Write("Enter path to .docx file: ");
            string path = Console.ReadLine();

            if (!File.Exists(path))
            {
                Console.WriteLine("File does not exist.");
                return;
            }
            IConfiguration configuration = new ConfigurationBuilder()
           .AddJsonFile("config.json", optional: true)
           .AddEnvironmentVariables()
           .Build();

            using var httpClient = new HttpClient();

            // Initialize the AI model
            var aiModel = new QuizAIModel(httpClient, configuration);
            string prompt = QuizAIModel.GetPrompt();
            // Read file bytes
            byte[] fileBytes = File.ReadAllBytes(path);

            prompt = fileBytes + " " + prompt;
            var result = test(prompt);

            await foreach (var chunk in result)
            {
                await Task.Delay(10);
                Console.Write(chunk);
            }
        }


        private static async IAsyncEnumerable<string> test(string prompt)
        {
            var _client = new HttpClient();
            _client.BaseAddress = new Uri("https://api.groq.com/openai/v1/chat/completions");
            _client.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", "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");
            var requestBody = new
            {
                messages = new[]
               {
                    new { role = "user", content = prompt }
                },
                model = "meta-llama/llama-4-scout-17b-16e-instruct",
                temperature = 1,
                max_completion_tokens = 1024,
                top_p = 1,
                stream = true, // stream = true is harder to implement without event source, keep false for now
                stop = (string?)null
            };

            var json = JsonConvert.SerializeObject(requestBody);
            var content = new StringContent(json, Encoding.UTF8, "application/json");

            var request = new HttpRequestMessage(HttpMethod.Post, "")
            {
                Content = content
            };
            request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("text/event-stream"));

            var response = await _client.SendAsync(request, HttpCompletionOption.ResponseHeadersRead);

            response.EnsureSuccessStatusCode();
            Stream stream = await response.Content.ReadAsStreamAsync();

            using StreamReader reader = new StreamReader(stream);

            while (!reader.EndOfStream)
            {
                var line = await reader.ReadLineAsync();
                if (string.IsNullOrWhiteSpace(line)) continue;
                if (!line.StartsWith("data:")) continue;

                var jsonPart = line.Substring("data:".Length).Trim();
                if (jsonPart == "[DONE]") yield break;

                ChatCompletionResponse? completion = JsonConvert.DeserializeObject<ChatCompletionResponse>(jsonPart);
                var messageContent = completion?.Choices?.FirstOrDefault()?.Delta?.Content;
                if (!string.IsNullOrWhiteSpace(messageContent))
                    yield return messageContent ?? string.Empty;
            }
        }
        class ChatCompletionResponse
        {
            [JsonProperty("choices")]
            public List<Choice> Choices { get; set; } = new();
        }

        class Choice
        {
            [JsonProperty("message")]
            public Message Message { get; set; } = new();

            [JsonProperty("delta")]
            public Delta Delta { get; set; } = new();
        }

        public class Message
        {
            [JsonProperty("role")]
            public string Role { get; set; } = string.Empty;

            [JsonProperty("content")]
            public string Content { get; set; } = string.Empty;
        }
        class Delta
        {
            [JsonProperty("content")]
            public string? Content { get; set; }
        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using System.Net.Http.Headers;
using Microsoft.Extensions.Configuration;
namespace WordDocumentTableParserProject.Ai
{
    public partial class QuizAIModel : IAIModel
    {
        private readonly string _apiKey = string.Empty;
        private readonly string _baseUrl = string.Empty;
        private readonly HttpClient _client;
        private readonly IConfiguration _configuration;

        public QuizAIModel(HttpClient client,IConfiguration configuration)
        {
            _client = client;
            _configuration = configuration;
            _apiKey = configuration?["ApiKey"] ?? string.Empty;
            _baseUrl = configuration?["BaseUrl"] ?? string.Empty;
            _client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", _apiKey);
            _client.BaseAddress = new Uri(_baseUrl);
        }

        public async Task<AIModelResult> GetResponseAsync(string prompt)
        {
            //var requestBody = new { prompt };
            var requestBody = new
            {
                messages = new[] { new { role = "user", content = prompt } },
                model = _configuration?["model"] ?? string.Empty,
                temperature = 1,
                max_completion_tokens = _configuration?["max_completion_tokens"] ?? string.Empty,
                top_p = _configuration?["top_p"] ?? string.Empty,
                stream = false, //TODO we have to add a feature to allow streamed answers 
                stop = (string?)null
            };

            string json = JsonConvert.SerializeObject(requestBody);
            StringContent content = new StringContent(json, Encoding.UTF8, "application/json");

            try
            {
                var response = await _client.PostAsync("", content);
                response.EnsureSuccessStatusCode();

                var responseBody = await response.Content.ReadAsStringAsync();
                var data = JsonConvert.DeserializeObject<ChatCompletionResponse>(responseBody);
                var resultText = data?.Choices?.FirstOrDefault()?.Message?.Content;

                if (string.IsNullOrWhiteSpace(resultText))
                    return AIModelResult.Fail("Empty response from AI");

                return AIModelResult.Success(resultText);
            }
            catch (Exception ex)
            {
                return AIModelResult.Fail($"Request failed: {ex.Message}");
            }
        }

        public async IAsyncEnumerable<string> GetStreamingResponseAsync(string prompt)
        {
            var requestBody = new
            {
                messages = new[] { new { role = "user", content = prompt } },
                model = _configuration?["model"] ?? string.Empty,
                temperature = 1,
                max_completion_tokens = int.TryParse(_configuration?["max_completion_tokens"], out var tokens) ? tokens : 1024,
                top_p = float.TryParse(_configuration?["top_p"], out var topP) ? topP : 1f,
                stream = true,
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

        class Message
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

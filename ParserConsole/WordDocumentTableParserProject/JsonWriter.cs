using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;

namespace WordDocumentTableParserProject
{
    /// <summary>
    /// Writing .docx file into .json file question by question
    /// </summary>
    public class JsonWriter : IWriter
    {
        
        private readonly FileStream _file;

        public JsonWriter(string path)
        {
            _file = File.Create(path);
        }

        public JsonWriter(FileStream file)
        {
            _file = file;
        }

        public Task WriteAsync(Question question)
        {
            return JsonSerializer.SerializeAsync<Question>(_file,question);
        }


        public ValueTask DisposeAsync()
        {
            return _file.DisposeAsync();
        }
    }
}

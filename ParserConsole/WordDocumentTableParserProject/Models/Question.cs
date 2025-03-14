using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.Json.Serialization;
using System.Threading.Tasks;

namespace WordDocumentTableParserProject
{
    public class Question
    {
        public List<List<QuestionSegment>> QuestionChoices { get; set; } = null!; //every choice is a list of Question segment
        public required List<QuestionSegment> QuestionText { get; set; }
        public string Answer { get; set; } = string.Empty;
    }

    public class QuestionSegment
    {
        public List<QuestionSentence> Sentences = null!;
    }

    public class QuestionSentence
    {
        public string Text { get; set; } = string.Empty;
        public QuestionSentenceType QuestionSentenceType { get; set; } = QuestionSentenceType.SimpleText;

        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public string? AltText { get; set; } = null;

    }

    public enum QuestionSentenceType
    {
        SimpleText,
        Equation,
        ImageUrl,
        CodeBlock,
        Table
    }
}

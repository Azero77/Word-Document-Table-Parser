using DocumentFormat.OpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordDocumentTableParserProject.Models
{
    public class RawQuestion
    {
        public OpenXmlElement QuestionText { get; set; } = null!;
        public OpenXmlElement QuestionChoices { get; set; } = null!;
        public OpenXmlElement QuestionAnswer { get; set; } = null!;
    }
}

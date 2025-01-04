using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Office.SpreadSheetML.Y2023.MsForms;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordDocumentTableParserProject
{
    /// <summary>
    /// Take the question from the table and lazily load it
    /// </summary>
    public class QuestionParser
    {
        private readonly WordprocessingDocument _document;
        public QuestionParser(WordprocessingDocument document)
        {
            _document = document;
        }
        public List<Question> ParseQuestions(string filePath)
        {
            List<Question> questions = new();

            using (_document)
            {
                var body = _document.MainDocumentPart.Document.Body;
                var table = body?.Elements<Table>().Single();
                var rows = table?.Elements<TableRow>();

                    // Process rows in pairs (odd row + even row)
                for (int i = 0; i < 12; i += 2)
                {
                    var oddRow = rows.ElementAt(i);
                    var evenRow = i + 1 < rows.Count() ? rows.ElementAt(i + 1) : null;

                    if (oddRow != null && evenRow != null)
                    {
                        var oddCells = oddRow.Elements<TableCell>();
                        var evenCells = evenRow.Elements<TableCell>();

                        if (oddCells.Count() == 3 && evenCells.Count() == 3)
                        {
                            Question question = new()
                            {
                                QuestionText = GetTextFromCell(oddCells.ElementAt(1)),
                                QuestionAnswerIndex = (GetTextFromCell(evenCells.ElementAt(2)))[0] - 'A'
                            };

                            // Extract choices from the second column of the even row
                            var choicesParagraphs = evenCells.ElementAt(1).Elements<Paragraph>();
                            foreach (var paragraph in choicesParagraphs)
                            {
                                question.Choices.Add(paragraph.InnerText.Trim());
                            }

                            questions.Add(question);
                        }
                    }
                    }
                return questions;
            }
            string GetTextFromCell(TableCell cell)
            {
                return cell.InnerText.Trim();
            }
        }

        /// <summary>
        /// Generator for questions to lazy load every question
        /// </summary>
        /// <returns></returns>
        public IEnumerable<Question> GetQuestion()
        {
            throw new NotImplementedException();
        }

    }
}
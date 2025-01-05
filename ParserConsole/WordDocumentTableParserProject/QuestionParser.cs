using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Office.SpreadSheetML.Y2023.MsForms;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;

namespace WordDocumentTableParserProject
{
    /// <summary>
    /// Take the question from the table and lazily load it
    /// </summary>
    public class QuestionParser : IDisposable
    {
        private readonly WordprocessingDocument _document;
        private Table? _table;
        private IEnumerator<TableRow> _enumerator;
        public QuestionParser(WordprocessingDocument document)
        {
            _document = document;
            AssignTable();
            _enumerator = RowGenerator().GetEnumerator();
        }

        private void AssignTable()
        {
            IEnumerable<Table> tables = _document?.MainDocumentPart?.Document?.Body?.Elements<Table>() ?? throw new InvalidDataException("Invalid Document");
            if (tables.Count() != 1)
                throw new InvalidDataException("Document Contains zero or More than one table please select the current table by index or name");
            _table = tables.Single();
        }


        /// <summary>
        /// Generator for questions to lazy load every question
        /// </summary>
        /// <returns></returns>
        public IEnumerable<Question> GetQuestions()
        {
            TableRow evenRow, oddRow;
            int i = 0;
            while (i++ < 8)
            {
                if (!_enumerator.MoveNext()) yield break; // No more rows to process
                oddRow = _enumerator.Current;

                if (!_enumerator.MoveNext())
                    throw new InvalidDataException("Document must have an even number of rows for questions and answers.");

                evenRow = _enumerator.Current;

                yield return ProcessQuestion(evenRow, oddRow);
            }
        }

        private Question ProcessQuestion(TableRow evenRow, TableRow oddRow)
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
                return question;

            }
            throw new InvalidDataException("Must Have Three columns for each row");

            string GetTextFromCell(TableCell cell)
            {
                return cell.InnerText.Trim();
            }
        }

        private (TableRow even,TableRow odd) AssignRows()
        {
            TableRow even,odd;
            _enumerator.MoveNext();
            odd = _enumerator.Current;
            _enumerator.MoveNext();
            even = _enumerator.Current;
            return (even, odd);
        }

        /// <summary>
        /// Lazily load every table row 
        /// </summary>
        /// <returns></returns>
        private IEnumerable<TableRow> RowGenerator()
        {
           
            IEnumerable<TableRow> rows = _table?.Elements<TableRow>() ?? throw new InvalidDataException("No Rows Are Added");
            foreach (var row in rows)
            {
                yield return row;
            }
        }

        public void Dispose()
        {
            _document?.Dispose();
        }
    }
}
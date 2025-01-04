using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordDocumentTableParserProject
{
    public interface IFileParser
    {
        Task ParseAsync();
    }
}

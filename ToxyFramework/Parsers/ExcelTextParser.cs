using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.Extractor;
using NPOI.HSSF.Extractor;

namespace Toxy.Parsers
{
    public class ExcelTextParser:PlainTextParser
    {
        public ExcelTextParser(ParserContext context)
            : base(context)
        {
            
        }
        public override string Parse()
        {
            if (!File.Exists(Context.Path))
                throw new FileNotFoundException("File " + Context.Path + " is not found");

            IWorkbook workbook = WorkbookFactory.Create(Context.Path);

            bool extractHeaderFooter = false;
            if (Context.Properties.ContainsKey("IncludeHeaderFooter"))
            {
                extractHeaderFooter = Utility.IsTrue(Context.Properties["IncludeHeaderFooter"]);
            }
            bool showCalculatedResult = false;
            if (Context.Properties.ContainsKey("ShowCalculatedResult"))
            {
                showCalculatedResult = Utility.IsTrue(Context.Properties["ShowCalculatedResult"]);
            }
            bool includeSheetNames = true;
            if (Context.Properties.ContainsKey("IncludeSheetNames"))
            {
                includeSheetNames = Utility.IsTrue(Context.Properties["IncludeSheetNames"]);
            }
            bool includeComment = true;
            if (Context.Properties.ContainsKey("IncludeComments"))
            {
                includeComment = Utility.IsTrue(Context.Properties["IncludeComments"]);
            }

            if (workbook is XSSFWorkbook)
            {
                XSSFExcelExtractor extractor = new XSSFExcelExtractor((XSSFWorkbook)workbook);
                extractor.SetIncludeHeadersFooters(extractHeaderFooter);
                extractor.SetIncludeCellComments(includeComment);
                extractor.SetIncludeSheetNames(includeSheetNames);
                extractor.SetFormulasNotResults(!showCalculatedResult);
                return extractor.Text;
            }
            else //if (workbook is HSSFWorkbook)
            {
                ExcelExtractor extractor = new ExcelExtractor((HSSFWorkbook)workbook);
                extractor.IncludeHeaderFooter = extractHeaderFooter;
                extractor.IncludeCellComments= includeComment;
                extractor.IncludeSheetNames = includeSheetNames;
                extractor.FormulasNotResults = !showCalculatedResult;
                return extractor.Text;
            }
        }
    }
}

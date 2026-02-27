using NPOI.HSSF.Extractor;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.Extractor;
using NPOI.XSSF.UserModel;
using PasswordProtectedChecker;
using System.IO;

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
            Utility.ValidateContext(Context);
            
            var checker = new Checker();
            if (checker.IsFileProtected(Context.Path).Protected)
                throw new System.InvalidOperationException($"file {Context.Path} is encrypted");
                
            IWorkbook workbook = null;
            if (Context.IsStreamContext)
                workbook = WorkbookFactory.Create(Context.Stream);
            else
                workbook = WorkbookFactory.Create(Context.Path);
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

            if (workbook is XSSFWorkbook xssWorkbook)
            {
                XSSFExcelExtractor extractor = new XSSFExcelExtractor(xssWorkbook);
                extractor.IncludeHeadersFooters = extractHeaderFooter;
                extractor.IncludeCellComments = includeComment;
                extractor.IncludeSheetNames = includeSheetNames;
                extractor.FormulasNotResults = !showCalculatedResult;
                return extractor.Text;
            }
            else //if (workbook is HSSFWorkbook)
            {
                ExcelExtractor extractor = new ExcelExtractor((HSSFWorkbook)workbook);
                extractor.IncludeHeadersFooters = extractHeaderFooter;
                extractor.IncludeCellComments= includeComment;
                extractor.IncludeSheetNames = includeSheetNames;
                extractor.FormulasNotResults = !showCalculatedResult;
                return extractor.Text;
            }
        }
    }
}
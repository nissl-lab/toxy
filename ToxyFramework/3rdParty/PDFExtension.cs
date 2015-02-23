using PdfSharp.Pdf;
using PdfSharp.Pdf.Content;
using PdfSharp.Pdf.Content.Objects;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Toxy
{
    public static class PdfSharpExtensions
    {
        public static string ExtractWholeText(this PdfPage page)
        {
            var content = ContentReader.ReadContent(page);

            var result = new StringBuilder();
            CObjectVistor visitor = new CObjectVistor();
            visitor.ExtractText(content, result);
            return result.ToString();
        }
        public static List<string> ExtractText(this PdfPage page)
        {
            var content = ContentReader.ReadContent(page);
            
            var result = new StringBuilder();
            CObjectVistor visitor = new CObjectVistor();
            visitor.ExtractText(content,result);
            string[] list=result.ToString().Split(new string[]{"\n\n"},StringSplitOptions.RemoveEmptyEntries);
            return new List<string>(list);
        }

        public class CObjectVistor
        {
            #region CObject Visitor
            public void ExtractText(CObject obj, StringBuilder target)
            {
                if (obj is CArray)
                    ExtractText((CArray)obj, target);
                else if (obj is CComment)
                    ExtractText((CComment)obj, target);
                else if (obj is CInteger)
                    ExtractText((CInteger)obj, target);
                else if (obj is CName)
                    ExtractText((CName)obj, target);
                else if (obj is CNumber)
                    ExtractText((CNumber)obj, target);
                else if (obj is COperator)
                    ExtractText((COperator)obj, target);
                else if (obj is CReal)
                    ExtractText((CReal)obj, target);
                else if (obj is CSequence)
                    ExtractText((CSequence)obj, target);
                else if (obj is CString)
                    ExtractText((CString)obj, target);
                else
                    throw new NotImplementedException(obj.GetType().AssemblyQualifiedName);
            }

            private void ExtractText(CArray obj, StringBuilder target)
            {
                foreach (var element in obj)
                {
                    ExtractText(element, target);
                }
            }
            private double lastY = 0;
            private double lastX = 0;
            private double diffX = 0;
            private double diffY = 0;
            private string lastString = null;
            private void ExtractText(CComment obj, StringBuilder target) { /* nothing */ }
            private void ExtractText(CInteger obj, StringBuilder target) { /* nothing */ }
            private void ExtractText(CName obj, StringBuilder target) { /* nothing */ }
            private void ExtractText(CNumber obj, StringBuilder target) { /* nothing */ }
            private void ExtractText(COperator obj, StringBuilder target)
            {
                if (obj.OpCode.OpCodeName == OpCodeName.Tj || obj.OpCode.OpCodeName == OpCodeName.TJ)
                {
                    foreach (var element in obj.Operands)
                    {
                        ExtractText(element, target);
                    }
                    //target.Append(" ");
                }
                else if (obj.OpCode.OpCodeName == OpCodeName.Td || obj.OpCode.OpCodeName == OpCodeName.TD)
                {
                    target.Append("\n");
                }
                else if (obj.OpCode.OpCodeName == OpCodeName.Tm)
                {
                    double x=0;
                    if (obj.Operands[4] is CReal)
                    {
                        x = ((CReal)obj.Operands[4]).Value;
                    }
                    else if(obj.Operands[4] is CInteger)
                    {
                        x = (double)((CInteger)obj.Operands[4]).Value;
                    }
                    double y=0;
                    if (obj.Operands[5] is CReal)
                    {
                        y = ((CReal)obj.Operands[5]).Value;
                    }
                    else if (obj.Operands[5] is CInteger)
                    {
                        y = (double)((CInteger)obj.Operands[5]).Value;
                    }

                    diffX = x - lastX;

                    if (lastString != null)
                    {
                        target.Append(lastString);
                        if (lastString.Length == 1)
                        {
                            if ((lastString[0] >= 'A' && lastString[0] <= 'Z'))
                            {
                                if (lastString[0] == 'M')
                                {
                                    if (diffX > 11.3)
                                        target.Append(" ");
                                }
                                else
                                {
                                    if (diffX > 9.5)
                                        target.Append(" ");
                                }
                            }
                            else if ((lastString[0] >= 'a' && lastString[0] <= 'z'))
                            {
                                if (lastString[0] == 'm')
                                {
                                    if (diffX > 9.4)
                                    {
                                        target.Append(" ");
                                    }
                                }
                                else if (diffX > 7)
                                    target.Append(" ");
                            }
                            else if (diffX >8.1)
                                target.Append(" ");

                        }
                    }
                    diffY = lastY - y;
                    if (lastY != 0 && diffY > 13)
                    {
                        for (int i = 1; i < diffY / 13;i++ )
                            target.Append("\n");
                    }
                    lastX = x;
                    lastY = y;
                }
                else
                {

                }
            }
            private void ExtractText(CReal obj, StringBuilder target) { /* nothing */ }
            private void ExtractText(CSequence obj, StringBuilder target)
            {
                foreach (var element in obj)
                {
                    ExtractText(element, target);
                }
            }
            private void ExtractText(CString obj, StringBuilder target)
            {
                lastString = obj.Value;
            }
            #endregion        
        }

    }
}

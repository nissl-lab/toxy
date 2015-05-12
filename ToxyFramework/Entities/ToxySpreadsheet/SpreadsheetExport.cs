using System;
using System.Collections.Generic;
using System.Data;
using System.Text;
using Clifton.Tools.Data;
using System.IO;
using System.IO.Compression;
using ICSharpCode.SharpZipLib.BZip2;
using ICSharpCode.SharpZipLib.Checksums;
using ICSharpCode.SharpZipLib.Zip;
using ICSharpCode.SharpZipLib.GZip;

using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.HSSF.Util;
using NPOI.HSSF.UserModel;

namespace Toxy
{
    public partial class ToxySpreadsheet : ICloneable
    {

        public bool ToCSV(string filename, string Separatedstring = ",")
        {





            FileStream fs = File.Create(filename);

            StreamWriter sw = new StreamWriter(fs, Encoding.UTF8);


            foreach (ToxyTable tsheet in Tables)
            {

                sw.WriteLine("Tabel:" + tsheet.Name);
                if (tsheet.HasHeader)
                {
                    for (int j = 0; j <= tsheet.HeaderRows.Count; j++)
                    {
                        string row = string.Empty;
                        foreach (ToxyRow trow in tsheet.HeaderRows)
                        {
                            if (trow.RowIndex == j)
                            {
                                for (int k = 0; k <= trow.LastCellIndex; k++)
                                {
                                    bool blankcell = true;
                                    foreach (ToxyCell tcell in trow.Cells)
                                    {
                                        if (tcell.CellIndex == k)
                                        {
                                            if (k == trow.LastCellIndex)
                                            {
                                                row += "\"" + tcell.Value + "\"";
                                            }
                                            else
                                            {
                                                row += "\"" + tcell.Value + "\"" + Separatedstring;
                                            }
                                            blankcell = false;
                                        }
                                    }
                                    if (blankcell)
                                    {
                                        row += Separatedstring;

                                    }
                                }
                            }
                        }
                        sw.WriteLine(row);
                    }
                }

                for (int j = 0; j <= tsheet.LastRowIndex; j++)
                {
                    string row = string.Empty;

                    foreach (ToxyRow trow in tsheet.Rows)
                    {
                        if (trow.RowIndex == j)
                        {
                            for (int k = 0; k <= trow.LastCellIndex; k++)
                            {
                                bool blankcell = true;
                                foreach (ToxyCell tcell in trow.Cells)
                                {
                                    if (tcell.CellIndex == k)
                                    {
                                        if (k == trow.LastCellIndex)
                                        {
                                            row += "\"" + tcell.Value + "\"";
                                        }
                                        else
                                        {
                                            row += "\"" + tcell.Value + "\"" + Separatedstring;
                                        }
                                        blankcell = false;
                                    }
                                }
                                if (blankcell)
                                {
                                    row += Separatedstring;
                                }
                            }
                        }
                    }
                    sw.WriteLine(row);
                }


                sw.WriteLine("");
                sw.WriteLine("");

            }


            sw.Flush();
            sw.Close();

            return true;
        }


        public bool ToTXT(string filename)
        {

            ToCSV(filename, "\t");
            return true;

        }

        public bool ToXLS(string filename)
        {
            ToXLSX(filename, 2003);
            return false;
        }
        public bool ToXLSX(string filename, int ver = 2007)
        {


            IWorkbook workbook;
            if (ver == 2007)
                workbook = new XSSFWorkbook();
            else
                workbook = new HSSFWorkbook();


            int sheetno = 0;
            foreach (ToxyTable tsheet in Tables)
            {
                sheetno++;
                ISheet sht;
                if (tsheet.Name == string.Empty)
                {
                    sht = workbook.CreateSheet("Sheet" + sheetno);
                }
                else
                {
                    sht = workbook.CreateSheet(tsheet.Name);
                }

                IDrawing patr = sht.CreateDrawingPatriarch();



                if (tsheet.HasHeader)
                {
                    foreach (ToxyRow trow in tsheet.HeaderRows)
                    {
                        IRow row = sht.CreateRow(trow.RowIndex);

                        foreach (ToxyCell tcell in trow.Cells)
                        {
                            ICell cell = row.CreateCell(tcell.CellIndex);

                            cell.SetCellValue(tcell.Value);

                            if (tcell.Formula != null)
                            {
                                cell.SetCellFormula(tcell.Formula);
                            }
                            if (tcell.Comment != null)
                            {
                                IComment comment1;
                                if (ver == 2007)
                                {
                                    comment1 = patr.CreateCellComment(new XSSFClientAnchor(0, 0, 0, 0, 4, 2, 6, 5));
                                    comment1.String = (new XSSFRichTextString(tcell.Comment));
                                }
                                else
                                {
                                    comment1 = patr.CreateCellComment(new HSSFClientAnchor(0, 0, 0, 0, 4, 2, 6, 5));
                                    comment1.String = (new HSSFRichTextString(tcell.Comment));
                                }

                                cell.CellComment = comment1;
                            }


                        }

                    }



                }

                foreach (ToxyRow trow in tsheet.Rows)
                {
                    IRow row = sht.CreateRow(trow.RowIndex);

                    foreach (ToxyCell tcell in trow.Cells)
                    {
                        ICell cell = row.CreateCell(tcell.CellIndex);

                        cell.SetCellValue(tcell.Value);

                        if (tcell.Formula != null)
                        {
                            cell.SetCellFormula(tcell.Formula);
                        }
                        if (tcell.Comment != null)
                        {
                            IComment comment1;
                            if (ver == 2007)
                            {
                                comment1 = patr.CreateCellComment(new XSSFClientAnchor(0, 0, 0, 0, 4, 2, 6, 5));
                                comment1.String = (new XSSFRichTextString(tcell.Comment));
                            }
                            else
                            {
                                comment1 = patr.CreateCellComment(new HSSFClientAnchor(0, 0, 0, 0, 4, 2, 6, 5));
                                comment1.String = (new HSSFRichTextString(tcell.Comment));
                            }

                            cell.CellComment = comment1;
                        }


                    }

                }


                foreach (MergeCellRange tmergecellrange in tsheet.MergeCells)
                {

                    CellRangeAddress regon = new CellRangeAddress(tmergecellrange.FirstRow, tmergecellrange.LastRow, tmergecellrange.FirstColumn, tmergecellrange.LastColumn);
                    sht.AddMergedRegion(regon);
                }


            }

            FileStream sw = File.Create(filename);
            workbook.Write(sw);
            sw.Close();
            return true;
        }




    }
}

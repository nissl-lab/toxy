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
        public bool ToXLSX(string filename, int type = 2007)
        {


            IWorkbook workbook;
            if (type == 2007)
                workbook = new XSSFWorkbook();
            else
                workbook = new HSSFWorkbook();
            ICellStyle[] cellstyle_array = new ICellStyle[6];
            IFont[] cfont1 = new IFont[6];
            IFontFormatting[] cfont_format = new IFontFormatting[6];
            CellRangeAddress regon;
            ICellStyle titlecellstyle;



            titlecellstyle = workbook.CreateCellStyle();
            titlecellstyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            titlecellstyle.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
            titlecellstyle.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
            titlecellstyle.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
            titlecellstyle.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
            titlecellstyle.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
            titlecellstyle.WrapText = true;


            foreach (ToxyTable tsheet in Tables)
            {
                ISheet sht = workbook.CreateSheet(tsheet.Name);
                IDrawing patr = (HSSFPatriarch)sht.CreateDrawingPatriarch();

                int currentrowindex = 0;

                if (tsheet.HasHeader)
                {

                    for (int j = 0; j < tsheet.HeaderRows.Count; j++)
                    {

                        for (int k = 0; k < tsheet.HeaderRows[j].LastCellIndex; k++)
                        {
                            ICell cell = sht.CreateRow(j + currentrowindex).CreateCell(k);
                            cell.SetCellValue(tsheet.HeaderRows[j].Cells[k].Value);
                            cell.SetCellFormula(tsheet.HeaderRows[j].Cells[k].Formula);
                            IComment comment1 = patr.CreateCellComment(new HSSFClientAnchor(0, 0, 0, 0, 4, 2, 6, 5));
                            comment1.String = (new HSSFRichTextString(tsheet.HeaderRows[j].Cells[k].Comment));
                            cell.CellComment = comment1;

                        }
                        currentrowindex++;
                    }

                }

                for (int j = 0; j < tsheet.Rows.Count; j++)
                {

                    for (int k = 0; k < tsheet.Rows[j].LastCellIndex; k++)
                    {
                        ICell cell = sht.CreateRow(j + currentrowindex).CreateCell(k);
                        cell.SetCellValue(tsheet.Rows[j].Cells[k].Value);
                        cell.SetCellFormula(tsheet.Rows[j].Cells[k].Formula);
                        IComment comment1 = patr.CreateCellComment(new HSSFClientAnchor(0, 0, 0, 0, 4, 2, 6, 5));
                        comment1.String = (new HSSFRichTextString(tsheet.HeaderRows[j].Cells[k].Comment));
                        cell.CellComment = comment1;

                    }
                    currentrowindex++;
                }


                for (int m = 0; m < tsheet.MergeCells.Count; m++)
                {

                    regon = new CellRangeAddress(tsheet.MergeCells[m].FirstRow, tsheet.MergeCells[m].LastRow, tsheet.MergeCells[m].FirstColumn, tsheet.MergeCells[m].LastColumn);
                    sht.AddMergedRegion(regon);
                }


            }

            FileStream sw = File.Create(filename);
            workbook.Write(sw);
            sw.Close();
            return true;
        }



        public bool ToHtml(string filename)
        {

            return false;
        }

    }
}

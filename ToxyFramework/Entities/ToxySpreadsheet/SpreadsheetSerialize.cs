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

namespace Toxy
{
    public partial class  ToxySpreadsheet:ICloneable
    {

        private static int compresslevel = 9;

        public MemoryStream Serialize(bool bcompress = true)
        {
            MemoryStream ms = new MemoryStream();
            RawSerializer rs = new RawSerializer(ms);

            //sheet Serialize
            rs.SerializeNullable(Name);

            //table Serialize
            rs.Serialize(Tables.Count);
            foreach (ToxyTable tb in Tables)
            {

                rs.Serialize(tb.HasHeader);
                rs.Serialize(tb.Name);
                rs.SerializeNullable(tb.PageHeader);
                rs.SerializeNullable(tb.PageFooter);
                rs.Serialize(tb.LastRowIndex);
                rs.Serialize(tb.LastColumnIndex);
                //mergecell 
                rs.Serialize(tb.MergeCells.Count);
                foreach (MergeCellRange mcr in tb.MergeCells)
                {
                    rs.Serialize(mcr.FirstRow);
                    rs.Serialize(mcr.LastRow);
                    rs.Serialize(mcr.FirstColumn);
                    rs.Serialize(mcr.LastColumn);
                }

                //columnheaders

                if (tb.HasHeader)
                {
                    rs.Serialize(tb.HeaderRows.Count);
                    foreach (ToxyRow row in tb.HeaderRows)
                    {
                        rs.Serialize(row.RowIndex);
                        rs.Serialize(row.LastCellIndex);
                        //cells
                        rs.Serialize(row.Cells.Count);
                        foreach (ToxyCell cell in row.Cells)
                        {
                            rs.Serialize(cell.Value);
                            rs.Serialize(cell.CellIndex);
                            rs.SerializeNullable(cell.Comment);
                            rs.SerializeNullable(cell.Formula);

                        }


                    }
                }

                //rows
                rs.Serialize(tb.Rows.Count);
                foreach (ToxyRow row in tb.Rows)
                {
                    rs.Serialize(row.RowIndex);
                    rs.Serialize(row.LastCellIndex);
                    //cells
                    rs.Serialize(row.Cells.Count);
                    foreach (ToxyCell cell in row.Cells)
                    {
                        rs.Serialize(cell.Value);
                        rs.Serialize(cell.CellIndex);
                        rs.SerializeNullable(cell.Comment);
                        rs.SerializeNullable(cell.Formula);

                    }


                }


            }

            rs.Flush();
            //http://m.blog.csdn.net/blog/jecray/1554669
            //http://blog.sina.com.cn/s/blog_40678c33010007zk.html


            MemoryStream ms2 = new MemoryStream();
            ms.Position = 0;

            if (bcompress)
            {

                BZip2.Compress(ms, ms2, false, compresslevel);
                //ms2.Position = 0;
                //BZip2.Decompress(ms2,ms3, false);
                return ms2;
            }
            else
            {
                return ms;
            }




        }




        public bool Deserialize(MemoryStream ms, bool bcompress = true)
        {


            ms.Position = 0;
            MemoryStream msd = new MemoryStream();

            RawDeserializer rd;
            if (bcompress)
            {

                BZip2.Decompress(ms, msd, false);
                msd.Position = 0;

                rd = new RawDeserializer(msd);
            }
            else
            {
                rd = new RawDeserializer(ms);
            }



            //sheet Serialize
            var sheetName = rd.DeserializeNullable(typeof(string));
            if (sheetName != null)
            {
                Name = (string)sheetName;
            }

            //table Serialize
            int TablesCount = rd.DeserializeInt();
            for (int i = 0; i < TablesCount; i++)
            {
                ToxyTable tb = new ToxyTable();
                bool tbHasHeader = rd.DeserializeBool();
                tb.Name = rd.DeserializeString();
                var tbPageHeader = rd.DeserializeNullable(typeof(string));
                if (tbPageHeader != null)
                {
                    tb.PageHeader = (string)tbPageHeader;
                }
                var tbPageFooter = rd.DeserializeNullable(typeof(string));
                if (tbPageFooter != null)
                {
                    tb.PageFooter = (string)tbPageFooter;
                }

                tb.LastRowIndex = rd.DeserializeInt();
                tb.LastColumnIndex = rd.DeserializeInt();
                //mergecell 
                int tbMergeCellsCount = rd.DeserializeInt();
                for (int j = 0; j < tbMergeCellsCount; j++)
                {
                    MergeCellRange mcr = new MergeCellRange();
                    mcr.FirstRow = rd.DeserializeInt();
                    mcr.LastRow = rd.DeserializeInt();
                    mcr.FirstColumn = rd.DeserializeInt();
                    mcr.LastColumn = rd.DeserializeInt();
                    tb.MergeCells.Add(mcr);
                }

                //columnheaders

                if (tbHasHeader)
                {
                    int tbHeaderCount = rd.DeserializeInt();
                    for (int m = 0; m < tbHeaderCount; m++)
                    {
                        int rowRowIndex = rd.DeserializeInt();
                        ToxyRow row = new ToxyRow(rowRowIndex);

                        row.LastCellIndex = rd.DeserializeInt();
                        //cells
                        int rowCellsCount = rd.DeserializeInt();
                        for (int n = 0; n < rowCellsCount; n++)
                        {

                            string cellValue = rd.DeserializeString();
                            int cellCellIndex = rd.DeserializeInt();
                            ToxyCell cell = new ToxyCell(cellCellIndex, cellValue);
                            var cellComment = rd.DeserializeNullable(typeof(string));
                            if (cellComment != null)
                            {
                                cell.Comment = (string)cellComment;
                            }
                            var cellformula = rd.DeserializeNullable(typeof(string));
                            if (cellformula != null)
                            {
                                cell.Formula = (string)cellformula;
                            }

                            row.Cells.Add(cell);
                        }

                        tb.HeaderRows.Add(row);
                    }
                }

                //rows
                int tbRowsCount = rd.DeserializeInt();
                for (int m = 0; m < tbRowsCount; m++)
                {
                    int rowRowIndex = rd.DeserializeInt();
                    ToxyRow row = new ToxyRow(rowRowIndex);

                    row.LastCellIndex = rd.DeserializeInt();
                    //cells
                    int rowCellsCount = rd.DeserializeInt();
                    for (int n = 0; n < rowCellsCount; n++)
                    {

                        string cellValue = rd.DeserializeString();
                        int cellCellIndex = rd.DeserializeInt();
                        ToxyCell cell = new ToxyCell(cellCellIndex, cellValue);
                        var cellComment = rd.DeserializeNullable(typeof(string));
                        if (cellComment != null)
                        {
                            cell.Comment = (string)cellComment;
                        }
                        var cellformula = rd.DeserializeNullable(typeof(string));
                        if (cellformula != null)
                        {
                            cell.Formula = (string)cellformula;
                        }

                        row.Cells.Add(cell);
                    }

                    tb.Rows.Add(row);
                }


                Tables.Add(tb);

            }






            return true;

        }




    }
}

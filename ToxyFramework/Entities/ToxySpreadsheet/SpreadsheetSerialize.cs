using System;
using System.Collections.Generic;
using System.Data;
using System.Text;
using Clifton.Tools.Data;
using System.IO;


namespace Toxy
{
    public partial class ToxySpreadsheet
    {

        public MemoryStream Serialize()
        {
            MemoryStream ms=new MemoryStream();
            RawSerializer rs = new RawSerializer(ms);

            //sheet Serialize
            rs.Serialize(Name);

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
                    rs.Serialize(tb.ColumnHeaders.RowIndex);
                    rs.Serialize(tb.ColumnHeaders.LastCellIndex);
                    //cells
                    rs.Serialize(tb.ColumnHeaders.Cells.Count);
                    foreach(ToxyCell cell in tb.ColumnHeaders.Cells)
                    {
                        rs.Serialize(cell.Value);
                        rs.Serialize(cell.CellIndex);
                        rs.SerializeNullable(cell.Comment);
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
                    foreach(ToxyCell cell in row.Cells)
                    {
                        rs.Serialize(cell.Value);
                        rs.Serialize(cell.CellIndex);
                        rs.SerializeNullable(cell.Comment);
                    
                    }


                }


            }

            rs.Flush();

            return ms;

        }


        public bool Deserialize(MemoryStream ms)
        {
            RawDeserializer rd=new RawDeserializer(ms);

            ms.Position = 0;

            

            //sheet Serialize
            Name=rd.DeserializeString();

            //table Serialize
            int TablesCount=rd.DeserializeInt();
            for (int i =0;i<TablesCount;i++)
            {
                ToxyTable tb=new ToxyTable();
               bool tbHasHeader=rd.DeserializeBool();
                tb.Name= rd   .DeserializeString()   ;
                tb.PageHeader = (string) rd.DeserializeNullable(typeof(string));
                tb.PageFooter = (string)rd.DeserializeNullable(typeof(string));
                tb.LastRowIndex= rd  .DeserializeInt()    ;
                tb.LastColumnIndex = rd.DeserializeInt();
                //mergecell 
               int  tbMergeCellsCount = rd.DeserializeInt();
               for (int j = 0; j < tbMergeCellsCount;j++ )
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
                    tb.ColumnHeaders.RowIndex = rd.DeserializeInt();
                    tb.ColumnHeaders.LastCellIndex = rd.DeserializeInt();
                    //cells
                    int tbColumnHeadersCellsCount = rd.DeserializeInt();
                    for (int k = 0; k < tbColumnHeadersCellsCount;k++ )
                    {
                        
                        string cellValue = rd.DeserializeString();
                        int cellCellIndex = rd.DeserializeInt();
                        ToxyCell cell = new ToxyCell(cellCellIndex,cellValue);
                        cell.Comment = rd.DeserializeString();
                        tb.ColumnHeaders.Cells.Add(cell);
                    }
                }

                //rows
                int tbRowsCount = rd.DeserializeInt();
                for (int m = 0; m < tbRowsCount;m++ )
                {
                    int rowRowIndex = rd.DeserializeInt();
                    ToxyRow row = new ToxyRow(rowRowIndex);

                    row.LastCellIndex = rd.DeserializeInt();
                    //cells
                    int rowCellsCount = rd.DeserializeInt();
                    for (int n = 0; n < rowCellsCount;n++ )
                    {

                        string cellValue = rd.DeserializeString();
                        int cellCellIndex = rd.DeserializeInt();
                        ToxyCell cell = new ToxyCell(cellCellIndex, cellValue);
                        cell.Comment = rd.DeserializeString();
                        row.Cells.Add(cell);
                    }


                }


                Tables.Add(tb);

            }






            return true;

        }




    }
}

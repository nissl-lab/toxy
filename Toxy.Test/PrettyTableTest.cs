using System;
using System.Collections.Generic;
using NUnit.Framework;
using NUnit.Framework.Legacy;

namespace Toxy.Test;

[TestFixture]
public class PrettyTableTest
{
    private ToxyTable BuildTable()
    {
        ToxyTable table = new ToxyTable();

        var singleRow = new ToxyRow(0)
        {
            Cells =
            [
                new ToxyCell(0, "City name"),
                new ToxyCell(1, "Area"),
                new ToxyCell(2, "Population"),
                new ToxyCell(3, "Annual Rainfall")
            ]
        };
        table.HeaderRows =
        [
            singleRow
        ];
        table.Rows = new List<ToxyRow>();
        var row1 = new ToxyRow(0)
        {
            Cells =
            [
                new ToxyCell(0, "Adelaide"),
                new ToxyCell(1, "1295"),
                new ToxyCell(2, "1158259"),
                new ToxyCell(3, "600.5")
            ]
        };
        table.Rows.Add(row1);
        var row2 = new ToxyRow(1)
        {
            Cells =
            [
                new ToxyCell(0, "Brisbane"),
                new ToxyCell(1, "5905"),
                new ToxyCell(2, "1857594"),
                new ToxyCell(3, "1146.4")
            ]
        };
        table.Rows.Add(row2);
        var row3 = new ToxyRow(2)
        {
            Cells =
            [
                new ToxyCell(0, "Darwin"),
                new ToxyCell(1, "112"),
                new ToxyCell(2, "120900"),
                new ToxyCell(3, "1714.7")
            ]
        };
        table.Rows.Add(row3);
        return table;
    }

    [Test]
    public void TestPrint()
    {
        ToxyTable table = BuildTable();
        ClassicAssert.AreEqual(8,table.Print().Split(Environment.NewLine).Length);
    }

    [Test]
    public void TestGetStringWithStartAndEnd()
    {
        ToxyTable table = BuildTable();
        var results = table.Print(1,2).Split(Environment.NewLine);
        ClassicAssert.AreEqual("+------+------------+", results[0]);
        ClassicAssert.AreEqual("| Area | Population |", results[1]);
        ClassicAssert.AreEqual("+------+------------+", results[2]);
        ClassicAssert.AreEqual("| 5905 | 1857594 |", results[4]);
        ClassicAssert.AreEqual("| 112 | 120900 |", results[5]);
        ClassicAssert.AreEqual("+------+------------+", results[6]);
    }
}
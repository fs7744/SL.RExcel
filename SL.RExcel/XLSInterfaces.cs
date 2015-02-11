using System.Collections.Generic;

namespace SL.RExcel
{
    public interface IWorkBook
    {
        IWorksheet[] Worksheets { get; }

        int Count { get; }

        IWorksheet GetSheetByName(string name);

        IWorksheet GetSheetByIndex(int index);

        IEnumerable<string> GetAllSheetNames();
    }

    public interface IWorksheet
    {
        string Name { get; }

        IDictionary<uint, IRow> Rows { get; }

        uint FirstRow { get; }

        uint LastRow { get; }

        //uint FirstCol { get; }

        //uint LastCol { get; }

        IRow GetRow(uint index);
    }

    public interface IRow
    {
        IDictionary<uint, ICell> Cells { get; }

        bool IsEmpty();

        ICell GetCell(uint index);
    }

    public interface ICell
    {
        object Value { get; }

        string GetStringValue();
    }
}
using SL.RExcel.XLS.Records;
using System;
using System.Collections.Generic;

namespace SL.RExcel.XLS
{
    public class XLSSheet : IWorksheet
    {
        public string Name { get; private set; }

        public IRow[] Rows
        {
            get { throw new NotImplementedException(); }
        }

        public uint FirstRow
        {
            get { throw new NotImplementedException(); }
        }

        public uint LastRow
        {
            get { throw new NotImplementedException(); }
        }

        public uint FirstCol
        {
            get { throw new NotImplementedException(); }
        }

        public uint LastCol
        {
            get { throw new NotImplementedException(); }
        }

        public XLSSheet(BoundSheetRecord record, List<Record> allRecords)
        {
            Name = record.Name;
        }

        public IRow GetRow(uint index)
        {
            throw new NotImplementedException();
        }
    }
}
using SL.RExcel.XLS.File;
using SL.RExcel.XLS.Records;
using System;
using System.Collections.Generic;
using System.IO;

namespace SL.RExcel.XLS
{
    public class XLSWorkBook : IWorkBook
    {
        private Dictionary<ushort, Type> m_Records;

        public IWorksheet[] Worksheets { get; private set; }

        public int Count
        {
            get { return Worksheets.Length; }
        }

        public XLSWorkBook(Stream stream)
        {
            m_Records = new Dictionary<ushort, Type>();
            //m_Records.Add((ushort)RecordType.Bof, typeof(BofRecord));
            m_Records.Add((ushort)RecordType.Boundsheet, typeof(BoundSheetRecord));
            m_Records.Add((ushort)RecordType.Index, typeof(IndexRecord));
            m_Records.Add((ushort)RecordType.DbCell, typeof(DbCellRecord));
            m_Records.Add((ushort)RecordType.Row, typeof(RowRecord));
            //m_Records.Add((ushort)RecordType.Continue, typeof(ContinueRecord));
            m_Records.Add((ushort)RecordType.Blank, typeof(BlankRecord));
            m_Records.Add((ushort)RecordType.BoolErr, typeof(BoolErrRecord));
            //m_Records.Add((ushort)RecordType.Formula, typeof(FormulaRecord));
            m_Records.Add((ushort)RecordType.Label, typeof(LabelRecord));
            m_Records.Add((ushort)RecordType.LabelSst, typeof(LabelSstRecord));
            //m_Records.Add((ushort)RecordType.MulBlank, typeof(MulBlankRecord));
            //m_Records.Add((ushort)RecordType.MulRk, typeof(MulRkRecord));
            m_Records.Add((ushort)RecordType.String, typeof(StringValueRecord));
            //m_Records.Add((ushort)RecordType.Xf, typeof(XfRecord));
            //m_Records.Add((ushort)RecordType.Rk, typeof(RkRecord));
            m_Records.Add((ushort)RecordType.Number, typeof(NumberRecord));
            m_Records.Add((ushort)RecordType.Array, typeof(ArrayRecord));
            //m_Records.Add((ushort)RecordType.ShrFmla, typeof(SharedFormulaRecord));
            //m_Records.Add((ushort)RecordType.Table, typeof(TableRecord));
            m_Records.Add((ushort)RecordType.Sst, typeof(SstRecord));
            m_Records.Add((ushort)RecordType.Eof, typeof(EofRecord));
            //m_Records.Add((ushort)RecordType.Font, typeof(FontRecord));
            //m_Records.Add((ushort)RecordType.Format, typeof(FormatRecord));
            //m_Records.Add((ushort)RecordType.Palette, typeof(PaletteRecord));
            //m_Records.Add((ushort)RecordType.Hyperlink, typeof(HyperLinkRecord));
            Load(new XLSFile(stream));
            m_Records = null;
        }

        private void Load(XLSFile file)
        {
            using (var stream = GetWorkBookStream(file))
            {
                List<Record> records = new List<Record>();
                SstRecord sst = null;
                while (stream.Length - stream.Position >= BIFFData.MinSize)
                {
                    var record = GetRecord(new BIFFData(stream), stream);
                    if (record != null)
                    {
                        sst = SetSST(sst, record);
                        records.Add(record);
                    }
                }
            }
        }

        private SstRecord SetSST(SstRecord sst, Record record)
        {
            if (record is SstRecord)
                sst = record as SstRecord;
            else if (record is LabelSstRecord && sst != null)
                ((LabelSstRecord)record).SetValue(sst);
            return sst;
        }

        private Record GetRecord(BIFFData data, Stream stream)
        {
            Type type = null;
            m_Records.TryGetValue(data.ID, out type);
            return type == null ? null : (Record)Activator.CreateInstance(type, data, stream);
        }

        private Stream GetWorkBookStream(XLSFile file)
        {
            try
            {
                return file.OpenStream("Workbook");
            }
            catch (IOException)
            {
                return file.OpenStream("Book");
            }
        }

        public IWorksheet GetSheetByName(string name)
        {
            throw new NotImplementedException();
        }

        public IWorksheet GetSheetByIndex(int index)
        {
            throw new NotImplementedException();
        }

        public System.Collections.Generic.IEnumerable<string> GetAllSheetNames()
        {
            throw new NotImplementedException();
        }
    }
}
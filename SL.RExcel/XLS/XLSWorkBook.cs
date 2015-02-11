using SL.RExcel.XLS.File;
using SL.RExcel.XLS.Records;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

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
            stream.Position = 0L;
            m_Records = new Dictionary<ushort, Type>();
            //m_Records.Add((ushort)RecordType.Bof, typeof(BofRecord));
            m_Records.Add((ushort)RecordType.Boundsheet, typeof(BoundSheetRecord));
            m_Records.Add((ushort)RecordType.Index, typeof(IndexRecord));
            m_Records.Add((ushort)RecordType.DbCell, typeof(DbCellRecord));
            m_Records.Add((ushort)RecordType.Row, typeof(RowRecord));
            //m_Records.Add((ushort)RecordType.Continue, typeof(ContinueRecord));
            //m_Records.Add((ushort)RecordType.Blank, typeof(BlankRecord));
            m_Records.Add((ushort)RecordType.BoolErr, typeof(BoolErrRecord));
            //m_Records.Add((ushort)RecordType.Formula, typeof(FormulaRecord));
            m_Records.Add((ushort)RecordType.Label, typeof(LabelRecord));
            m_Records.Add((ushort)RecordType.LabelSst, typeof(LabelSstRecord));
            //m_Records.Add((ushort)RecordType.MulBlank, typeof(MulBlankRecord));
            m_Records.Add((ushort)RecordType.MulRk, typeof(MulRkRecord));
            m_Records.Add((ushort)RecordType.String, typeof(StringValueRecord));
            //m_Records.Add((ushort)RecordType.Xf, typeof(XfRecord));
            m_Records.Add((ushort)RecordType.Rk, typeof(RkRecord));
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
            m_Records.Add((ushort)RecordType.MergeCells, typeof(MergeCellsRecord));
            Load(new XLSFile(stream));
            m_Records = null;
        }

        #region load

        private void Load(XLSFile file)
        {
            List<SheetRecord> sheets = new List<SheetRecord>();
            using (var stream = GetWorkBookStream(file))
            {
                SstRecord sst = null;
                SheetRecord currentSheet = null;
                while (stream.Length - stream.Position >= BIFFData.MinSize)
                {
                    var record = GetRecord(new BIFFData(stream), stream);
                    if (record != null)
                        GetSheetRecords(sheets, ref sst, ref currentSheet, record);
                }
                sheets.ForEach(i => i.SST = sst);
            }

            if (sheets.Count > 0)
                GetSheets(sheets);
        }

        private void GetSheets(List<SheetRecord> sheets)
        {
            Worksheets = sheets.Select(i => new XLSSheet(i)).ToArray();
        }

        private static void GetSheetRecords(List<SheetRecord> sheets, ref SstRecord sst, ref SheetRecord currentSheet, Record record)
        {
            if (record is SstRecord)
                sst = record as SstRecord;
            else if (record is BoundSheetRecord)
                sheets.Add(new SheetRecord() { Sheet = record as BoundSheetRecord });
            else if (record is IndexRecord)
            {
                currentSheet = sheets.FirstOrDefault(i => i.Index == null);
                if (currentSheet != null)
                    currentSheet.Index = record as IndexRecord;
            }
            else if (record is RowRecord && currentSheet != null)
            {
                currentSheet.Rows.Add(record as RowRecord);
            }
            else if (record is CellRecord && currentSheet != null)
            {
                currentSheet.Cells.Add(record as CellRecord);
            }
            else if (record is MergeCellsRecord && currentSheet != null)
            {
                currentSheet.MergeCells.Add(record as MergeCellsRecord);
            }
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

        #endregion load

        public IWorksheet GetSheetByName(string name)
        {
            return Worksheets.FirstOrDefault(i => i.Name == name);
        }

        public IWorksheet GetSheetByIndex(int index)
        {
            return Worksheets[index];
        }

        public IEnumerable<string> GetAllSheetNames()
        {
            return Worksheets.Select(i => i.Name);
        }
    }
}
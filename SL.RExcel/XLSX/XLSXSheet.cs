using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace SL.RExcel.XLSX
{
    public class MergeCell
    {
        public uint FirstRow { get; set; }

        public uint LastRow { get; set; }

        public uint FirstCol { get; set; }

        public uint LastCol { get; set; }
    }

    public class XLSXSheet : IWorksheet
    {
        public string Name { get; private set; }

        public IDictionary<uint, IRow> Rows { get; private set; }

        public uint FirstRow { get; private set; }

        public uint LastRow { get; private set; }

        public XLSXSheet(string name, string id, UnZipper unzip, Dictionary<string, string> sharedStrings, List<string> fileNames)
        {
            Name = name;
            Rows = new Dictionary<uint, IRow>();
            var xml_sheetFileName = string.Format(XLSXCommon.Sheet, id);
            if (!fileNames.Any(i => i.EqualsIgnoreCase(xml_sheetFileName))) return;
            {
                var element = unzip.GetXLSXPart(xml_sheetFileName);
                foreach (var row in element.Descendants(XLSXCommon.ExcelNamespace + XLSXCommon.XML_Row))
                {
                    var indexStr = row.Attribute(XLSXCommon.XML_R);
                    var index = int.Parse(indexStr.Value);
                    Rows.Add((uint)index, new XLSXRow(row, sharedStrings));
                }
                FillMergeCells(GetMergeCells(element));
            }
        }

        private void FillMergeCells(List<MergeCell> list)
        {
            foreach (var mergeCell in list)
            {
                var value = GetFirstMergeCellValue(mergeCell);
                if (value != null)
                    SetMergeCellValue(mergeCell, value);
            }
        }

        private void SetMergeCellValue(MergeCell mergeCell, object value)
        {
            for (uint i = mergeCell.FirstRow; i <= mergeCell.LastRow; i++)
            {
                IRow row = null;
                if (!Rows.TryGetValue(i, out row))
                    Rows.Add(i, new XLSXRow());
                for (uint j = mergeCell.FirstCol; j <= mergeCell.LastCol; j++)
                {
                    ICell cell = null;
                    if (!row.Cells.TryGetValue(j, out cell))
                    {
                        cell = new XLSXCell(value);
                        row.Cells.Add(j, cell);
                    }
                }
            }
        }

        private object GetFirstMergeCellValue(MergeCell mergeCell)
        {
            for (uint i = mergeCell.FirstRow; i <= mergeCell.LastRow; i++)
            {
                IRow row = null;
                if (!Rows.TryGetValue(i, out row))
                    continue;
                for (uint j = mergeCell.FirstCol; j <= mergeCell.LastCol; j++)
                {
                    ICell cell = null;
                    if (!row.Cells.TryGetValue(j, out cell) || cell.Value == null)
                        continue;
                    else
                        return cell.Value;
                }
            }
            return null;
        }

        private static List<MergeCell> GetMergeCells(XElement element)
        {
            List<MergeCell> areas = new List<MergeCell>();
            foreach (var item in element.Descendants(XLSXCommon.ExcelNamespace + XLSXCommon.XML_MergeCell))
            {
                var refStr = item.Attribute(XLSXCommon.XML_Ref).Value;
                var indexs = refStr.ToUpper().Split(XLSXCommon.RefSplit, StringSplitOptions.RemoveEmptyEntries);
                var first = indexs[0].ConvertIndex();
                var last = indexs[1].ConvertIndex();
                areas.Add(new MergeCell()
                {
                    FirstCol = (uint)first[0],
                    LastCol = (uint)last[0],
                    FirstRow = (uint)first[1],
                    LastRow = (uint)last[1]
                });
            }
            return areas;
        }

        public IRow GetRow(uint index)
        {
            IRow result = null;
            Rows.TryGetValue(index, out result);
            return result;
        }
    }
}
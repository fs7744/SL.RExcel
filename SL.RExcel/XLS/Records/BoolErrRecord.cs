using System.IO;

namespace SL.RExcel.XLS.Records
{
    public class BoolErrRecord : RowColXfCellRecord
    {
        //public object Value { get; private set; }

        public bool Error { get; private set; }

        public BoolErrRecord(BIFFData data, Stream stream)
        {
            BinaryReader reader = new BinaryReader(data.ToStream());
            SetRowColXf(reader);
            byte boolErr = reader.ReadByte();
            Error = reader.ReadByte() == 1;
            if (Error)
                SetBoolErrValue(boolErr);
            else
                Value = boolErr == 1;
        }

        private void SetBoolErrValue(byte boolErr)
        {
            switch (boolErr)
            {
                case 0x00:
                    Value = "#NULL!";
                    break;

                case 0x07:
                    Value = "#DIV/0";
                    break;

                case 0x0f:
                    Value = "#VALUE!";
                    break;

                case 0x17:
                    Value = "#REF!";
                    break;

                case 0x1d:
                    Value = "#NAME?";
                    break;

                case 0x24:
                    Value = "#NUM!";
                    break;

                case 0x2a:
                    Value = "#N/A";
                    break;

                default:
                    Value = "";
                    break;
            }
        }
    }
}
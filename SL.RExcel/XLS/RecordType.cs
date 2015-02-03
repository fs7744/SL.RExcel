namespace SL.RExcel.XLS
{
    public enum RecordType : ushort
    {
        Bof = 0x0809,
        Boundsheet = 0x0085,

        Index = 0x020B,

        DbCell = 0x00d7,

        Row = 0x0208,

        Continue = 0x003c,

        Sst = 0x00fc,

        Blank = 0x0201,

        BoolErr = 0x0205,

        Formula = 0x006,

        Label = 0x0204,

        LabelSst = 0x00fd,

        MulBlank = 0x00be,

        MulRk = 0x00bd,

        String = 0x0207,

        Xf = 0x00e0,

        Eof = 0x000a,

        Rk = 0x027e,

        Number = 0x0203,

        Array = 0x0221,

        ShrFmla = 0x00bc,

        Table = 0x0036,

        Font = 0x0031,

        Format = 0x041e,

        Palette = 0x0092,

        Hyperlink = 0x01B8,
    }
}
namespace SL.RExcel.XLS.File
{
    public enum Stgty : byte
    {
        Invalid = 0,
        Storage,
        Stream,
        LockBytes,
        Property,
        Root
    }

    public enum DeColor : byte
    {
        Red = 0,
        Black
    }
}
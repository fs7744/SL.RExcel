namespace SL.RExcel.XLS.File
{
    public class SectorIndex
    {
        public const uint DifSectorIndex = 0xFFFFFFFC;
        public const uint FatSectorIndex = 0xFFFFFFFD;
        public const uint EndOfChain = 0xFFFFFFFE;
        public const uint FreeSectorIndex = 0xFFFFFFFF;
        public static readonly SectorIndex ZERO = new SectorIndex(0);

        private uint Value { get; private set; }

        public bool IsFree { get { return Value == FreeSectorIndex; } }

        public bool IsEndOfChain { get { return Value == EndOfChain; } }
        
        public SectorIndex(uint value)
        {
            Value = value;
        }

        public override bool Equals(object obj)
        {
            return Value.Equals(obj != null && obj is SectorIndex ? ((SectorIndex)obj).Value : obj);
        }

        public override int GetHashCode()
        {
            return Value.GetHashCode();
        }

        public int ToInt()
        {
            return (int)Value;
        }
    }
}
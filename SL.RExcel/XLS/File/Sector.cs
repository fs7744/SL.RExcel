namespace SL.RExcel.XLS.File
{
    public class Sector
    {
    }

    public static class SectorHelper
    {
        public static Storage ToStorage(this Sector sector)
        {
            return sector as Storage;
        }
    }
}
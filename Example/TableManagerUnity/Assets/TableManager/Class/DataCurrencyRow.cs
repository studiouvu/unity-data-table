using Enums;

namespace TableManager
{
    public partial class DataCurrencyRow : IRow
    {
        /*  */
        public string id { get; set; } 

        /*  */
        public string nameId { get; set; } 

        /*  */
        public string descId { get; set; } 

        /*  */
        public string iconAddress { get; set; } 

        /*  */
        public bool displayCondense { get; set; } 

        /* 0 */
        public long[] testArray { get; set; } 


    }
}

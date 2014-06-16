namespace FXSpotTrading.Models
{
    using System;

    public class Deal
    {
        public string Ccy1 { get; set; }
        public string Ccy2 { get; set; }
        public double Amount1 { get; set; }
        public double Amount2 { get; set; }
        public double Rate { get; set; }
        public bool IsCcy1Fixed { get; set; }

        public bool TryDeriveXcrossAmount()
        {
            if (IsCcy1Fixed)
            {
                Amount2 = Amount1*Rate;
                return true;
            }
            else if (Math.Abs(Rate) > 0.000001)
            {
                Amount1 = Amount2/Rate;
                return true;
            }
            return false;
        }
    }
}

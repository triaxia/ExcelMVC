namespace FXSpotTrading.Models
{
    using System;

    public class Deal
    {
        public string BuyCcy { get; set; }
        public string SellCcy { get; set; }
        public double BuyAmount { get; set; }
        public double SellAmount { get; set; }
        public bool IsCcy1Fixed { get; set; }

        public ExchangeRate Rate { get; set; }

        public bool TryDeriveXAmount()
        {
            if (Rate == null)
                return false;

            double fx;
            if (Rate.Pair.Ccy1 == BuyCcy)
                fx = Rate.Ask;
            else
                fx = 1 / Rate.Bid;

            if (IsCcy1Fixed)
            {
                SellAmount = BuyAmount * fx;
                return true;
            }
            else if (Math.Abs(fx) > 0.000001)
            {
                BuyAmount = SellAmount / fx;
                return true;
            }
            return false;
        }
    }
}

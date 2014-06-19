namespace SpotTrading.BusinessModels
{
    using System;

    public class ExchangeRate
    {
        public CcyPair Pair { get; set; }
        public double Bid { get; set; }
        public double Ask { get; set; }

        public ExchangeRate Flip()
        {
            var rate = new ExchangeRate { Pair = new CcyPair { Ccy1 = Pair.Ccy2, Ccy2 = Pair.Ccy1, Pip = Pair.Pip } };
            rate.Pair.Pip = Pair.Pip;
            rate.Bid = 1.0 / Ask;
            rate.Ask = 1.0 / Bid;
            return rate;
        }

        public static ExchangeRate Cross(ExchangeRate lhs, ExchangeRate rhs)
        {
            Func<CcyPair, string> nonBaseCcy = x => x.Ccy1 == "USD" ? x.Ccy2 : x.Ccy1;
            var rate = new ExchangeRate { Pair = new CcyPair { Ccy1 = nonBaseCcy(lhs.Pair), Ccy2 = nonBaseCcy(rhs.Pair) } };
            rate.Pair.Pip = Math.Max(lhs.Pair.Pip, rate.Pair.Pip);

            if (rate.Pair.Ccy1 == lhs.Pair.Ccy2)
                lhs = lhs.Flip();

            if (rate.Pair.Ccy1 == rhs.Pair.Ccy2)
                rhs = rhs.Flip();

            rate.Bid = lhs.Bid * rhs.Ask;
            rate.Ask = lhs.Ask * rhs.Bid;

            return rate;
        }
    }
}

using System.Collections.Generic;
using System.Linq;

namespace FXSpotTrading.Models
{
    public class ExchangeRates : List<ExchangeRate>
    {
        public ExchangeRates(IEnumerable<CcyPair> pairs)
        {
            Create(pairs);
        }

        public void Create(IEnumerable<CcyPair> pairs)
        {
            Clear();
            AddRange(pairs.Where(x => x.IsValid).Select(y => new ExchangeRate { Pair = y, Bid = y.Spot - y.Pip, Ask = y.Spot + y.Pip }));
        }

        public ExchangeRate Find(string ccy1, string ccy2)
        {
            if (ccy1 == null || ccy2 == null)
                return null;

            var rate = this.FirstOrDefault(x => (x.Pair.Ccy1 == ccy1 && x.Pair.Ccy2 == ccy2)
                || (x.Pair.Ccy1 == ccy2 && x.Pair.Ccy2 == ccy1));

            if (rate != null)
                return rate;

            var lhs = Find(ccy1, "USD");
            if (lhs == null)
                return null;

            var rhs = Find(ccy2, "USD");
            if (rhs == null)
                return null;

            return ExchangeRate.Cross(lhs, rhs);
        }
    }
}

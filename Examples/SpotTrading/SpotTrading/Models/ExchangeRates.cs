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
            AddRange(pairs.Where(x => x.IsValid).Select(y => new ExchangeRate { Pair = y, Bid = y.Spot, Ask = y.Spot }));
        }

        public ExchangeRate Find(string ccy1, string ccy2)
        {
            if (ccy1 == null || ccy2 == null)
                return null;

            return this.FirstOrDefault(x => (x.Pair.Ccy1 == ccy1 && x.Pair.Ccy2 == ccy2)
                || (x.Pair.Ccy1 == ccy2 && x.Pair.Ccy2 == ccy1));
        }
    }
}

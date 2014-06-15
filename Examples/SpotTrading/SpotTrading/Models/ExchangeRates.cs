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
    }
}

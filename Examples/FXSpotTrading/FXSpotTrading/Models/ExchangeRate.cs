namespace FXSpotTrading.Models
{
    public class ExchangeRate
    {
        public CcyPair Pair { get; set; }
        public double Bid { get; set; }
        public double Ask { get; set; }
    }
}

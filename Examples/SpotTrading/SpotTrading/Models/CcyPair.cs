namespace FXSpotTrading.Models
{
    public class CcyPair
    {
        public string Ccy1 { get; set; }
        public string Ccy2 { get; set; }
        public double Pip { get; set; }
        public double Spot { get; set; }
      
        public string Code
        {
            get { return string.Format("{0}/{1}", Ccy1, Ccy2); }
        }

        public bool IsValid
        {
            get { return !string.IsNullOrEmpty(Ccy1) && !string.IsNullOrEmpty(Ccy2); }
        }
    }
}

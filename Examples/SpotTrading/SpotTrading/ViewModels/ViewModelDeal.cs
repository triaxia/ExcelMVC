namespace FXSpotTrading.ViewModels
{
    using System.ComponentModel;
    using Models;

    public class ViewModelDeal : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged = delegate { };

        public Deal Model { get; private set; }
        public ExchangeRates Rates { get; private set; }

        public ViewModelDeal(ExchangeRates rates)
        {
            Model = new Deal();
            Rates = rates;
        }

        public string Ccy1
        {
            get 
            {
                return Model.Ccy1; 
            }
            set
            {
                Model.Ccy1 = value;
                SetRate();
            }
        }

        public string Ccy2
        {
            get
            {
                return Model.Ccy2;
            }
            set
            {
                Model.Ccy2 = value;
                SetRate();
            }
        }


        public double Amount1
        {
            get
            {
                return Model.Amount1;
            }
            set
            {
                Model.Amount1 = value;
                Model.IsCcy1Fixed = true;
                DeriveAmount();
            }
        }

        public double Amount2
        {
            get
            {
                return Model.Amount2;
            }
            set
            {
                Model.Amount2 = value;
                Model.IsCcy1Fixed = false;
                DeriveAmount();
            }
        }

        public double Rate
        {
            get
            {
                return Model.Rate;
            }
            set
            {
                Model.Rate = value;
            }
        }

        public void SetRate()
        {
            var fx = Rates.Find(Ccy1, Ccy2);
            if (fx == null)
                return;
            Rate = fx.Pair.Ccy1 == Ccy1 ? fx.Ask : fx.Bid;
        }

        public void DeriveAmount()
        {
            if (Model.TryDeriveXcrossAmount())
                RaiseChanged(Model.IsCcy1Fixed ? "Amoun2" : "Amount1");
        }

        public void RaiseChanged(string name)
        {
            PropertyChanged(this, new PropertyChangedEventArgs((name)));
        }
    }
}

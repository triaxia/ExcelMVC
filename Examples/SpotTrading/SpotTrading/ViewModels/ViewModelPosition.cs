namespace SpotTrading.ViewModels
{
    using System.ComponentModel;
    using BusinessModels;

    public class ViewModelPosition : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged = delegate { }; 

        public Position Model { get; private set; }

        public ViewModelPosition(Position model)
        {
            Model = model;
        }

        public void Net(Deal deal, ExchangeRates rates, bool raiseChanged)
        {
            Model.Net(deal, rates);
            if (raiseChanged)
            {
                PropertyChanged(this, new PropertyChangedEventArgs("Model.Amount1"));
                PropertyChanged(this, new PropertyChangedEventArgs("Model.Amount2"));
                PropertyChanged(this, new PropertyChangedEventArgs("Model.BaseAmount"));
            }
        }
    }
}

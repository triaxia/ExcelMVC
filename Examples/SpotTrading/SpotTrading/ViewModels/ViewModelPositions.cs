namespace SpotTrading.ViewModels
{
    using System.Collections.Generic;
    using System.Collections.Specialized;
    using System.Linq;
    using BusinessModels;

    public class ViewModelPositions :  List<ViewModelPosition>, INotifyCollectionChanged
    {
        public event NotifyCollectionChangedEventHandler CollectionChanged = delegate { };

        public Positions Model { get; private set; }

        public void Net(Deal deal, ExchangeRates rates)
        {
            var item = this.FirstOrDefault(x => CcyPair.IsMatched(deal.BuyCcy, deal.SellCcy, x.Model.Ccy1, x.Model.Ccy2));
            if (item == null)
            {
                item = new ViewModelPosition(new Position());
                item.Net(deal, rates, false);
                Add(item);
                CollectionChanged(this, new NotifyCollectionChangedEventArgs(NotifyCollectionChangedAction.Reset));
            }
            else
            {
                item.Net(deal, rates, true);
            }
        }
    }
}

namespace SpotTrading.CmdModels
{
    using System;
    using System.Windows.Input;
    using ExcelMvc.Controls;
    using ViewModels;

    public class CmdModelDeal:  ICommand
    {
        private ViewModelDeal Deal { get; set; }
        private ViewModelPositions Positions { get; set; }
        private ViewModelExchangeRates Rates { get; set; }

        public CmdModelDeal(ViewModelDeal deal, ViewModelPositions positions, ViewModelExchangeRates rates)
        {
            Deal = deal;
            Deal.PropertyChanged += Deal_PropertyChanged;
            Positions = positions;
            Rates = rates;
            CanExecuteChanged(this, null);
        }

        void Deal_PropertyChanged(object sender, System.ComponentModel.PropertyChangedEventArgs e)
        {
            CanExecuteChanged(this, null);
        }

        public bool CanExecute(object parameter)
        {
            return Deal.Model.BuyCcy != null
                && Deal.Model.SellCcy != null
                && Deal.Model.BuyCcy != Deal.Model.SellCcy;
        }

        public event EventHandler CanExecuteChanged = delegate { };

        public void Execute(object parameter)
        {
            var args = parameter as CommandEventArgs;
            Positions.Net(Deal.Model, Rates.Model);
        }
    }
}
namespace FXSpotTrading.CmdModels
{
    using System;
    using System.Windows.Input;
    using ExcelMvc.Controls;
    using ViewModels;

    public class CmdModelDeal:  ICommand
    {
        private ViewModelDeal Model { get; set; }

        public CmdModelDeal(ViewModelDeal model)
        {
            Model = model;
        }

        public bool CanExecute(object parameter)
        {
            var args = parameter as CommandEventArgs;
            return args.Source.State == null;
        }

        public event EventHandler CanExecuteChanged = delegate { };

        public void Execute(object parameter)
        {
            var args = parameter as CommandEventArgs;
        }
    }
}
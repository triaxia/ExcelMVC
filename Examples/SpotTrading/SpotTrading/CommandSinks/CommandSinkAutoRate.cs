namespace SpotTrading.CommandSinks
{
    using System;
    using System.Windows.Input;
    using ExcelMvc.Controls;
    using ViewModels;

    public class CommandSinkAutoRate :  ICommand
    {
        private ViewModelExchangeRates Model { get; set; }

        public CommandSinkAutoRate(ViewModelExchangeRates rates)
        {
            Model = rates;
        }

        public bool CanExecute(object parameter)
        {
            // toggling betweeen start and stop
            return true;
        }

        public event EventHandler CanExecuteChanged = delegate { };

        public void Execute(object parameter)
        {
            var args = parameter as CommandEventArgs;
            ExecuteAutoRate(args.Source);
        }

        private void ExecuteAutoRate(Command cmd)
        {
            if (cmd.State == null)
            {
                cmd.Caption = "Stop Simulation";
                cmd.State = 1;
                Model.StartSimulate();
            }
            else
            {
                cmd.Caption = "Start Simulation";
                cmd.State = null;
                Model.StopSimulate();
            }
        }
    }
}
namespace FXSpotTrading.CmdModels
{
    using System;
    using System.Windows.Input;
    using ExcelMvc.Controls;
    using ViewModels;

    public class CmdModelAutoRate :  ICommand
    {
        private ViewModelExchangeRates Model { get; set; }

        public CmdModelAutoRate(ViewModelExchangeRates rates)
        {
            Model = rates;
        }

        public bool CanExecute(object parameter)
        {
            //var args = parameter as CommandEventArgs;
            //return args.Source.State == null;

            // always enabled as we are toggling the status
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
                cmd.State = "Started";
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
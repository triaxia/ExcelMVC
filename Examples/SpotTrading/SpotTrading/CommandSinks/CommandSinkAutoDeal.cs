namespace SpotTrading.CommandSinks
{
    using System;
    using System.Windows.Input;
    using ExcelMvc.Controls;
    using ViewModels;

    public class CommandSinkAutoDeal :  ICommand
    {
        private ViewModelDealing Model { get; set; }

        public CommandSinkAutoDeal(ViewModelDealing deals)
        {
            Model = deals;
            CanExecuteChanged(this, new EventArgs());
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
            ExecuteAutoDeal(args.Source);
        }

        private void ExecuteAutoDeal(Command cmd)
        {
            if (cmd.State == null)
            {
                cmd.Caption = "Stop Dealing";
                cmd.State = 1;
                Model.StartSimulate();
            }
            else
            {
                cmd.Caption = "Start Dealing";
                cmd.State = null;
                Model.StopSimulate();
            }
        }
    }
}
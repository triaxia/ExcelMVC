namespace FXSpotTrading.ViewModels
{
    using System;
    using System.Collections.Generic;
    using System.Collections.Specialized;
    using System.Linq;
    using System.Threading;
    using System.Windows.Input;
    using ExcelMvc.Controls;
    using ExcelMvc.Runtime;
    using Models;

    internal class ViewModelExchangeRates : List<ViewModelExchangeRate>, INotifyCollectionChanged, ICommand
    {
        private ManualResetEvent AutoUpDateEvent { get; set; }
        public event NotifyCollectionChangedEventHandler CollectionChanged = delegate { };

        public ExchangeRates Model { get; private set; }

        public ViewModelExchangeRates(ExchangeRates rates)
        {
            Model = rates;
            Create();
        }

        private void Create()
        {
            Clear();
            AddRange(Model.Select(x => new ViewModelExchangeRate { Model = x }));
            CollectionChanged(this, new NotifyCollectionChangedEventArgs(NotifyCollectionChangedAction.Reset));
        }

        public bool CanExecute(object parameter)
        {
            return true;
        }

        public event System.EventHandler CanExecuteChanged = delegate { };

        public void Execute(object parameter)
        {
            var args = parameter as CommandEventArgs;
            if (args == null)
                return;

            if (args.Source.Name == "AutoRate")
                ExecuteAutoRate(args.Source);
        }

        private void ExecuteAutoRate(Command cmd)
        {
            if (cmd.Value == null)
            {
                cmd.Caption = "Stop Simulation";
                cmd.Value = "Started";
                StartSimulate();
            }
            else
            {
                cmd.Caption = "Start Simulation";
                cmd.Value = null;
                StopSimulate();
            }
        }

        private void StartSimulate()
        {
            AutoUpDateEvent = new ManualResetEvent(false);
            var thread = new Thread(Update) { Name = RangeUpdator.NameOfAsynUpdateThread, IsBackground = true };
            thread.Start();
        }

        private void StopSimulate()
        {
            if (AutoUpDateEvent != null)
                AutoUpDateEvent.Set();
        }

        private void Update(object state)
        {
            var random = new Random();
            while (!AutoUpDateEvent.WaitOne(1000))
            {
                var idx = (int)(random.NextDouble() * Count);
                if (idx >= Count) idx--;
                this[idx].Update();
            }
        }
    }
}
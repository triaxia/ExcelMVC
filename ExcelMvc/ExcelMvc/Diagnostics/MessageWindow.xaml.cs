namespace ExcelMvc.Diagnostics
{
    using System;
    using System.ComponentModel;
    using System.Runtime.CompilerServices;
    using System.Windows;
    using System.Windows.Input;
    using Runtime;

    /// <summary>
    /// Implements a visual sink for exception and information messages
    /// </summary>
    public partial class MessageWindow
    {
        private MessageWindow()
        {
            InitializeComponent();
            Closed += MessageWindow_Closed;
            Closing += MessageWindow_Closing;
            LayoutRoot.DataContext = Messages.Instance;
        }

        private static MessageWindow Instance { get; set; }

        /// <summary>
        /// Creates and shows to the status window
        /// </summary>
        [MethodImpl(MethodImplOptions.Synchronized)]
        public static void ShowInstance()
        {
            AsyncActions.Post(
                state =>
                {
                    CreateInstance();
                    // var interop = new WindowInteropHelper(Instance) { Owner = App.Instance.MainWindow.Handle };
                    Instance.Show();
                },
                null,
                false);
        }

        /// <summary>
        /// Hides the singleton
        /// </summary>
        [MethodImpl(MethodImplOptions.Synchronized)]
        public static void HideInstance()
        {
            AsyncActions.Post(
                state =>
                { 
                    Instance?.Hide(); 
                },
                null,
                false);
        }

        private static void CreateInstance()
        {
            Instance = Instance ?? new MessageWindow();
        }

        private static void MessageWindow_Closing(object sender, CancelEventArgs e)
        {
            if (ReferenceEquals(sender, Instance))
            {
                e.Cancel = true;
                HideInstance();
            }
        }

        private static void MessageWindow_Closed(object sender, EventArgs e)
        {
            if (ReferenceEquals(sender, Instance))
                Instance = null;
        }

        private void ButtonClear_OnClick(object sender, RoutedEventArgs e)
        {
            Messages.Instance.Clear();
        }

        private void ButtonHide_OnClick(object sender, RoutedEventArgs e)
        {
            Hide();
        }

        private void LineLimit_OnKeyDown(object sender, KeyEventArgs e)
        {
            var key = Convert.ToInt32(e.Key);
            e.Handled = (key < Convert.ToInt32(Key.D0) || key > Convert.ToInt32(Key.D9))
                     && (key < Convert.ToInt32(Key.NumPad0) || key > Convert.ToInt32(Key.NumPad9));
        }
    }
}

using System.ComponentModel;

namespace ExcelMvc.Diagnostics
{
    using System;
    using System.Runtime.CompilerServices;
    using System.Windows;
    using System.Windows.Input;

    /// <summary>
    /// Implements a visual sink for exception and information messages
    /// </summary>
    public partial class MessageWindow
    {
        #region Constructors
        private MessageWindow()
        {
            InitializeComponent();
            Closed += MessageWindow_Closed;
            Closing += MessageWindow_Closing;
        }

        #endregion 

        #region Properties
        private static MessageWindow Instance { get; set; }

        private Message Model
        {
            get { return (Message)LayoutRoot.DataContext; }
        }

        #endregion

        #region Methods

        /// <summary>
        /// Creates and shows the singleton
        /// </summary>
        [MethodImpl(MethodImplOptions.Synchronized)]
        public static void ShowInstance()
        {
            if (Instance == null)
                Instance = new MessageWindow();

            // var interop = new WindowInteropHelper(Instance) { Owner = App.Instance.MainWindow.Handle };
            Instance.Show();
        }

        /// <summary>
        /// Hides the singleton
        /// </summary>
        [MethodImpl(MethodImplOptions.Synchronized)]
        public static void HideInstance()
        {
            if (Instance != null)
                Instance.Hide();
        }

        /// <summary>
        /// Adds an exception to the singleton (only if it has been created by the ShowInstane method)
        /// </summary>
        /// <param name="ex">Exception to be addded</param>
        [MethodImpl(MethodImplOptions.Synchronized)]
        public static void AddErrorLine(Exception ex)
        {
            CreateInstance();
            Instance.Model.AddErrorLine(ex);
        }

        /// <summary>
        /// Adds an error to the singleton
        /// </summary>
        /// <param name="error">Error to be added</param>
        [MethodImpl(MethodImplOptions.Synchronized)]
        public static void AddErrorLine(string error)
        {
            CreateInstance();
            Instance.Model.AddErrorLine(error);
        }

        /// <summary>
        /// Adds a message to the singleton
        /// </summary>
        /// <param name="message">Message to be added</param>
        [MethodImpl(MethodImplOptions.Synchronized)]
        public static void AddInfoLine(string message)
        {
            CreateInstance();
            Instance.Model.AddInfoLine(message);
        }

        private static void CreateInstance()
        {
            if (Instance == null)
                Instance = new MessageWindow();
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
        #endregion

        private void ButtonClear_OnClick(object sender, RoutedEventArgs e)
        {
            Model.Clear();
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

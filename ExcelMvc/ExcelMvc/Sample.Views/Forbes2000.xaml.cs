using System.Collections;
using System.Windows;

namespace Sample.Views
{
    /// <summary>
    /// Interaction logic for Forbes.xaml
    /// </summary>
    public partial class Forbes2000 : Window
    {
        public Forbes2000()
        {
            InitializeComponent();
        }

        public IEnumerable Model
        {
            get 
            {
                return CompanyList.ItemsSource;
            }
            set
            {
                CompanyList.ItemsSource = value;
            }
        }
    }
}

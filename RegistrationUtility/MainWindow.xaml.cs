using System.Windows;
using System.Windows.Input;

namespace RegistrationUtility
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            DataContext = new ViewModel();
        }

        private void DLLTextBox_MouseEnter(object sender, MouseEventArgs e)
        {
            ((ViewModel)DataContext).FolderIconVisibility = Visibility.Visible;
        }

        private void DLLTextBox_MouseLeave(object sender, MouseEventArgs e)
        {
            ((ViewModel)DataContext).FolderIconVisibility = Visibility.Hidden;
        }
    }
}

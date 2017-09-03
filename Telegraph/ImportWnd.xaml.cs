using System.Windows;
using Telegraph.ViewModel;

namespace Telegraph
{
    /// <summary>
    /// Логика взаимодействия для ImportWnd.xaml
    /// </summary>
    public partial class ImportWnd : Window
    {
        public ImportWnd()
        {
            InitializeComponent();
            DataContext = ImportViewModel.SharedViewModel().InitPages();
        }
    }
}

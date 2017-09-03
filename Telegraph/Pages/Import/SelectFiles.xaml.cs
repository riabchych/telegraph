using System.Windows.Controls;
using Telegraph.ViewModel;

namespace Telegraph.Pages.Import
{
    /// <summary>
    /// Логика взаимодействия для ImportCustomFilesPage.xaml
    /// </summary>
    public partial class SelectFiles : Page
    {
        public SelectFiles()
        {
            InitializeComponent();
            DataContext = ImportViewModel.SharedViewModel();
        }
    }
}

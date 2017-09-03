using System.Windows.Controls;
using Telegraph.ViewModel;

namespace Telegraph.Pages.Import
{
    /// <summary>
    /// Логика взаимодействия для ImportPage.xaml
    /// </summary>
    public partial class SelectImportType : Page
    {
        public SelectImportType()
        {
            InitializeComponent();
            DataContext = ImportViewModel.SharedViewModel();
        }

    }
}

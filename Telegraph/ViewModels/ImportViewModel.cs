using System.Windows.Controls;

namespace Telegraph.ViewModel
{

    public class ImportViewModel: MainViewModel
    {
        private static ImportViewModel importViewModel;
        private RelayCommand showPageRestoreFromFiles;

        public static ImportViewModel SharedViewModel()
        {
            return importViewModel ?? (importViewModel = new ImportViewModel());
        }

        public ImportViewModel InitPages()
        {
            SelectImportType = new Pages.Import.SelectImportType();
            CurrentPage = SelectImportType;
            return this;
        }

        public RelayCommand ShowPageRestoreFromFiles
        {
            get
            {
                return showPageRestoreFromFiles ??
                (showPageRestoreFromFiles = new RelayCommand((o) => 
                {
                    SelectFiles = new Pages.Import.SelectFiles();
                    CurrentPage = SelectFiles;
                }));
            }
        }
    }
}
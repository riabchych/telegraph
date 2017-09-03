using System.Windows.Controls;

namespace Telegraph.ViewModels
{

    public class ImportViewModel: MainViewModel
    {
        private RelayCommand showPageRestoreFromFiles;

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
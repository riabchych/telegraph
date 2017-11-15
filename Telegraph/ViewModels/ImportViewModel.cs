using Microsoft.Practices.ServiceLocation;
using MvvmDialogs;
using MvvmDialogs.FrameworkDialogs.OpenFile;
using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Reflection;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;

namespace Telegraph.ViewModels
{

    public class ImportViewModel : MainViewModel
    {
        private Page selectFilesPage;
        private Page selectImportTypePage;
        private RelayCommand importFromFilesCommand;
        private RelayCommand importFromBackupCommand;
        private RelayCommand startImportCommand;
        private RelayCommand pauseImportCommand;
        private RelayCommand stopImportCommand;
        private RelayCommand cancelImportCommand;
        private RelayCommand goToBackCommand;
        private RelayCommand chooseFilesCommand;
        private RelayCommand onClose;
        private readonly IDialogService dialogService;
        private readonly ApplicationViewModel applicationViewModel;

        public ImportViewModel()
        {
            dialogService = ServiceLocator.Current.GetInstance<IDialogService>();
            applicationViewModel = ServiceLocator.Current.GetInstance<ApplicationViewModel>();
        }

        private ObservableCollection<string> Import(IProgress<int> progress)
        {
            ObservableCollection<string> failureList = new ObservableCollection<string>();

            for (int i = 0; i < Filenames.Count; i++)
            {
                progress.Report(i + 1);

                if (!applicationViewModel.ImportTelegrams(Filenames[i]))
                {
                    failureList.Add(Filenames[i]);
                }
            }
            return failureList;
        }

        public RelayCommand ImportFromFilesCommand
        {
            get => importFromFilesCommand ??
                (importFromFilesCommand = new RelayCommand((o) =>
                {
                    CurrentPage = SelectFilesPage = new Pages.Import.SelectFilesPage();
                }));
        }

        public RelayCommand ImportFromBackupCommand
        {
            get => importFromBackupCommand ??
                (importFromBackupCommand = new RelayCommand((o) =>
                {
                    CurrentPage = SelectFilesPage = new Pages.Import.SelectFilesPage();
                }));
        }

        public RelayCommand StartImportCommand
        {
            get => startImportCommand ??
                (startImportCommand = new RelayCommand(async (o) =>
                {
                    if (Filenames == null || (ProgressMaximum = Filenames.Count + 1) <= 1)
                    {
                        return;
                    }
                    IsBusy = true;
                    Progress<int> progress = new Progress<int>(current => ProgressText = $"Імпортується телеграма {ProgressValue = current}/{ProgressMaximum}...");
                    Filenames = await Task.Factory.StartNew(() => Import(progress), TaskCreationOptions.LongRunning);
                    IsBusy = false;
                    dialogService.ShowMessageBox(this, Filenames.Count > 0 ? "Імпортування закінчено! Деякі телеграми імпортувати не вдалося."
                        : "Телеграми успішно імпортовані!", "Імпортування закінчено", MessageBoxButton.OK, MessageBoxImage.Information);
                }));
        }

        public RelayCommand PauseImportCommand
        {
            get => pauseImportCommand ??
                (pauseImportCommand = new RelayCommand((o) =>
                {
                }));
        }

        public RelayCommand StopImportCommand
        {
            get => stopImportCommand ??
                (stopImportCommand = new RelayCommand((o) =>
                {
                }));
        }

        public RelayCommand CancelImportCommand
        {
            get => cancelImportCommand ??
                (cancelImportCommand = new RelayCommand((o) =>
                {
                    applicationViewModel.ImportWnd.DialogResult = false;
                }));
        }

        public RelayCommand GoToBackCommand
        {
            get => goToBackCommand ??
                (goToBackCommand = new RelayCommand((o) =>
                {
                    CurrentPage = SelectImportTypePage;
                }));
        }

        public RelayCommand ChooseFilesCommand
        {
            get => chooseFilesCommand ??
                (chooseFilesCommand = new RelayCommand((o) =>
                {
                    OpenFileDialogSettings settings = new OpenFileDialogSettings
                    {
                        Title = "Обрання файлів",
                        Multiselect = true,
                        InitialDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location),
                        Filter = "Документи Word|*.doc;*.docx|Всі файли|*.*"
                    };

                    if (dialogService.ShowOpenFileDialog(this, settings) == true)
                    {
                        Filenames = new ObservableCollection<string>(settings.FileNames);
                    }
                }));
        }

        public RelayCommand OnClose
        {
            get => onClose ??
                (onClose = new RelayCommand((o) =>
                {
                    applicationViewModel.RefreshViewSource();
                }));
        }

        public ObservableCollection<string> Filenames
        {
            get { return GetValue(() => Filenames); }
            set { SetValue(() => Filenames, value); }
        }

        public Page CurrentPage
        {
            get { return GetValue(() => CurrentPage); }
            set { SetValue(() => CurrentPage, value); }
        }

        public new bool IsBusy
        {
            get { return GetValue(() => IsBusy); }
            set { SetValue(() => IsBusy, value); }
        }

        public string ProgressText
        {
            get { return GetValue(() => ProgressText); }
            set { SetValue(() => ProgressText, value); }
        }

        public int ProgressMaximum
        {
            get { return GetValue(() => ProgressMaximum); }
            set { SetValue(() => ProgressMaximum, value); }
        }

        public int ProgressValue
        {
            get { return GetValue(() => ProgressValue); }
            set { SetValue(() => ProgressValue, value); }
        }

        public Page SelectFilesPage { get => selectFilesPage; set => selectFilesPage = value; }
        public Page SelectImportTypePage { get => selectImportTypePage; set => selectImportTypePage = value; }
    }
}
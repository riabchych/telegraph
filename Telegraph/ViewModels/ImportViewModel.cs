using Microsoft.Practices.ServiceLocation;
using MvvmDialogs;
using MvvmDialogs.FrameworkDialogs.OpenFile;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Reflection;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;

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
        private ManualResetEvent _busy;
        private BackgroundWorker importBackgroundWorker;


        public ImportViewModel()
        {
            dialogService = ServiceLocator.Current.GetInstance<IDialogService>();
            applicationViewModel = ServiceLocator.Current.GetInstance<ApplicationViewModel>();
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
                (startImportCommand = new RelayCommand((o) =>
                {
                    if (Filenames == null || Filenames.Count <= 0)
                    {
                        return;
                    }

                    ImportBackgroundWorker = new BackgroundWorker();
                    _busy = new ManualResetEvent(false);
                    ImportBackgroundWorker.DoWork += Import_DoWork;
                    ImportBackgroundWorker.ProgressChanged += ImportDoWork_ProgressChanged;
                    ImportBackgroundWorker.RunWorkerCompleted += ImportDoWork_Completed;
                    ImportBackgroundWorker.WorkerSupportsCancellation = true;
                    ImportBackgroundWorker.WorkerReportsProgress = true;
                    ImportBackgroundWorker.RunWorkerAsync();
                    
                    _busy.Set();
                }));
        }

        public RelayCommand PauseImportCommand
        {
            get => pauseImportCommand ??
                (pauseImportCommand = new RelayCommand((o) =>
                {
                    if (!_busy.WaitOne(0))
                    {
                        _busy.Set();
                    }
                    else
                    {
                        _busy.Reset();
                    }

                }));
        }

        public RelayCommand StopImportCommand
        {
            get => stopImportCommand ??
                (stopImportCommand = new RelayCommand((o) =>
                {
                    _busy.Set();
                    if (ImportBackgroundWorker.IsBusy)
                    {
                        ImportBackgroundWorker.CancelAsync();
                        ImportBackgroundWorker.Dispose();
                        _busy.Dispose();
                    }
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
                    var settings = new OpenFileDialogSettings
                    {
                        Title = "Обрання файлів",
                        Multiselect = true,
                        InitialDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location),
                        Filter = "Документи Word|*.doc;*.docx|Всі файли|*.*"
                    };

                    if (dialogService.ShowOpenFileDialog(this, settings) == true)
                    {
                        Filenames = new List<string>(settings.FileNames);
                    }
                }));
        }

        public RelayCommand OnClose
        {
            get => onClose ??
                (onClose = new RelayCommand((o) =>
                {
                    _busy.Set();
                    if (ImportBackgroundWorker == null)
                    {
                        return;
                    }
                    else if (ImportBackgroundWorker.IsBusy)
                    {
                        ImportBackgroundWorker.CancelAsync();
                    }
                    ImportBackgroundWorker.Dispose();
                    _busy.Dispose();
                }));
        }

        private void ImportDoWork_Completed(object sender, RunWorkerCompletedEventArgs e)
        {
            IsBusy = false;
            applicationViewModel.RefreshViewSource();
            dialogService.ShowMessageBox(this, Filenames.Count > 0 ? "Імпортування закінчено! Деякі телеграми імпортувати не вдалося."
                : "Телеграми успішно імпортовані!", "Імпортування закінчено", MessageBoxButton.OK, MessageBoxImage.Information);
            ImportBackgroundWorker.CancelAsync();
        }

        private void ImportDoWork_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            ProgressValue = e.ProgressPercentage;
            ProgressText = string.Format("Імпортується телеграма {0}/{1}...", e.ProgressPercentage, ProgressMaximum);
        }

        private void Import_DoWork(object sender, DoWorkEventArgs e)
        {
            IsBusy = true;
            int i = 0;
            ProgressMaximum = Filenames.Count;
            string[] files = new string[Filenames.Count];
            Filenames.CopyTo(files);

            foreach (string filename in files)
            {
                _busy.WaitOne();
                ImportBackgroundWorker.ReportProgress(percentProgress: i++);

                if (ImportBackgroundWorker.CancellationPending)
                {
                    break;
                }

                Application.Current.Dispatcher.Invoke(priority: DispatcherPriority.Normal, method: (Action)delegate ()
                {
                    if (applicationViewModel.ImportTelegrams(filename))
                    {
                        Filenames.RemoveAt(Filenames.IndexOf(filename));
                    }
                });
            }
        }

        public List<string> Filenames
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
        public BackgroundWorker ImportBackgroundWorker { get => importBackgroundWorker; set => importBackgroundWorker = value; }
    }
}
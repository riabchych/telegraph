using GalaSoft.MvvmLight.Messaging;
using IFilterTextReader;
using Microsoft.Practices.ServiceLocation;
using Microsoft.Win32;
using MvvmDialogs;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Threading;
using Telegraph.LogModule.Loggers;

namespace Telegraph.ViewModels
{
    public class ApplicationViewModel : MainViewModel
    {
        private static ApplicationViewModel applicationViewModel;
        private readonly ITelegramDataService dataService;
        private readonly IDialogService dialogService;
        private readonly ILogger logger;
        private ImportWnd importWnd;
        private RelayCommand newCommand;
        private RelayCommand addCommand;
        private RelayCommand editCommand;
        private RelayCommand cutCommand;
        private RelayCommand copyCommand;
        private RelayCommand deleteCommand;
        private RelayCommand saveCommand;
        private RelayCommand sendToWord;
        private RelayCommand printCommand;
        private RelayCommand importCommand;
        private RelayCommand selectedTelegramsChanged;

        public static ApplicationViewModel SharedViewModel()
        {
            return applicationViewModel ?? (applicationViewModel = new ApplicationViewModel());
        }

        public string FilterText
        {
            get { return GetValue(() => FilterText); }
            set
            {
                SetValue(() => FilterText, value);
                TelegramsViewSource.View.Refresh();
            }
        }

        public string StatusText
        {
            get { return GetValue(() => StatusText); }
            set { SetValue(() => StatusText, value); }
        }

        public int FilterType
        {
            get { return GetValue(() => FilterType); }
            set { SetValue(() => FilterType, value); }
        }

        public bool IsImport
        {
            get { return GetValue(() => IsImport); }
            set { SetValue(() => IsImport, value); }
        }

        public IList<Telegram> SelectedTelegrams
        {
            get { return GetValue(() => SelectedTelegrams); }
            set { SetValue(() => SelectedTelegrams, value); }
        }

        public Telegram ActiveTelegram
        {
            get { return GetValue(() => ActiveTelegram); }
            set { SetValue(() => ActiveTelegram, value); }
        }

        public bool IsSelected
        {
            get { return GetValue(() => IsSelected); }
            set { SetValue(() => IsSelected, value); }
        }

        public ImportWnd ImportWnd { get => importWnd; set => importWnd = value; }

        public ApplicationViewModel()
        {
            dialogService = ServiceLocator.Current.GetInstance<IDialogService>();
            dataService = ServiceLocator.Current.GetInstance<ITelegramDataService>();
            logger = ServiceLocator.Current.GetInstance<ILogger>();
            TelegramsViewSource = new CollectionViewSource();
            TelegramsViewSource.Filter += TelegramsFilter;
            LoadDataAsync();
        }

        private async void LoadDataAsync()
        {
            StatusText = "Відбувається завантаження даних. Будь ласка, зачекайте ...";
            IsBusy = true;
            Telegrams = await Task.Factory.StartNew(() => dataService.LoadTelegrams(), TaskCreationOptions.LongRunning);
            RefreshViewSource(Telegrams);
            IsBusy = false;
        }

        public void RefreshViewSource()
        {
            Refresh();
        }

        private void RefreshViewSource(ICollection<Telegram> telegrams)
        {
            Refresh(telegrams);
        }

        private void Refresh(ICollection<Telegram> telegrams = null)
        {
            TelegramsViewSource.Dispatcher.Invoke(() =>
            {
                TelegramsViewSource.Source = new ObservableCollection<Telegram>(telegrams ?? Telegrams);
                StatusText = $"Всього телеграм: {Telegrams.Count}";
            });
        }

        private void TelegramsFilter(object sender, FilterEventArgs e)
        {
            if (string.IsNullOrEmpty(FilterText))
            {
                e.Accepted = true;
                return;
            }

            object result = null;

            if (e.Item is Telegram tlg)
            {
                switch (FilterType)
                {
                    case 0:
                        result = tlg.SelfNum;
                        break;
                    case 1:
                        result = tlg.IncNum;
                        break;
                    case 2:
                        result = tlg.From;
                        break;
                    case 3:
                        result = tlg.To;
                        break;
                    case 4:
                        result = tlg.Text;
                        break;
                    case 5:
                        result = tlg.SubNum;
                        break;
                    case 6:
                        result = tlg.Date;
                        break;
                    case 7:
                        result = tlg.SenderPos;
                        break;
                    case 8:
                        result = tlg.SenderRank;
                        break;
                    case 9:
                        result = tlg.SenderName;
                        break;
                    case 10:
                        result = tlg.Executor;
                        break;
                    case 11:
                        result = tlg.Phone;
                        break;
                    case 12:
                        result = tlg.HandedBy;
                        break;
                    case 13:
                        result = tlg.Dispatcher;
                        break;
                    default:
                        result = null;
                        break;
                }
            }

            e.Accepted = result == null || result.ToString().ToUpper().Contains(FilterText.ToUpper());
        }

        private void ShowError(string text)
        {
            logger.Error(text);
            Dispatcher.CurrentDispatcher.Invoke(() => dialogService.ShowMessageBox(
              this,
              text,
              "Помилка",
              MessageBoxButton.OK,
              MessageBoxImage.Error));
        }

        public RelayCommand NewCommand => newCommand ?? (newCommand = new RelayCommand(o => OpenTlgWnd()));

        public RelayCommand SaveCommand
        {
            get
            {
                return saveCommand ??
                    (saveCommand = new RelayCommand((o) =>
                    {
                        if (o is TelegramWnd telegramWindow)
                            telegramWindow.DialogResult = true;
                    }));
            }
        }

        public RelayCommand SelectedTelegramsChanged
        {
            get
            {
                return selectedTelegramsChanged ??
                    (selectedTelegramsChanged = new RelayCommand((o) =>
                    {
                        SelectedTelegrams = ConvertListOfSelectedItem<Telegram>(o);
                        ActiveTelegram = (IsSelected = (SelectedTelegrams.Count == 1)) ? SelectedTelegrams[0] : null;
                    }));
            }
        }

        public RelayCommand EditCommand
        {
            get
            {
                return editCommand ??
                  (editCommand = new RelayCommand((o) =>
                  {
                      if (ActiveTelegram is Telegram tlg)
                      {
                          ActiveTelegram = tlg.Clone() as Telegram;
                          OpenTlgWnd(ActiveTelegram);
                      }
                  }));
            }
        }

        public RelayCommand CutCommand => cutCommand ?? (cutCommand = new RelayCommand((o) => CopySelectedTelegramsAsync(true)));

        public RelayCommand CopyCommand => copyCommand ?? (copyCommand = new RelayCommand((o) => CopySelectedTelegramsAsync()));

        private async void CopySelectedTelegramsAsync(bool removable = false)
        {
            IsBusy = true;
            StatusText = "Відбувається копіювання телеграм...";
            logger.Debug(StatusText);

            try
            {
                StringCollection files = new StringCollection();

                foreach (Telegram tlg in SelectedTelegrams)
                {
                    string fileName = $"{Path.GetTempPath()}{tlg.SelfNum}.docx";
                    new WordDocument(tlg).CreatePackage(fileName);
                    files.Add(fileName);
                    logger.Debug($"\"{fileName}\" створений успішно.");
                }
                DataObject data = new DataObject();
                data.SetFileDropList(files);
                data.SetData("Preferred DropEffect", DragDropEffects.Move);
                Application.Current.Dispatcher.Invoke(DispatcherPriority.Normal, (Action)(() =>
                {
                    Clipboard.Clear();
                    Clipboard.SetDataObject(data, true);
                    logger.Debug($"Обрані телеграми успішно занесені в буфер обміну.");
                }));

                if (removable && await Task.Factory.StartNew(() => dataService.RemoveTelegram(SelectedTelegrams), TaskCreationOptions.LongRunning))
                {
                    RefreshViewSource();
                }
                logger.Info("Копіювання успішно закінчено!");
                IsBusy = false;
            }
            catch (Exception e)
            {
                logger.Error(e.Message);
                ShowError("Виникла помилка при копіюванні телеграм.");
            }
        }

        public RelayCommand AddCommand
        {
            get
            {
                return addCommand ??
                  (addCommand = new RelayCommand((o) =>
                  {
                      string[] files = null;

                      if (o == null)
                      {
                          OpenFileDialog dlg = new OpenFileDialog()
                          {
                              DefaultExt = ".doc",
                              Filter = "Документи Word|*.doc;*.docx|Всі файли|*.*",
                              CheckFileExists = true,
                              Multiselect = true
                          };

                          bool? result = dlg.ShowDialog();

                          if (result == true)
                          {
                              files = dlg.FileNames;
                              ImportTelegrams(files);
                          }
                      }
                      else
                      {
                          DragEventArgs e = o as DragEventArgs;
                          files = (string[])e.Data.GetData(DataFormats.FileDrop);
                          ImportTelegrams(files);
                      }
                  }));
            }
        }

        public RelayCommand DeleteCommand
        {
            get
            {
                return deleteCommand ??
                    (deleteCommand = new RelayCommand(async (o) =>
                    {
                        StatusText = "Видалення телеграм...";
                        IsBusy = true;
                        if (await Task.Factory.StartNew(() => dataService.RemoveTelegram(SelectedTelegrams), TaskCreationOptions.LongRunning))
                        {
                            RefreshViewSource();
                        }
                        IsBusy = false;
                    }));
            }
        }

        public RelayCommand SendToWord
        {
            get
            {
                return sendToWord ??
                    (sendToWord = new RelayCommand((o) =>
                    {
                        try
                        {
                            if (ActiveTelegram == null)
                                return;

                            if (!isMSWordInstalled())
                            {
                                ShowError("Виникла помилка при відправці документа в Microsoft Office Word. Microsoft Office Word не інстальований у Вашій системі.");
                                return;
                            }
                            string fileName = GetTempFile("docx");
                            new WordDocument(ActiveTelegram).CreatePackage(fileName);
                            logger.Debug($"\"{fileName}\" створений успішно.\n\"{fileName}\" відкриваєтья в Microsoft Office Word...");
                            Process.Start(fileName);
                        }
                        catch (Exception e)
                        {
                            logger.Error(e.Message);
                            ShowError("Виникла помилка при відправці документа в Microsoft Office Word.");
                        }
                    }));
            }
        }

        public RelayCommand PrintCommand
        {
            get
            {
                return printCommand ??
                    (printCommand = new RelayCommand((o) =>
                    {
                        try
                        {
                            string fileName = GetTempFile("docx");
                            new WordDocument(ActiveTelegram).CreatePackage(fileName);
                            logger.Debug($"\"{fileName}\" створений успішно.");
                            PrintDialog printDialog = new PrintDialog();
                            logger.Debug("Відкривається вікно друку...");
                            if (printDialog.ShowDialog() == true)
                            {
                                ProcessStartInfo info = new ProcessStartInfo(fileName)
                                {
                                    CreateNoWindow = true,
                                    WindowStyle = ProcessWindowStyle.Hidden,
                                    UseShellExecute = true,
                                    Verb = "PrintTo"
                                };
                                Process.Start(info);
                                logger.Info($"Виконується друк документу \"{fileName}\".");
                            }
                        }
                        catch (Exception e)
                        {
                            logger.Error(e.Message);
                            ShowError("Виникла помилка при друкуванні документу.");
                        }
                    }));
            }
        }

        public RelayCommand ImportCommand
        {
            get
            {
                return importCommand ??
                    (importCommand = new RelayCommand((wnd) =>
                    {
                        try
                        {
                            ServiceLocator.Current.GetInstance<ImportViewModel>().SelectImportTypePage =
                                new Pages.Import.SelectImportTypePage();
                            ServiceLocator.Current.GetInstance<ImportViewModel>().CurrentPage =
                                ServiceLocator.Current.GetInstance<ImportViewModel>().SelectImportTypePage;

                            ImportWnd = new ImportWnd();
                            logger.Debug("Відкривається вікно імпорту...");

                            if (ImportWnd.ShowDialog() == true)
                            {
                                throw new NotImplementedException("Метод не реализован");
                            }
                            ImportWnd = null;
                        }
                        catch (Exception e)
                        {
                            logger.Error(e.Message);
                            ShowError("Виникла помилка при імпортуванні документу.");
                        }
                    }));
            }
        }

        public bool ImportTelegrams(string filePath)
        {
            if ((ActiveTelegram = HandleTelegrams(filePath, true)) == null)
            {
                Dispatcher.CurrentDispatcher.Invoke(() => Messenger.Default.Send(ImportWnd));
                return false;
            }
            else
            {
                if (!CheckData(ActiveTelegram))
                {
                    ShowError($"Не вдалося отримати деяку інформацію, перевірте структуру документа \"{filePath}\".");
                    bool? result = OpenTlgWnd(ActiveTelegram);
                    Dispatcher.CurrentDispatcher.Invoke(() => Messenger.Default.Send(ImportWnd));
                    return result.Value;
                }

                return dataService.AddTelegram(ActiveTelegram);
            }
        }

        public bool ImportTelegrams(string[] files)
        {
            if (files == null)
                return false;

            foreach (string filePath in files)
            {
                if ((ActiveTelegram = HandleTelegrams(filePath)) == null)
                {
                    return false;
                }
                else
                {
                    if (!CheckData(ActiveTelegram))
                    {
                        ShowError($"Не вдалося отримати деяку інформацію, перевірте структуру документа \"{filePath}\".");
                    }

                    bool? result = OpenTlgWnd(ActiveTelegram);
                    RefreshViewSource();
                    return result.Value;
                }
            }
            return false;
        }

        private Telegram HandleTelegrams(string filePath, bool isImport = false)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
                {
                    throw new FileNotFoundException("Неможливо імпортувати телеграму, файл з даним іменем не існує.", filePath);
                }
                else
                {
                    string ext = Path.GetExtension(filePath);

                    if (!(ext.Equals(".doc") || ext.Equals(".docx")))
                    {
                        throw new FileFormatException($"Формат файлу \"{filePath}\" нажаль не підтримується.");
                    }
                }
            }
            catch (Exception e)
            {
                if (!isImport) ShowError(e.Message);
                return null;
            }
            try
            {
                using (TextReader reader = new FilterReader(filePath))
                {
                    string docText = null;

                    try
                    {
                        docText = reader.ReadToEnd();
                    }
                    catch
                    {
                        throw new FileFormatException($"Помилка при зчитуванні інформації з файлу \"{filePath}\".");
                    }

                    docText = docText.ToUpper();

                    if (string.IsNullOrWhiteSpace(docText))
                    {
                        throw new FileFormatException($"Файл \"{filePath}\" не містить даних.");
                    }

                    string regexString = string.Concat(TlgRegex.MainRegex, TlgRegex.FooterRegex);
                    Regex regex = null;
                    GroupCollection groups = null;
                    bool isUrgency = false;

                    try
                    {
                        regex = new Regex(regexString, RegexOptions.Multiline, TimeSpan.FromSeconds(3));
                        groups = regex.Match(input: docText).Groups;

                        if (groups.Count < 2)
                        {
                            throw new InvalidDataException();
                        }

                        isUrgency = new Regex(TlgRegex.UrgencyRegex).IsMatch(docText);
                    }
                    catch
                    {
                        throw new FileFormatException($"Не вдалося отримати інформацію, перевірте чи не порушена структура документу \"{filePath}\".");
                    }

                    return ActiveTelegram = new Telegram()
                    {
                        SelfNum = Convert.ToInt32(value: groups[1].Value.Trim() ?? "0"),
                        From = groups[2].Value.Trim(),
                        IncNum = Convert.ToInt32(value: groups[3].Value.Trim() ?? "0"),
                        To = groups[4].Value.Trim(),
                        Text = groups[5].Value.Trim(),
                        SubNum = groups[6].Value.Trim(),
                        Date = groups[8].Value.Trim(),
                        SenderPos = string.Concat(str0: groups[7].Value.Trim(), str1: " ", str2: groups[9].Value.Trim()),
                        SenderRank = groups[10].Value.Trim(),
                        SenderName = groups[11].Value.Trim(),
                        Executor = groups[13].Value.Trim(),
                        Phone = groups[14].Value.Trim(),
                        HandedBy = groups[16].Value.Trim(),
                        Urgency = isUrgency ? 1 : 0,
                        Dispatcher = groups[19].Value.Trim(),
                        Time = groups[20].Value.Trim()
                    };
                }
            }
            catch (Exception e)
            {
                if (!isImport) ShowError(e.Message);
                return null;
            }
        }

        private bool? OpenTlgWnd(Telegram tlg = null)
        {
            if (tlg == null)
            {
                tlg = new Telegram();
            }

            bool isNew = false;
            ActiveTelegram = tlg;

            if (ActiveTelegram.Id < 1)
            {
                ActiveTelegram.Time = DateTime.Now.ToString(new CultureInfo("ru-RU"));
                isNew = true;
            }

            TelegramWnd TelegramWindow = new TelegramWnd();
            logger.Debug("Відкривається вікно телеграми...");

            if (TelegramWindow.ShowDialog() == true)
            {
                if (isNew)
                {
                    logger.Debug("Відбувається додання нової телеграми...");
                    dataService.AddTelegram(ActiveTelegram);
                }
                else
                {
                    logger.Debug($"Відбувається редагування телеграми з номером ID - {ActiveTelegram.Id}.");
                    dataService.EditTelegram(ActiveTelegram.Id, ActiveTelegram);
                }
            }
            else
            {
                ActiveTelegram = null;
            }
            return TelegramWindow.DialogResult.Value;
        }
    }
}

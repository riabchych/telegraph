using GalaSoft.MvvmLight.Messaging;
using IFilterTextReader;
using Microsoft.Practices.ServiceLocation;
using Microsoft.Win32;
using MvvmDialogs;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
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
        private RelayCommand deleteCommand;
        private RelayCommand saveCommand;
        private RelayCommand sendToWord;
        private RelayCommand printCommand;
        private RelayCommand importCommand;

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

        public Telegram ActiveTelegram
        {
            get { return GetValue(() => ActiveTelegram); }
            set { SetValue(() => ActiveTelegram, value); }
        }

        public ApplicationViewModel()
        {
            dialogService = ServiceLocator.Current.GetInstance<IDialogService>();
            dataService = ServiceLocator.Current.GetInstance<ITelegramDataService>();
            logger = ServiceLocator.Current.GetInstance<ILogger>();
            TelegramsViewSource = new CollectionViewSource();
            StatusText = "Завантаження даних...";
            IsBusy = true;
            dataService.LoadTelegrams(TelegramsLoaded);
        }

        private void TelegramsLoaded(IEnumerable<Telegram> telegrams)
        {
            Telegrams = telegrams;
            TelegramsViewSource.Filter += TelegramsFilter;
            RefreshViewSource(Telegrams);
            IsBusy = false;
        }

        private void ShowError(string text)
        {
            logger.Error(text);
            dialogService.ShowMessageBox(
              this,
              text,
              "Помилка",
              MessageBoxButton.OK,
              MessageBoxImage.Error);
        }

        public RelayCommand NewCommand
        {
            get
            {
                return newCommand ?? (newCommand = new RelayCommand((o) => OpenTlgWnd()));
            }
        }

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
                    (deleteCommand = new RelayCommand((o) =>
                    {
                        if (ActiveTelegram is Telegram)
                        {
                            dataService.RemoveTelegram(ActiveTelegram, RefreshViewSource);
                        }
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
                                throw new Exception(message: "Microsoft Office Word не інстальований у Вашій системі");
                            }
                            string fileName = GetTempFile("docx");
                            new WordDocument(ActiveTelegram).CreatePackage(fileName);
                            logger.Debug(String.Format("\"{0}\" створений успішно.", fileName));
                            logger.Info(String.Format("\"{0}\" відкриваєтья в Microsoft Office Word", fileName));
                            Process.Start(fileName);
                        }
                        catch (Exception e)
                        {
                            logger.Error(e.Message);
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
                        string fileName = GetTempFile("docx");
                        new WordDocument(ActiveTelegram).CreatePackage(fileName);
                        logger.Debug(String.Format("\"{0}\" створений успішно.", fileName));
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
                            logger.Info(String.Format("Виконується друк документу \"{0}\".", fileName));
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
                        ServiceLocator.Current.GetInstance<ImportViewModel>().SelectImportTypePage = 
                            new Pages.Import.SelectImportTypePage();
                        ServiceLocator.Current.GetInstance<ImportViewModel>().CurrentPage = 
                            ServiceLocator.Current.GetInstance<ImportViewModel>().SelectImportTypePage;

                        ImportWnd = new ImportWnd();

                        if (ImportWnd.ShowDialog() == true)
                        {

                        }
                        ImportWnd = null;
                    }));
            }
        }

        public ImportWnd ImportWnd { get => importWnd; set => importWnd = value; }

        public bool ImportTelegrams(string filePath)
        {
            if ((ActiveTelegram = HandleTelegrams(filePath, true)) == null)
            {
                Messenger.Default.Send(new ImportWnd());
                return false;
            }
            else
            {
                try
                {
                    if (!CheckData(ActiveTelegram))
                    {
                        throw new Exception(message: $"Не вдалося отримати деяку інформацію, перевірте структуру документа \"{filePath}\".");
                    }

                    dataService.AddTelegram(ActiveTelegram);
                    return true;
                }
                catch (Exception e)
                {
                    ShowError(e.Message);
                    bool result = OpenTlgWnd(ActiveTelegram);
                    Messenger.Default.Send(new ImportWnd());
                    return result;
                }
            }
        }

        public bool ImportTelegrams(string[] files)
        {
            if (files == null)
            {
                return false;
            }
            else
            {
                foreach (string filePath in files)
                {
                    if ((ActiveTelegram = HandleTelegrams(filePath)) == null)
                    {
                        return false;
                    }
                    else
                    {
                        try
                        {
                            if (!CheckData(ActiveTelegram))
                            {
                                throw new Exception(message: $"Не вдалося отримати деяку інформацію, перевірте структуру документа \"{filePath}\".");
                            }
                        }
                        catch (Exception e)
                        {
                            ShowError(e.Message);
                        }
                        return OpenTlgWnd(ActiveTelegram);
                    }
                }
                return true;
            }
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
                        return null;
                    }
                }
            }
            catch (Exception e)
            {
                if (isImport) return null;
                ShowError(e.Message);
                return null;
            }
            try
            {
                using (TextReader reader = new FilterReader(filePath))
                {
                    if ((reader.ReadToEnd() is string docText))
                    {
                        if (string.IsNullOrWhiteSpace(docText = docText.ToUpper()))
                        {
                            throw new Exception(message: $"Файл \"{filePath}\" не містить даних.");
                        }

                        string regexString = string.Concat(TlgRegex.MainRegex, TlgRegex.FooterRegex);
                        Regex regex = new Regex(regexString, RegexOptions.Multiline, TimeSpan.FromSeconds(3));
                        GroupCollection groups = null;
                        groups = regex.Match(input: docText).Groups;

                        if (groups.Count < 2)
                        {
                            throw new Exception(message: $"Не вдалося отримати інформацію, перевірте чи не порушена структура документу \"{filePath}\".");
                        }

                        bool isUrgency = new Regex(pattern: TlgRegex.UrgencyRegex).IsMatch(docText);

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
                    else
                    {
                        return null;
                    }

                }
            }
            catch (Exception e)
            {
                if (isImport) return null;
                ShowError(e.Message);
                return null;
            }
        }

        private bool OpenTlgWnd(Telegram tlg = null)
        {
            bool isNew = false;

            tlg = tlg ?? (tlg = new Telegram());

            if (tlg.Id < 1)
            {
                tlg.Time = DateTime.Now.ToString(new CultureInfo("ru-RU"));
                isNew = true;
            }

            TelegramWnd TelegramWindow = new TelegramWnd();
            logger.Debug("Відкривається вікно телеграми...");
            if (TelegramWindow.ShowDialog() == true)
            {
                if (isNew)
                {
                    logger.Debug("Відбувається додання нової телеграми...");
                    dataService.AddTelegram(ActiveTelegram, RefreshViewSource);
                    return true;
                }
                else
                {
                    logger.Debug(String.Format("Відбувається редагування телеграми з номером ID - {0}", ActiveTelegram.Id));
                    dataService.EditTelegram(ActiveTelegram.Id, ActiveTelegram, RefreshViewSource);
                    return true;
                }
            }
            else
            {
                ActiveTelegram = null;
                return false;
            }
        }

        public void RefreshViewSource()
        {
            Refresh();
        }

        private void RefreshViewSource(IEnumerable<Telegram> telegrams)
        {
            Refresh(telegrams);
        }

        private void Refresh(IEnumerable<Telegram> telegrams = null)
        {
            logger.Debug("Оновлення списку телеграм.");
            ObservableCollection<Telegram> collection;
            TelegramsViewSource.Source = collection = new ObservableCollection<Telegram>(telegrams ?? Telegrams);
            StatusText = String.Format("Всього телеграм: {0} ", collection.Count);
        }

        private void TelegramsFilter(object sender, FilterEventArgs e)
        {
            if (string.IsNullOrEmpty(FilterText))
            {
                e.Accepted = true;
                return;
            }

            string result = null;

            if (e.Item is Telegram tlg)
            {
                switch (FilterType)
                {
                    case 0:
                        result = tlg.SelfNum.ToString();
                        break;
                    case 1:
                        result = tlg.IncNum.ToString();
                        break;
                    case 2:
                        result = tlg.From.ToString();
                        break;
                    case 3:
                        result = tlg.To.ToString();
                        break;
                    case 4:
                        result = tlg.Text.ToString();
                        break;
                    case 5:
                        result = tlg.SubNum.ToString();
                        break;
                    case 6:
                        result = tlg.Date.ToString();
                        break;
                    case 7:
                        result = tlg.SenderPos.ToString();
                        break;
                    case 8:
                        result = tlg.SenderRank.ToString();
                        break;
                    case 9:
                        result = tlg.SenderName.ToString();
                        break;
                    case 10:
                        result = tlg.Executor.ToString();
                        break;
                    case 11:
                        result = tlg.Phone.ToString();
                        break;
                    case 12:
                        result = tlg.HandedBy.ToString();
                        break;
                    case 13:
                        result = tlg.Dispatcher.ToString();
                        break;
                    default:
                        result = null;
                        break;
                }
            }

            e.Accepted = result == null ?
                true : e.Accepted = result.ToUpper().Contains(FilterText.ToUpper()) ? true : false;
        }
    }
}

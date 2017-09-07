using IFilterTextReader;
using Microsoft.Win32;
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

namespace Telegraph.ViewModels
{

    public class ApplicationViewModel : MainViewModel
    {
        private static ApplicationViewModel applicationViewModel;
        private ITelegramDataService _dataService;

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

        public ApplicationViewModel() { }

        public ApplicationViewModel(ITelegramDataService dataService)
        {
            TelegramsViewSource = new CollectionViewSource();
            StatusText = "Завантаження даних...";
            _dataService = dataService;
            IsBusy = true;
            _dataService.LoadTelegrams(TelegramsLoaded, TelegramsLoadFiled);
        }

        private void TelegramsLoadFiled(Exception obj)
        {
            throw new NotImplementedException();
        }

        private void TelegramsLoaded(IEnumerable<Telegram> telegrams)
        {
            Telegrams = telegrams;
            TelegramsViewSource.Filter += TelegramsFilter;
            RefreshViewSource(Telegrams);
            IsBusy = false;
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
                              GetTelegrams(files);
                          }
                      }
                      else
                      {
                          DragEventArgs e = o as DragEventArgs;
                          files = (string[])e.Data.GetData(DataFormats.FileDrop);
                          GetTelegrams(files);
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
                            _dataService.RemoveTelegram(ActiveTelegram, RefreshViewSource);
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
                        if (ActiveTelegram == null)
                            return;

                        string fileName = GetTempFile("docx");
                        new WordDocument(ActiveTelegram).CreatePackage(fileName);
                        Process.Start(fileName);
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

                        PrintDialog printDialog = new PrintDialog();

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
                        CurrentPage = new Pages.Import.SelectImportType();
                        ImportWnd importWnd = new ImportWnd();

                        if (importWnd.ShowDialog() == true)
                        {

                        }
                    }));
            }
        }

        private void GetTelegrams(string[] files)
        {
            if (files == null)
                return;

            foreach (string filePath in files)
            {
                try
                {
                    string ext = Path.GetExtension(filePath);
                    if (!(ext.Equals(".doc") || ext.Equals(".docx")))
                    {
                        continue;
                    }
                    else
                    {
                        using (TextReader reader = new FilterReader(filePath))
                        {
                            string docText = reader.ReadToEnd().ToUpper();

                            if (string.IsNullOrWhiteSpace(docText))
                                throw new Exception(message: "Документ якиий Ви намагаєтесь відкрити не містить данних.");

                            string regexString = string.Concat(TlgRegex.BasePartRegex, TlgRegex.FirstPartRegex);
                            Regex regex = new Regex(regexString, RegexOptions.Multiline, TimeSpan.FromSeconds(3));
                            GroupCollection groups = null;
                            try
                            {
                                groups = regex.Match(docText).Groups;
                            }
                            catch
                            {
                                throw new Exception(message: "Не вдалося отримати інформацію, перевірте чи не порушена структура документа.");
                            }
                            if (groups.Count < 2)
                            {
                                regexString = string.Concat(TlgRegex.BasePartRegex, TlgRegex.SecondPartRegex);
                                regex = new Regex(regexString, RegexOptions.Multiline);
                                groups = regex.Match(docText).Groups;
                                if (groups.Count < 2)
                                {
                                    throw new Exception(message: "Не вдалося отримати інформацію, перевірте чи не порушена структура документа.");
                                }
                            }

                            bool isUrgency = new Regex(TlgRegex.UrgencyRegex).IsMatch(docText);
                            string num = groups[2].Value.Trim().ToString();
                            int number = Int32.Parse((!string.IsNullOrWhiteSpace(num)) ? num : "0");

                            ActiveTelegram = new Telegram()
                            {
                                SelfNum = 0,
                                IncNum = number,
                                From = groups[1].Value.Trim(),
                                To = groups[3].Value.Trim(),
                                Text = groups[4].Value.Trim(),
                                SubNum = groups[5].Value.Trim(),
                                Date = groups[7].Value.Trim(),
                                SenderPos = string.Concat(groups[6].Value.Trim(), " ", groups[8].Value.Trim()),
                                SenderRank = groups[9].Value.Trim(),
                                SenderName = groups[10].Value.Trim(),
                                Executor = groups[11].Value.Trim(),
                                Phone = groups[12].Value.Trim(),
                                HandedBy = groups[13].Value.Trim(),
                                Urgency = isUrgency ? 1 : 0,
                                Dispatcher = string.Empty,
                                Time = string.Empty
                            };

                            if (CheckData(ActiveTelegram))
                            {
                                OpenTlgWnd(ActiveTelegram);
                            }
                            else
                            {
                                throw new Exception(message: "Не вдалося отримати деяку інформацію, перевірте структуру документа");
                            }
                        }
                    }
                }
                catch
                {

                }
            }
        }

        private void OpenTlgWnd(Telegram tlg = null)
        {
            bool isNew = false;

            tlg = tlg ?? (tlg = new Telegram());

            if (tlg.Id < 1)
            {
                tlg.Time = DateTime.Now.ToString(new CultureInfo("ru-RU"));
                isNew = true;
            }

            TelegramWnd TelegramWindow = new TelegramWnd();

            if (TelegramWindow.ShowDialog() == true)
            {
                if (isNew)
                {
                    _dataService.AddTelegram(ActiveTelegram, RefreshViewSource);
                }
                else
                {
                    _dataService.EditTelegram(ActiveTelegram.Id, ActiveTelegram, RefreshViewSource);
                }
            }
        }

        private void RefreshViewSource()
        {
            Refresh();
        }

        private void RefreshViewSource(IEnumerable<Telegram> telegrams)
        {
            Refresh(telegrams);
        }

        private void Refresh(IEnumerable<Telegram> telegrams = null)
        {
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

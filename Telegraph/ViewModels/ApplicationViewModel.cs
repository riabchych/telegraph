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
            FilterType = 3;
            TelegramsViewSource = new CollectionViewSource();
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
            IsBusy = false;
            Telegrams = telegrams;
            TelegramsViewSource.Source = new ObservableCollection<Telegram>(Telegrams);
            TelegramsViewSource.Filter += TelegramsFilter;
        }

        private void RefreshViewSource()
        {
            TelegramsViewSource.Source = new ObservableCollection<Telegram>(Telegrams);
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

            if (result == null)
                e.Accepted = true;
            else
                e.Accepted = result.ToUpper().Contains(FilterText.ToUpper())
                    ? true : false;
        }

        public RelayCommand NewCommand
        {
            get
            {
                return newCommand ??
                    (newCommand = new RelayCommand((o) =>
                    {
                        ActiveTelegram = new Telegram();
                        OpenTlgWnd(ActiveTelegram);
                    }));
            }
        }

        public RelayCommand SaveCommand
        {
            get
            {
                return saveCommand ??
                    (saveCommand = new RelayCommand((wnd) =>
                    {
                        if (wnd is TelegramWnd telegramWindow)
                            telegramWindow.DialogResult = true;
                    }));
            }
        }

        // команда редактирования
        public RelayCommand EditCommand
        {
            get
            {
                return editCommand ??
                  (editCommand = new RelayCommand((o) =>
                  {
                      if (ActiveTelegram is Telegram tlg)
                      {
                          ActiveTelegram = (Telegram)tlg.Clone();
                          OpenTlgWnd(ActiveTelegram);
                      }
                  }));
            }
        }

        // команда добавления
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
                              InsertIntoDb(files);
                          }
                      }
                      else
                      {
                          DragEventArgs e = o as DragEventArgs;
                          files = (string[])e.Data.GetData(DataFormats.FileDrop);
                          InsertIntoDb(files);
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
                        if (ActiveTelegram is Telegram tlg)
                        {
                            _dataService.RemoveTelegram(tlg, RefreshViewSource);
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
                        ImportWnd importWnd = new ImportWnd();

                        if (importWnd.ShowDialog() == true)
                        {

                        }
                        importWnd = null;
                    }));
            }
        }

        private void InsertIntoDb(string[] files)
        {
            if (files == null)
                return;

            foreach (string filePath in files)
            {
                string ext = Path.GetExtension(filePath);
                if ( !(ext.Equals(".doc") || ext.Equals(".docx")))
                    continue;

                using (TextReader reader = new FilterReader(filePath))
                {
                    string docText = reader.ReadToEnd().ToUpper();

                    if (string.IsNullOrWhiteSpace(docText)) throw new Exception("Документ якиий Ви намагаєтесь відкрити не містить данних.");

                    string regexString = string.Concat(TlgRegex.BasePartRegex, TlgRegex.FirstPartRegex);
                    Regex regex = new Regex(regexString, RegexOptions.Multiline, TimeSpan.FromSeconds(3));
                    GroupCollection groups = null;
                    try
                    {
                        groups = regex.Match(docText).Groups;
                    }
                    catch 
                    {
                        throw new Exception("Не вдалося отримати інформацію, перевірте чи не порушена структура документа.");
                    }
                    if (groups.Count < 2)
                    {
                        regexString = string.Concat(TlgRegex.BasePartRegex, TlgRegex.SecondPartRegex);
                        regex = new Regex(regexString, RegexOptions.Multiline);
                        groups = regex.Match(docText).Groups;
                        if (groups.Count < 2)
                        {
                            throw new Exception("Не вдалося отримати інформацію, перевірте чи не порушена структура документа.");
                        }
                    }

                    bool isUrgency = new Regex(TlgRegex.UrgencyRegex).IsMatch(docText);

                    string num = groups[2].Value.Trim().ToString();
                    int number = Int32.Parse((!string.IsNullOrWhiteSpace(num)) ? num : "0");
                    string SenderPosPart1 = groups[6].Value.Trim() + " ";
                    string SenderPosPart2 = groups[8].Value.Trim();

                    int urgency = isUrgency ? 1 : 0;
                    ActiveTelegram = null;
                    ActiveTelegram = new Telegram()
                    {
                        SelfNum = 0,
                        IncNum = number,
                        From = groups[1].Value.Trim(),
                        To = groups[3].Value.Trim(),
                        Text = groups[4].Value.Trim(),
                        SubNum = groups[5].Value.Trim(),
                        Date = groups[7].Value.Trim(),
                        SenderPos = string.Concat(SenderPosPart1, SenderPosPart2),
                        SenderRank = groups[9].Value.Trim(),
                        SenderName = groups[10].Value.Trim(),
                        Executor = groups[11].Value.Trim(),
                        Phone = groups[12].Value.Trim(),
                        HandedBy = groups[13].Value.Trim(),
                        Urgency = urgency,
                        Dispatcher = string.Empty,
                        Time = string.Empty
                    };

                    CheckData(ActiveTelegram);
                    OpenTlgWnd(ActiveTelegram);
                }
            }

        }

        private void OpenTlgWnd(Telegram tlg)
        {
            bool isNew = false;

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

        private bool CheckData(Telegram data)
        {
            if (data.IncNum < 1 ||
                string.IsNullOrWhiteSpace(data.From) ||
                string.IsNullOrWhiteSpace(data.To) ||
                string.IsNullOrWhiteSpace(data.Text) ||
                string.IsNullOrWhiteSpace(data.SubNum) ||
                string.IsNullOrWhiteSpace(data.Date) ||
                string.IsNullOrWhiteSpace(data.SenderPos) ||
                string.IsNullOrWhiteSpace(data.SenderRank) ||
                string.IsNullOrWhiteSpace(data.SenderName) ||
                string.IsNullOrWhiteSpace(data.Executor) ||
                string.IsNullOrWhiteSpace(data.Phone) ||
                string.IsNullOrWhiteSpace(data.HandedBy))
            {
                throw new Exception("Не вдалося отримати деяку інформацію, перевірте структуру документа");
            }
            else
            {
                return true;
            }
        }
    }
}

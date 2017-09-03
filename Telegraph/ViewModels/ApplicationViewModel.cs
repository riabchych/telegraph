using IFilterTextReader;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data.Entity;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Threading;
using Telegraph.ViewModel;

namespace Telegraph
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
        private RelayCommand windowLoaded;
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

        public RelayCommand WindowLoaded
        {
            get
            {
                return windowLoaded ??
                  (windowLoaded = new RelayCommand((lv) =>
                  {

                  }));
            }
        }

        public ApplicationViewModel()   { }

        public ApplicationViewModel(ITelegramDataService dataService)
        {
            FilterType = 3;
            TelegramsViewSource = new CollectionViewSource();
            _dataService = dataService;
            IsBusy = true;
            DbTask = Task.Factory.StartNew(() =>
            {
                _dataService.LoadTelegrams(TelegramsLoaded, TelegramsLoadFiled);
            });
        }

        private void TelegramsLoadFiled(Exception obj)
        {
            throw new NotImplementedException();
        }

        private void TelegramsLoaded(IEnumerable<Telegram> telegrams)
        {
            Application.Current.Dispatcher.Invoke(DispatcherPriority.Normal, (Action)delegate ()
            {
                IsBusy = false;
                Telegrams = Db.Telegrams.Local.ToBindingList();
                TelegramsViewSource.Source = new ObservableCollection<Telegram>(Telegrams);
                TelegramsViewSource.Filter += TelegramsFilter;
            });
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
                e.Accepted = (result.ToUpper().Contains(FilterText.ToUpper()))
                    ? true : false;

        }

        public RelayCommand NewCommand
        {
            get
            {
                return newCommand ??
                    (newCommand = new RelayCommand((o) =>
                    {
                        OpenTlgWnd(new Telegram());
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
                  (editCommand = new RelayCommand((selectedItem) =>
                  {
                      if (selectedItem is Telegram tlg)
                      {
                          tlg = (Telegram)tlg.Clone();
                          OpenTlgWnd(tlg);
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
                          // Create OpenFileDialog 
                          OpenFileDialog dlg = new OpenFileDialog()
                          {
                              // Set filter for file extension and default file extension 
                              DefaultExt = ".doc",
                              Filter = "Word Files (*.doc)|*.doc|Word Files (*.docx)|*.docx|Всі файли (*.*)|*.*",
                              CheckFileExists = true,
                              Multiselect = true
                          };

                          // Display OpenFileDialog by calling ShowDialog method 
                          bool? result = dlg.ShowDialog();


                          // Get the selected file name and display in a TextBox 
                          if (result == true)
                          {
                              // Open document 
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
                    (deleteCommand = new RelayCommand((selectedItem) =>
                    {
                        if (selectedItem is Telegram tlg)
                        {
                            Db.Telegrams.Remove(tlg);
                            Db.SaveChanges();
                            RefreshViewSource();
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
                        if (o == null)
                            return;

                        Telegram tlg;
                        string type = o.GetType().Name;

                        switch (type)
                        {
                            case "Telegram":
                                tlg = o as Telegram;
                                break;
                            case "TelegramWnd":
                                TelegramWnd wnd = o as TelegramWnd;
                                tlg = wnd.Telegram;
                                break;
                            default:
                                tlg = null;
                                break;
                        }
                        if (tlg == null)
                            return;

                        string fileName = GetTempFile("docx");
                        new WordDocument(tlg).CreatePackage(fileName);
                        Process.Start(fileName);
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
            foreach (string filePath in files)
            {

                using (TextReader reader = new FilterReader(filePath))
                {
                    string docText = reader.ReadToEnd().ToUpper();

                    if (string.IsNullOrWhiteSpace(docText)) throw new Exception("Документ якиий Ви намагаєтемь відкрити не містить данних."); ;

                    string regexString = string.Concat(TlgRegex.BasePartRegex, TlgRegex.FirstPartRegex);
                    Regex regex = new Regex(regexString, RegexOptions.Multiline);

                    GroupCollection groups = regex.Match(docText).Groups;
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

                    Telegram tlg = new Telegram()
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

                    CheckData(tlg);
                    OpenTlgWnd(tlg);
                }
            }

        }

        private void OpenTlgWnd(Telegram tlg)
        {
            bool isNew = false;

            if (tlg.Id < 1)
            {
                var culture = new CultureInfo("ru-RU");
                DateTime localDate = DateTime.Now;
                tlg.Time = localDate.ToString(culture);
                isNew = true;
            }

            TelegramWnd TelegramWindow = new TelegramWnd(this,tlg);

            if (TelegramWindow.ShowDialog() == true)
            {
                if (isNew)
                {
                    Db.Telegrams.Add(TelegramWindow.Telegram);
                }
                else
                {
                    tlg = Db.Telegrams.Find(TelegramWindow.Telegram.Id);

                    tlg.From = TelegramWindow.Telegram.From;
                    tlg.To = TelegramWindow.Telegram.To;
                    tlg.SelfNum = TelegramWindow.Telegram.SelfNum;
                    tlg.IncNum = TelegramWindow.Telegram.IncNum;
                    tlg.Text = TelegramWindow.Telegram.Text;
                    tlg.SubNum = TelegramWindow.Telegram.SubNum;
                    tlg.Date = TelegramWindow.Telegram.Date;
                    tlg.SenderPos = TelegramWindow.Telegram.SenderPos;
                    tlg.SenderRank = TelegramWindow.Telegram.SenderRank;
                    tlg.SenderName = TelegramWindow.Telegram.SenderName;
                    tlg.Executor = TelegramWindow.Telegram.Executor;
                    tlg.Phone = TelegramWindow.Telegram.Phone;
                    tlg.HandedBy = TelegramWindow.Telegram.HandedBy;
                    tlg.Urgency = TelegramWindow.Telegram.Urgency;
                    tlg.Dispatcher = TelegramWindow.Telegram.Dispatcher;
                    tlg.Time = TelegramWindow.Telegram.Time;

                    Db.Entry(tlg).State = EntityState.Modified;
                }
                Db.SaveChanges();
                RefreshViewSource();
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

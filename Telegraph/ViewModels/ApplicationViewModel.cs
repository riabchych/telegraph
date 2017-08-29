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

namespace Telegraph
{
    struct TlgRegex
    {
        private const string firstPartRegex = @"^ВИК[\s\.]+([А-ЯЁЇІЄҐ\1]+){1}[\s]+([0-9\-\:]+){1}[\s]*$[\s\-]+^ПРД[\s\.]+([А-ЯЁЇІЄҐ\1]+){1}";
        private const string secondPartRegex = @"^ПРД[\s\.]+([[А-ЯЁЇІЄҐ\1]+)[\s]*$[\s\r\n]+^ВИК[\s\.]+([[А-ЯЁЇІЄҐ\1]+){1}[\s]+([0-9\-\:]+)[\s]?";
        private const string basePartRegex = @"[А-ЯЁЇІЄҐ\s\d\W]+^([А-ЯЁЇІЄҐ\s\d]*).НР[\s\.]+([0-9]*)?[\s]?$[\s]+^([А-ЯЁЇІЄҐ\d\/\s]+)[\s\n\r]+^([А-ЯЁЇІЄҐ\w\W\s]+)^[\s]*НР[\.]?[\s]*([\d\/\s\-]+[\s]*[\d]+)[\s]+([А-ЯЁЇІЄҐ\s\.\-0-9]+)[\r\n]^([\d]+[\s\.\\\/]+[\d]+[\s\.\\\/]+[\d]+)[\s]+([А-ЯЁЇІЄҐ\s\-\1]+)$[\s]+([А-ЯЁЇІЄҐ\-\1]+)?[\s]+([А-ЯЁЇІЄҐ\s\1\.]+)$[\r\n]?[\s\-]+";
        private const string urgencyRegex = @"[\s\d]{10,}....(ТЕРМІНОВ.)";

        public static string FirstPartRegex => firstPartRegex;

        public static string SecondPartRegex => secondPartRegex;

        public static string BasePartRegex => basePartRegex;

        public static string UrgencyRegex => urgencyRegex;
    }

    public class ApplicationViewModel : PropertyChangedNotification
    {
        private ApplicationContext db;
        private Task dbTask;
        private static ApplicationViewModel applicationViewModel;
        private TelegramWnd telegramWindow;

        RelayCommand newCommand;
        RelayCommand addCommand;
        RelayCommand editCommand;
        RelayCommand deleteCommand;
        RelayCommand sendToWord;
        RelayCommand windowLoaded;

        public static ApplicationViewModel SharedViewModel()
        {
            return applicationViewModel ?? (applicationViewModel = new ApplicationViewModel());
        }

        public IEnumerable<Telegram> Telegrams
        {
            get { return GetValue(() => Telegrams); }
            set { SetValue(() => Telegrams, value); }
        }

        ObservableCollection<Telegram> TelegramsCollection
        {
            get { return GetValue(() => TelegramsCollection); }
            set { SetValue(() => TelegramsCollection, value); }
        }

        public CollectionViewSource TelegramsViewSource
        {
            get { return GetValue(() => TelegramsViewSource); }
            set { SetValue(() => TelegramsViewSource, value); }
        }

        public ICollectionView SourceCollection
        {
            get { return TelegramsViewSource.View; }
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
            set
            {
                SetValue(() => FilterType, value);
            }
        }

        public bool IsBusy
        {
            get { return GetValue(() => IsBusy); }
            set { SetValue(() => IsBusy, value); }
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
                      ListView listview = lv as ListView;
                      IsBusy = true;
                      FilterType = 3;
                      Db = new ApplicationContext();
                      TelegramsViewSource = new CollectionViewSource();
                      DbTask = Task.Factory.StartNew(async () =>
                      {
                          Db.Telegrams.Load();

                          await DbTask;

                          Application.Current.Dispatcher.Invoke(DispatcherPriority.Normal, (Action)delegate ()
                          {
                              IsBusy = false;
                              Telegrams = Db.Telegrams.Local.ToBindingList();
                              TelegramsViewSource.Source = new ObservableCollection<Telegram>(Telegrams);
                              TelegramsViewSource.Filter += TelegramsFilter;
                          });

                      });
                  }));
            }
        }
        public void RefreshViewSource()
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

            Telegram tlg = e.Item as Telegram;
            string result = null;

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

        // команда редактирования
        public RelayCommand EditCommand
        {
            get
            {
                return editCommand ??
                  (editCommand = new RelayCommand((selectedItem) =>
                  {
                      if (selectedItem == null) return;
                      // получаем выделенный объект
                      Telegram tlg = selectedItem as Telegram;
                      tlg = (Telegram)tlg.Clone();
                      OpenTlgWnd(tlg);

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
                        if (selectedItem == null) return;
                        Telegram tlg = selectedItem as Telegram;
                        Db.Telegrams.Remove(tlg);
                        Db.SaveChanges();
                        RefreshViewSource();
                    }));
            }
        }

        public RelayCommand SendToWord
        {
            get
            {
                return sendToWord ??
                    (sendToWord = new RelayCommand((selectedItem) =>
                    {
                        Telegram tlg;
                        if (TelegramWindow == null)
                            if (selectedItem == null)
                                return;
                            else
                                tlg = selectedItem as Telegram;
                        else
                            tlg = TelegramWindow.Telegram;

                        string fileName = GetTempFile("docx");
                        new WordDocument(tlg).CreatePackage(fileName);
                        Process.Start(fileName);
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

            TelegramWindow = new TelegramWnd(tlg);
            
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

                TelegramWindow.Close();
                TelegramWindow = null;
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

        private string GetTempFile(string fileExtension)
        {
            string temp = System.IO.Path.GetTempPath();
            string res = string.Empty;
            while (true)
            {
                res = string.Format("{0}.{1}", Guid.NewGuid().ToString(), fileExtension);
                res = System.IO.Path.Combine(temp, res);
                if (!System.IO.File.Exists(res))
                {
                    try
                    {
                        System.IO.FileStream s = System.IO.File.Create(res);
                        s.Close();
                        break;
                    }
                    catch (Exception)
                    {

                    }
                }
            }
            return res;
        }

        public ApplicationContext Db { get => db; set => db = value; }
        public Task DbTask { get => dbTask; set => dbTask = value; }
        public TelegramWnd TelegramWindow { get => telegramWindow; set => telegramWindow = value; }
    }
}

using IFilterTextReader;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Globalization;
using System.IO;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;

namespace Telegraph
{
    struct TlgRegex
    {
        private const string firstPartRegex = @"^ВИК[\s\.]+([ЙЦУКЕНГШЩЗХЇЄЖДЛОРПАВІФЯЧСМИТЬБЮ0-9]+){1}[\s]+([0-9\-\:]+){1}[\s]*$[\n\r\s\-]+^ПРД[\s\.]+([ЙЦУКЕНГШЩЗХЇЄЖДЛОРПАВІФЯЧСМИТЬБЮ0-9]+){1}";
        private const string secondPartRegex = @"^ПРД[\s\.]+([ЙЦУКЕНГШЩЗХЇЄЖДЛОРПАВІФЯЧСМИТЬБЮ0-9]+)[\s]*$[\s\r\n]+^ВИК[\s\.]+([ЙЦУКЕНГШЩЗХЇЄЖДЛОРПАВІФЯЧСМИТЬБЮ0-9]+){1}[\s]+([0-9\-\:]+)[\s]?";
        private const string basePartRegex = @"^[\.]?[\w\W]*.НР[\s\.]*([0-9]*){1}[\s]*$[\n\r]+^[ЙЦУКЕНГШЩЗХЇЄЖДЛОРПАВІФЯЧСМИТЬБЮ\s\/]+([ЙЦУКЕНГШЩЗХЇЄЖДЛОРПАВІФЯЧСМИТЬБЮ\d\s]+)$[\n\r]+^([ЙЦУКЕНГШЩЗХЇЄЖДЛОРПАВІФЯЧСМИТЬБЮ\w\d\W\n\r]+)^[\s]*НР\.[\s]*([\d\/\s\-]+[\s]*[\d]+)[\s]+([ЙЦУКЕНГШЩЗХЇЄЖДЛОРПАВІФЯЧСМИТЬБЮ\s\.\-0-9]+)^([\d]+[\s\.\\]+[\d]+[\s\.\\]+[\d]+)[\s]*([ЙЦУКЕНГШЩЗХЇЄЖДЛОРПАВІФЯЧСМИТЬБЮ\s\-0-9]+)$[\n\r\s]+([ЙЦУКЕНГШЩЗХЇЄЖДЛОРПАВІФЯЧСМИТЬБЮ\-0-9]+){1}[\s]+([ЙЦУКЕНГШЩЗХЇЄЖДЛОРПАВІФЯЧСМИТЬБЮ\s\-0-9\.]+)[\s][\r\n][\n\r\s\-]+";
        private const string urgencyRegex = @"[\s\t\d]{10,}....(ТЕРМІНОВ.)";

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

        RelayCommand newCommand;
        RelayCommand addCommand;
        RelayCommand editCommand;
        RelayCommand deleteCommand;
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

        public int EditSelfNum
        {
            get { return GetValue(() => EditSelfNum); }
            set { SetValue(() => EditSelfNum, value); }
        }

        public int EditIncNum
        {
            get { return GetValue(() => EditIncNum); }
            set { SetValue(() => EditIncNum, value); }
        }

        public bool IsBusy
        {
            get { return GetValue(() => IsBusy); }
            set { SetValue(() => IsBusy, value); }
        }

        public RelayCommand WindowLoaded
        {
            get
            {
                return windowLoaded ??
                  (windowLoaded = new RelayCommand((o) =>
                  {
                      IsBusy = true;
                      Db = new ApplicationContext();
                      DbTask = Task.Factory.StartNew(async () =>
                      {

                          Db.Telegrams.Load();

                          await DbTask;

                          Telegrams = Db.Telegrams.Local.ToBindingList();
                          IsBusy = false;

                      });
                  }));
            }
        }

        public RelayCommand NewCommand
        {
            get
            {
                return newCommand ??
                    (newCommand = new RelayCommand((o) =>
                    {
                        EditIncNum = 0;
                        EditSelfNum = 0;
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
                      EditIncNum = tlg.IncNum;
                      EditSelfNum = tlg.SelfNum;
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
                              EditIncNum = 0;
                              EditSelfNum = 0;
                              InsertIntoDb(files);
                          }
                      }
                      else
                      {
                          DragEventArgs e = o as DragEventArgs;
                          files = (string[])e.Data.GetData(DataFormats.FileDrop);
                          EditIncNum = 0;
                          EditSelfNum = 0;
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
                        db.Telegrams.Remove(tlg);
                        db.SaveChanges();
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

                    string num = groups[1].Value.Trim().ToString();
                    int number = Int32.Parse((!string.IsNullOrWhiteSpace(num)) ? num : "0");
                    string SenderPosPart1 = groups[5].Value.Trim();
                    string SenderPosPart2 = groups[7].Value.Trim();

                    int urgency = isUrgency ? 1 : 0;

                    Telegram tlg = new Telegram()
                    {
                        SelfNum = 0,
                        IncNum = number,
                        To = groups[2].Value.Trim(),
                        Text = groups[3].Value.Trim(),
                        SubNum = groups[4].Value.Trim(),
                        Date = groups[6].Value.Trim(),
                        SenderPos = string.Concat(SenderPosPart1, SenderPosPart2),
                        SenderRank = groups[8].Value.Trim(),
                        SenderName = groups[9].Value.Trim(),
                        Executor = groups[10].Value.Trim(),
                        Phone = groups[11].Value.Trim(),
                        HandedBy = groups[12].Value.Trim(),
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

            TelegramWnd telegramWindow = new TelegramWnd(tlg);

            if (telegramWindow.ShowDialog() == true)
            {
                if (isNew)
                {
                    db.Telegrams.Add(telegramWindow.Telegram);
                }
                else
                {
                    tlg = db.Telegrams.Find(telegramWindow.Telegram.Id);

                    tlg.SelfNum = telegramWindow.Telegram.SelfNum;
                    tlg.IncNum = telegramWindow.Telegram.IncNum;
                    tlg.To = telegramWindow.Telegram.To;
                    tlg.Text = telegramWindow.Telegram.Text;
                    tlg.SubNum = telegramWindow.Telegram.SubNum;
                    tlg.Date = telegramWindow.Telegram.Date;
                    tlg.SenderPos = telegramWindow.Telegram.SenderPos;
                    tlg.SenderRank = telegramWindow.Telegram.SenderRank;
                    tlg.SenderName = telegramWindow.Telegram.SenderName;
                    tlg.Executor = telegramWindow.Telegram.Executor;
                    tlg.Phone = telegramWindow.Telegram.Phone;
                    tlg.HandedBy = telegramWindow.Telegram.HandedBy;
                    tlg.Urgency = telegramWindow.Telegram.Urgency;
                    tlg.Dispatcher = telegramWindow.Telegram.Dispatcher;
                    tlg.Time = telegramWindow.Telegram.Time;

                    db.Entry(tlg).State = EntityState.Modified;
                }
                db.SaveChanges();
            }


        }

        private bool CheckData(Telegram data)
        {
            if (data.IncNum < 1 ||
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

        public ApplicationContext Db { get => db; set => db = value; }
        public Task DbTask { get => dbTask; set => dbTask = value; }
    }
}

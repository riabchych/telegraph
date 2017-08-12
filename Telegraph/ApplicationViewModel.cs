using IFilterTextReader;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data.Entity;
using System.IO;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;

namespace Telegraph
{
    struct TlgRegex
    {
        private const string number = @".З[\w\W]*.НР[\W]*([0-9]*)[\W]*[\n]";
        private const string to = @"(КОМАН[ЙЦУКЕНГШЩЗХЇЄЖДЛОРПАВІФЯЧСМИТЬБЮ]+[\w\d\W].*[\n\r])";
        private const string text = @"КОМАН[ЙЦУКЕНГШЩЗХЇЄЖДЛОРПАВІФЯЧСМИТЬБЮ]+[\w\d\W].*[\n\r].?([ЙЦУКЕНГШЩЗХЇЄЖДЛОРПАВІФЯЧСМИТЬБЮ\w\d\W\n\r]*)НР";
        private const string urgency = @"[\s\t\d]{10,}....(ТЕРМІНОВ.)";
        private const string all = @"^.З[\w\W]*.НР[\s\.]*([0-9]*){1}[\s]*$[\n\r]+^[ЙЦУКЕНГШЩЗХЇЄЖДЛОРПАВІФЯЧСМИТЬБЮ\s\/]+([ЙЦУКЕНГШЩЗХЇЄЖДЛОРПАВІФЯЧСМИТЬБЮ\d\s]+)$[\n\r]+^([ЙЦУКЕНГШЩЗХЇЄЖДЛОРПАВІФЯЧСМИТЬБЮ\w\d\W\n\r]+)^НР\.([\d\/]+[\-]{1}[\s]*[\d]+)[\s]+([ЙЦУКЕНГШЩЗХЇЄЖДЛОРПАВІФЯЧСМИТЬБЮ\s\.\-0-9]+)^([\d\.\/\\]+){1}[\s]+([ЙЦУКЕНГШЩЗХЇЄЖДЛОРПАВІФЯЧСМИТЬБЮ\s\-0-9]+)$[\n\r\s]+([ЙЦУКЕНГШЩЗХЇЄЖДЛОРПАВІФЯЧСМИТЬБЮ\-0-9]+){1}[\s]+([ЙЦУКЕНГШЩЗХЇЄЖДЛОРПАВІФЯЧСМИТЬБЮ\s\-0-9\.]+)[\s][\r\n][\n\r\s\-]+^ВИК[\s\.]+([ЙЦУКЕНГШЩЗХЇЄЖДЛОРПАВІФЯЧСМИТЬБЮ0-9]+){1}[\s]+([0-9\-\:]+){1}[\s]*$[\n\r\s\-]+^ПРД[\s\.]+([ЙЦУКЕНГШЩЗХЇЄЖДЛОРПАВІФЯЧСМИТЬБЮ0-9]+){1}";

        public static string Number => number;
        public static string To => to;
        public static string Text => text;
        public static string Urgency => urgency;
        public static string All => all;
    }

    public class ApplicationViewModel : INotifyPropertyChanged
    {
        private ApplicationContext db;
        private IEnumerable<Telegram> telegrams;
        private IEnumerable<User> Users;
        private Task dbTask;

        RelayCommand newCommand;
        RelayCommand addCommand;
        RelayCommand editCommand;
        RelayCommand deleteCommand;
        private Telegram telegram;

        public ApplicationViewModel()
        {
            InitApplicationViewModel();
        }

        public ApplicationViewModel(SynchronizationContext context, SendOrPostCallback callback)
        {
            InitApplicationViewModel(context, callback);
        }

        private void InitApplicationViewModel(SynchronizationContext context = null, SendOrPostCallback callback = null)
        {
            Db = new ApplicationContext();

            DbTask = Task.Factory.StartNew(async () =>
            {
                Db.Telegrams.Load();

                if (await Task.WhenAny(DbTask, Task.Delay(1000 * 60)) == DbTask)
                {
                    Telegrams = Db.Telegrams.Local.ToBindingList();
                    if (context != null && callback != null)
                    {
                        context.Send(callback, null);
                    }
                }
            });
        }

        public RelayCommand NewCommand
        {
            get
            {
                return newCommand ??
                    (newCommand = new RelayCommand((o) =>
                    {
                        TelegramWnd telegramWnd = new TelegramWnd(new Telegram());
                        if (telegramWnd.ShowDialog() == true)
                        {
                            Telegram tlg = telegramWnd.Telegram;
                            db.Telegrams.Add(tlg);
                            db.SaveChanges();
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
                          OpenFileDialog dlg = new OpenFileDialog();



                          // Set filter for file extension and default file extension 
                          dlg.DefaultExt = ".doc";
                          dlg.Filter = "Word Files (*.doc)|*.doc|Word Files (*.docx)|*.docx|Всі файли (*.*)|*.*";
                          dlg.CheckFileExists = true;
                          dlg.Multiselect = true;

                          // Display OpenFileDialog by calling ShowDialog method 
                          Nullable<bool> result = dlg.ShowDialog();


                          // Get the selected file name and display in a TextBox 
                          if (result == true)
                          {
                              // Open document 
                              files = dlg.FileNames;
                          }
                      }
                      else
                      {
                          DragEventArgs e = o as DragEventArgs;
                          files = (string[])e.Data.GetData(DataFormats.FileDrop);
                      }

                      InsertIntoDb(files);


                  }));
            }
        }

        private void InsertIntoDb(string[] files)
        {
            foreach (string filePath in files)
            {
                using (TextReader reader = new FilterReader(filePath))
                {
                    string docText = reader.ReadToEnd();

                    if (string.IsNullOrWhiteSpace(docText)) continue;

                    GroupCollection groups = new Regex(TlgRegex.All, RegexOptions.Multiline).Match(docText).Groups;
                    bool isUrgency = new Regex(TlgRegex.Urgency).IsMatch(docText);

                    string num = groups[1].Value.Trim().ToString();
                    int number = Int32.Parse((!string.IsNullOrWhiteSpace(num)) ? num : "0");
                    string SenderPosPart1 = groups[5].Value.Trim();
                    string SenderPosPart2 = groups[7].Value.Trim();

                    int urgency = isUrgency ? 1 : 0;

                    try
                    {
                        Telegram tlg = new Telegram()
                        {
                            Number = number,
                            To = groups[2].Value.Trim(),
                            Text = groups[3].Value.Trim(),
                            Subnum = groups[4].Value.Trim(),
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

                        //TelegramViewModel vm = new TelegramViewModel();
                        TelegramWnd telegramWnd = new TelegramWnd(tlg);
                        if (telegramWnd.ShowDialog() == true)
                        {
                            telegramWnd.Telegram = tlg;
                            db.Telegrams.Add(tlg);
                        }


                    }
                    catch (Exception err)
                    {
                        MessageBox.Show(err.Message);
                    }

                    /*using (StreamWriter sw = new StreamWriter(wordDoc.MainDocumentPart.GetStream(FileMode.Create)))
                    {
                        sw.Write(docText);
                    }*/
                }
            }
            db.SaveChanges();
        }

        private void OpenTlgWnd(Telegram tlg)
        {
            TelegramWnd telegramWindow = new TelegramWnd(tlg);

            if (telegramWindow.ShowDialog() == true)
            {
                // получаем измененный объект
                telegramWindow.Telegram = tlg;

                db.Telegrams.Add(tlg);
                telegram = db.Telegrams.Find(telegramWindow.Telegram.id);
                if (telegram != null)
                {
                    telegram.Number = telegramWindow.Telegram.Number;
                    telegram.To = telegramWindow.Telegram.To;
                    telegram.Text = telegramWindow.Telegram.Text;
                    telegram.Subnum = telegramWindow.Telegram.Subnum;
                    telegram.Date = telegramWindow.Telegram.Date;
                    telegram.Urgency = telegramWindow.Telegram.Urgency;
                    telegram.Dispatcher = telegramWindow.Telegram.Dispatcher;

                    db.Entry(telegram).State = EntityState.Modified;
                    db.SaveChanges();
                }
            }
        }

        private bool CheckData(Telegram data)
        {
            if (data.Number < 1 || string.IsNullOrWhiteSpace(data.To) || string.IsNullOrWhiteSpace(data.Text) ||
                string.IsNullOrWhiteSpace(data.Subnum) || string.IsNullOrWhiteSpace(data.Date))
            {
                return false; //throw new Exception("Не удалось извлечь некоторую информацию.");
            }
            else
            {
                return true;
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
                      Telegram telegram = selectedItem as Telegram;

                      Telegram tlg = new Telegram()
                      {
                          id = telegram.id,
                          Number = telegram.Number,
                          To = telegram.To,
                          Text = telegram.Text,
                          Subnum = telegram.Subnum,
                          Date = telegram.Date,
                          Urgency = telegram.Urgency,
                          Dispatcher = telegram.Dispatcher,
                      };

                      TelegramWnd telegramWindow = new TelegramWnd(tlg);


                      if (telegramWindow.ShowDialog() == true)
                      {
                          // получаем измененный объект
                          telegram = db.Telegrams.Find(telegramWindow.Telegram.id);
                          if (telegram != null)
                          {
                              telegram.Number = telegramWindow.Telegram.Number;
                              telegram.To = telegramWindow.Telegram.To;
                              telegram.Text = telegramWindow.Telegram.Text;
                              telegram.Subnum = telegramWindow.Telegram.Subnum;
                              telegram.Date = telegramWindow.Telegram.Date;
                              telegram.Urgency = telegramWindow.Telegram.Urgency;
                              telegram.Dispatcher = telegramWindow.Telegram.Dispatcher;

                              db.Entry(telegram).State = EntityState.Modified;
                              db.SaveChanges();
                          }
                      }
                  }));
            }
        }

        public ApplicationContext Db { get => db; set => db = value; }
        public IEnumerable<Telegram> Telegrams { get => telegrams; set => telegrams = value; }
        public Task DbTask { get => dbTask; set => dbTask = value; }
        public RelayCommand DeleteCommand { get => deleteCommand; set => deleteCommand = value; }
        public IEnumerable<User> Users1 { get => Users; set => Users = value; }

        public event PropertyChangedEventHandler PropertyChanged;
        public void OnPropertyChanged([CallerMemberName]string prop = "")
        {
            if (PropertyChanged != null)
                PropertyChanged(this, new PropertyChangedEventArgs(prop));
        }
    }
}

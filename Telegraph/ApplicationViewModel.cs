using IFilterTextReader;
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
        private const string all = @".З[\w\W]*.НР[\W]*([0-9]*)[\W\n\r]*КОМАН[ЙЦУКЕНГШЩЗХЇЄЖДЛОРПАВІФЯЧСМИТЬБЮ\s\/]+([\d\s].*)[\n\r].?([ЙЦУКЕНГШЩЗХЇЄЖДЛОРПАВІФЯЧСМИТЬБЮ\w\d\W\n\r]*)НР\.([\d\/\-]*)[[ЙЦУКЕНГШЩЗХЇЄЖДЛОРПАВІФЯЧСМИТЬБЮ\s\.]*]?([\d\.]+)";

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
        private Task dbTask;

        RelayCommand addCommand;
        RelayCommand editCommand;
        RelayCommand deleteCommand;

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
        // команда добавления
        public RelayCommand AddCommand
        {
            get
            {
                return addCommand ??
                  (addCommand = new RelayCommand((o) =>
                  {
                      if (o == null) return;
                      // получаем выделенный объект
                      DragEventArgs e = o as DragEventArgs;

                      string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                      TextReader reader;
                      bool isUrgency;

                      foreach (string filePath in files)
                      {
                          reader = new FilterReader(filePath);

                          using (reader)
                          {
                              string docText = null;
                              docText = reader.ReadToEnd();

                              GroupCollection groups = new Regex(TlgRegex.All).Match(docText).Groups;
                              isUrgency = new Regex(TlgRegex.Urgency).IsMatch(docText);

                              string num = groups[1].Value.Trim().ToString();

                              int number = Int32.Parse((num != "") ? num : "0");
                              string to = groups[2].Value.Trim();
                              string text = groups[3].Value.Trim();
                              string subNum = groups[4].Value.Trim();
                              string date = groups[5].Value.Trim();
                              int urgency = isUrgency ? 1 : 0;

                              try
                              {
                                  if (number < 1)
                                      throw new Exception("Не удалось извлечь некоторую информацию.");
                                  if (to == "")
                                      throw new Exception("Не удалось извлечь некоторую информацию.");
                                  if (text == "")
                                      throw new Exception("Не удалось извлечь некоторую информацию.");
                                  if (subNum == "")
                                      throw new Exception("Не удалось извлечь некоторую информацию.");
                                  if (date == "")
                                      throw new Exception("Не удалось извлечь некоторую информацию.");

                                  Telegram tlg = new Telegram()
                                  {
                                      Number = number,
                                      To = to,
                                      Text = text,
                                      Subnum = subNum,
                                      Date = date,
                                      Urgency = urgency,
                                  };

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

        public event PropertyChangedEventHandler PropertyChanged;
        public void OnPropertyChanged([CallerMemberName]string prop = "")
        {
            if (PropertyChanged != null)
                PropertyChanged(this, new PropertyChangedEventArgs(prop));
        }
    }
}

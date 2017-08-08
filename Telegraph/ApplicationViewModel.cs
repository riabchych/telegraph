using System.Collections.Generic;
using System.ComponentModel;
using System.Data.Entity;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Threading.Tasks;

namespace Telegraph
{
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
        /*public RelayCommand AddCommand
        {
            get
            {
                return addCommand ??
                  (addCommand = new RelayCommand((o) =>
                  {
                      TelegramWindow telegramWindow = new TelegramWindow(new Telegram());
                      if (telegramWindow.ShowDialog() == true)
                      {
                          Telegram telegram = telegramWindow.Telegram;
                          db.Telegrams.Add(telegram);
                          db.SaveChanges();
                      }
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

                      Telegram vm = new Telegram()
                      {
                          Id = telegram.Id,
                          Company = telegram.Company,
                          Price = telegram.Price,
                          Title = telegram.Title
                      };
                      TelegramWindow telegramWindow = new TelegramWindow(vm);


                      if (telegramWindow.ShowDialog() == true)
                      {
                          // получаем измененный объект
                          telegram = db.Telegrams.Find(telegramWindow.Telegram.Id);
                          if (telegram != null)
                          {
                              telegram.Company = telegramWindow.Telegram.Company;
                              telegram.Title = telegramWindow.Telegram.Title;
                              telegram.Price = telegramWindow.Telegram.Price;
                              db.Entry(telegram).State = EntityState.Modified;
                              db.SaveChanges();
                          }
                      }
                  }));
            }
        }*/
        // команда удаления
        public RelayCommand DeleteCommand
        {
            get
            {
                return deleteCommand ??
                  (deleteCommand = new RelayCommand((selectedItem) =>
                  {
                      if (selectedItem == null) return;
                      // получаем выделенный объект
                      Telegram telegram = selectedItem as Telegram;
                      Db.Telegrams.Remove(telegram);
                      Db.SaveChanges();
                  }));
            }
        }

        public ApplicationContext Db { get => db; set => db = value; }

        public IEnumerable<Telegram> Telegrams
        {
            get { return telegrams; }
            set
            {
                telegrams = value;
                OnPropertyChanged("Telegrams");
            }
        }

        public Task DbTask { get => dbTask; set => dbTask = value; }

        public event PropertyChangedEventHandler PropertyChanged;
        public void OnPropertyChanged([CallerMemberName]string prop = "")
        {
            if (PropertyChanged != null)
                PropertyChanged(this, new PropertyChangedEventArgs(prop));
        }
    }
}

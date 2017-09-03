using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Globalization;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Threading;

namespace Telegraph
{
    public interface ITelegramDataService
    {
        void LoadTelegrams(
            Action<IEnumerable<Telegram>> success = null,
            Action<Exception> fail = null);

        void AddTelegram(
            Telegram tlg,
            Action success = null,
            Action<Exception> fail = null);

        void EditTelegram(
            int id,
            Telegram tlg,
            Action success = null,
            Action<Exception> fail = null);

        void RemoveTelegram(
            Telegram tlg,
            Action success = null,
            Action<Exception> fail = null);
    }

    public class TelegramDataService : ITelegramDataService
    {
        private ApplicationContext Db;

        public TelegramDataService()
        {
            Db = new ApplicationContext();
        }

        public void AddTelegram(Telegram tlg, Action success = null, Action<Exception> fail = null)
        {
            using (var transaction = Db.Database.BeginTransaction())
            {
                try
                {
                    Db.Telegrams.Add(tlg);
                    Db.SaveChanges();
                    transaction.Commit();
                    success?.Invoke();
                }
                catch (Exception ex)
                {
                    transaction.Rollback();
                    fail?.Invoke(ex);
                }
            }
        }

        public void EditTelegram(int id, Telegram tlg, Action success = null, Action<Exception> fail = null)
        {

            using (var transaction = Db.Database.BeginTransaction())
            {
                try
                {
                    Telegram t = Db.Telegrams.Find(id);

                    t.From = tlg.From;
                    t.To = tlg.To;
                    t.SelfNum = tlg.SelfNum;
                    t.IncNum = tlg.IncNum;
                    t.Text = tlg.Text;
                    t.SubNum = tlg.SubNum;
                    t.Date = tlg.Date;
                    t.SenderPos = tlg.SenderPos;
                    t.SenderRank = tlg.SenderRank;
                    t.SenderName = tlg.SenderName;
                    t.Executor = tlg.Executor;
                    t.Phone = tlg.Phone;
                    t.HandedBy = tlg.HandedBy;
                    t.Urgency = tlg.Urgency;
                    t.Dispatcher = tlg.Dispatcher;
                    t.Time = tlg.Time;

                    Db.Entry(t).State = EntityState.Modified;
                    Db.SaveChanges();
                    transaction.Commit();
                    success?.Invoke();
                }
                catch (Exception ex)
                {
                    transaction.Rollback();
                    fail?.Invoke(ex);
                }
            }
        }

        public void LoadTelegrams(
            Action<IEnumerable<Telegram>> success,
            Action<Exception> fail)
        {
            try
            {
                new Task(() =>
                {
                    try
                    {
                        Db.Telegrams.Load();
                        Application.Current.Dispatcher.Invoke(DispatcherPriority.Normal, (Action)delegate ()
                        {
                            success?.Invoke(Db.Telegrams.Local.ToBindingList());
                        });
                    }
                    catch (Exception ex)
                    {
                        fail?.Invoke(ex);
                    }

                }).Start();
            }
            catch (Exception e)
            {
                fail?.Invoke(e);
            }
        }

        public void RemoveTelegram(Telegram tlg, Action success = null, Action<Exception> fail = null)
        {

            using (var transaction = Db.Database.BeginTransaction())
            {
                try
                {
                    Db.Telegrams.Remove(tlg);
                    Db.SaveChanges();
                    transaction.Commit();
                    success?.Invoke();
                }
                catch (Exception ex)
                {
                    transaction.Rollback();
                    fail?.Invoke(ex);
                }
            }
        }
    }

    public class TelegramDisignDataService : ITelegramDataService
    {
        private ApplicationContext Db;


        public TelegramDisignDataService()
        {
            Db = new ApplicationContext();
        }

        public void AddTelegram(Telegram tlg, Action success = null, Action<Exception> fail = null)
        {
            throw new NotImplementedException();
        }

        public void EditTelegram(int id, Telegram tlg, Action success = null, Action<Exception> fail = null)
        {
            throw new NotImplementedException();
        }

        public void LoadTelegrams(
            Action<IEnumerable<Telegram>> success,
            Action<Exception> fail)
        {
            success?.Invoke(CreateDisignTelegrams());
        }

        public void RemoveTelegram(Telegram tlg, Action  success = null, Action<Exception> fail = null)
        {
            throw new NotImplementedException();
        }

        private IEnumerable<Telegram> CreateDisignTelegrams()
        {
            var telegrams = new List<Telegram>();
            var tlgCounter = 1;
            var culture = new CultureInfo("ru-RU");
            DateTime localDate = DateTime.Now;
            var n = new Random(100);
            while (tlgCounter <= 10)
            {
                var tlg = new Telegram()
                {
                    SelfNum = tlgCounter++,
                    IncNum = n.Next(100, 200),
                    From = "1З КИЕВА   П1ВН1ЧНЕ КИЇВСЬКЕ ТЕРИТОРІАЛЬНЕ УПРАВЛІННЯ",
                    To = "КОМАНДИРОВІ  В/Ч 306" + n.Next(1, 100),
                    Text = " НА ВАШ ВИХІДНИЙ ВІД 31.09.2017 № 1/9-1072" + tlgCounter,
                    SubNum = tlgCounter.ToString(),
                    Date = localDate.ToString(culture),
                    SenderPos = "НАЧАЛЬНИК ПІВНІЧНОГО КИЇВСЬКОГО ТЕРИТОРІАЛЬНОГО УПРАВЛІННЯ НГ УКРАЇНИ",
                    SenderRank = "ПОЛКОВНИК",
                    SenderName = "Ю.А. МАКАРЧУК",
                    Executor = "ШИРМАНОВ",
                    Phone = "30-17",
                    HandedBy = "ТАРАСЮК",
                    Urgency = n.Next(0, 1),
                    Dispatcher = "РЯБЧИЧ",
                    Time = localDate.ToString(culture)
                };
                telegrams.Add(tlg);
            }
            return telegrams;
        }
    }
}

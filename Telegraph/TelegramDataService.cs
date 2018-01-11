using Microsoft.Practices.ServiceLocation;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data.Entity;
using Telegraph.LogModule.Loggers;

namespace Telegraph
{
    public class TelegramDataService : ITelegramDataService
    {
        private readonly ApplicationContext Db;
        private readonly ILogger logger;

        public TelegramDataService()
        {
            logger = ServiceLocator.Current.GetInstance<ILogger>();
            Db = new ApplicationContext();
        }

        public bool AddTelegram(Telegram tlg)
        {
            using (var transaction = Db.Database.BeginTransaction())
            {
                try
                {
                    Db.Telegrams.Add(tlg);
                    Db.SaveChanges();
                    transaction.Commit();
                    logger.Info("Телеграма успішно додана в базу.");
                    return true;
                }
                catch (Exception ex)
                {
                    transaction.Rollback();
                    logger.Debug("Виникла помилка при доданні телеграми в базу.");
                    logger.Error(ex.Message);
                    return false;
                }
            }
        }

        public bool EditTelegram(int id, Telegram tlg)
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
                    logger.Info("Телеграма успішно збережена.");
                    return true;
                }
                catch (Exception ex)
                {
                    transaction.Rollback();
                    logger.Info("Виникла помилка при редагуванні телеграми.");
                    logger.Error(ex.Message);
                    return false;
                }
            }
        }

        public BindingList<Telegram> LoadTelegrams()
        {
            try
            {
                logger.Debug("Відбувається завантаження телеграм з бази.");
                Db.Telegrams.Load();
                logger.Info("Телеграми успішно завантажені.");
                return Db.Telegrams.Local.ToBindingList();
            }
            catch (Exception ex)
            {
                logger.Info("Виникла помилка при завантаженні телеграм.");
                logger.Error(ex.Message);
                return null;
            }
        }

        public bool RemoveTelegram(object tlg)
        {
            using (DbContextTransaction transaction = Db.Database.BeginTransaction())
            {
                try
                {
                    foreach (Telegram t in tlg as IList<Telegram>)
                    {
                        Db.Telegrams.Remove(t);
                    }
                    Db.SaveChanges();
                    transaction.Commit();
                    logger.Info("Телеграми успішно видалені.");
                    return true;
                }
                catch (Exception ex)
                {
                    transaction.Rollback();
                    logger.Info("Виникла помилка при видаленні телеграми.");
                    logger.Error(ex.Message);
                    return false;
                }
            }
        }
    }
}

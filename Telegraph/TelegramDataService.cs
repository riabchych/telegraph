using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Text;
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
    }

    public class TelegramDataService : ITelegramDataService
    {
        private ApplicationContext Db;
        private Task DbTask;

        public TelegramDataService()
        {
            Db = new ApplicationContext();
        }

        public void LoadTelegrams(
            Action<IEnumerable<Telegram>> success,
            Action<Exception> fail)
        {
            try
            {
                Db.Telegrams.Load();
                success(Db.Telegrams.Local.ToBindingList());
            }
            catch(Exception e)
            {
                fail(e);
            }
        }
    }
}

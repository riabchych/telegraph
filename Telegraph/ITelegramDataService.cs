using System.ComponentModel;

namespace Telegraph
{
    public interface ITelegramDataService
    {
        BindingList<Telegram> LoadTelegrams();

        bool AddTelegram(Telegram tlg);

        bool EditTelegram(int id, Telegram tlg);

        bool RemoveTelegram(object tlg);
    }
}

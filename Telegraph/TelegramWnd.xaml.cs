using System.Windows;

namespace Telegraph
{
    /// <summary>
    /// Логика взаимодействия для TelegramWnd.xaml
    /// </summary>
    public partial class TelegramWnd : Window
    {
        public Telegram Telegram { get; set; }
        
        public TelegramWnd(Telegram t)
        {
            InitializeComponent();
            Telegram = t;
            DataContext = Telegram;
            VM = ApplicationViewModel.SharedViewModel();
        }
    }
}

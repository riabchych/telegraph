using System.Windows;
using Telegraph.ViewModels;
using GalaSoft.MvvmLight;
using GalaSoft.MvvmLight.Command;
using GalaSoft.MvvmLight.Messaging;


namespace Telegraph
{
    /// <summary>
    /// Логика взаимодействия для ImportWnd.xaml
    /// </summary>
    public partial class ImportWnd : Window
    {
        public ImportWnd()
        {
            InitializeComponent();
            Messenger.Default.Register<ImportWnd>(this, (msg) => Dispatcher.Invoke(() => Activate()));
        }
    }
}

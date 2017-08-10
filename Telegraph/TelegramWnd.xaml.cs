using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

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
            this.DataContext = Telegram;
        }

        private void Accept_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = true;
        }
    }
}

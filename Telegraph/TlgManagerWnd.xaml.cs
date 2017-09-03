using System.ComponentModel;
using System.Threading;
using System.Windows;

namespace Telegraph
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class TlgManagerWnd : System.Windows.Window
    {

        public TlgManagerWnd()
        {
            InitializeComponent();
        }

        private void telegramsList_DragLeave(object sender, DragEventArgs e)
        {
            this.telegramsList.BorderThickness = new Thickness(0);
        }

        private void telegramsList_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop, true))
            {
                this.telegramsList.BorderThickness = new Thickness(3);
            }
        }

        private void telegramsList_Drop(object sender, DragEventArgs e)
        {
            this.telegramsList.BorderThickness = new Thickness(0);
        }
    }
}

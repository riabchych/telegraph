using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Text;
using System.Threading;
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
    /// Логика взаимодействия для SplashWnd.xaml
    /// </summary>
    public partial class SplashWnd : Window
    {
        private ApplicationViewModel appViewModel;

        public SplashWnd()
        {
            InitializeComponent();
            InitViewModel();
        }

        private void InitViewModel(object state = null)
        {
            if (appViewModel == null)
            {
                appViewModel = new ApplicationViewModel(SynchronizationContext.Current,
                    new SendOrPostCallback(InitViewModel));
            }
            else
            {
                new TlgManagerWnd(appViewModel).Show();
                Close();
            }
        }
    }
}

using System;
using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Navigation;
using System.IO;
using System.Text.RegularExpressions;
using IFilterTextReader;
using System.Windows.Interactivity;

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

        public TlgManagerWnd(ApplicationViewModel model)
        {
            InitializeComponent();
            this.telegramsList.DataContext = model;
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

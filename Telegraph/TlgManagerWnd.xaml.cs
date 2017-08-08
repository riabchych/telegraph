using System;
using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Navigation;
using System.IO;
using System.Text.RegularExpressions;
using IFilterTextReader;

namespace Telegraph
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class TlgManagerWnd : System.Windows.Window
    {
        struct TlgRegex
        {
            private const string number = @".З[\w\W]*.НР[\W]*([0-9]*)[\W]*[\n]";
            private const string to = @"(КОМАН[ЙЦУКЕНГШЩЗХЇЄЖДЛОРПАВІФЯЧСМИТЬБЮ]+[\w\d\W].*[\n\r])";
            private const string text = @"КОМАН[ЙЦУКЕНГШЩЗХЇЄЖДЛОРПАВІФЯЧСМИТЬБЮ]+[\w\d\W].*[\n\r].?([ЙЦУКЕНГШЩЗХЇЄЖДЛОРПАВІФЯЧСМИТЬБЮ\w\d\W\n\r]*)НР";
            private const string urgency = @"[\s\t\d]{10,}....(ТЕРМІНОВ.)";
            private const string all = @".З[\w\W]*.НР[\W]*([0-9]*)[\W\n\r]*КОМАН[ЙЦУКЕНГШЩЗХЇЄЖДЛОРПАВІФЯЧСМИТЬБЮ\s\/]+([\d\s].*)[\n\r].?([ЙЦУКЕНГШЩЗХЇЄЖДЛОРПАВІФЯЧСМИТЬБЮ\w\d\W\n\r]*)НР\.([\d\/\-]*)([ЙЦУКЕНГШЩЗХЇЄЖДЛОРПАВІФЯЧСМИТЬБЮ\s]*)?([\d\.]+)";

            public static string Number => number;
            public static string To => to;
            public static string Text => text;
            public static string Urgency => urgency;
            public static string All => all;
        }

        private ApplicationContext db = new ApplicationContext();
        NavigationService nav;

        public ObservableCollection<string> Files
        {
            get
            {
                return _files;
            }
        }
        private ObservableCollection<string> _files = new ObservableCollection<string>();

        public ObservableCollection<Telegram> myList { get; private set; }

        public TlgManagerWnd()
        {
            InitializeComponent();
            nav = NavigationService.GetNavigationService(this);
        }

        public TlgManagerWnd(ApplicationViewModel model)
        {
            InitializeComponent();
            nav = NavigationService.GetNavigationService(this);
            this.telegramsList.DataContext = model;
        }

        // To search and replace content in a document part.
        public void SearchAndReplace(string document)
        {
            try
            {
                TextReader reader = new FilterReader(document);

                using (reader)
                {
                    string docText = null;

                    docText = reader.ReadToEnd();

                    string number = new Regex(TlgRegex.Number).Match(docText).Groups[1].Value;
                    string to = new Regex(TlgRegex.To).Match(docText).Groups[1].Value;
                    string text = new Regex(TlgRegex.Text).Match(docText).Groups[1].Value;
                    string urgency = new Regex(TlgRegex.Urgency).Match(docText).Groups[1].Value;

                    MessageBox.Show(number + " : " + text);

                    /*using (StreamWriter sw = new StreamWriter(wordDoc.MainDocumentPart.GetStream(FileMode.Create)))
                    {
                        sw.Write(docText);
                    }*/
                }
            }
            catch(Exception e)
            {
                MessageBox.Show(e.Message, "Error");
            }
        }

        private void telegramsList_DragLeave(object sender, DragEventArgs e)
        {
            this.telegramsList.BorderThickness = new Thickness(0);
        }

        private void telegramsList_Drop(object sender, DragEventArgs e)
        {
            bool isValidFile = false;

            if (e.Data.GetDataPresent(DataFormats.FileDrop, true))
            {
                this.telegramsList.BorderThickness = new Thickness(3);
                _files.Clear();
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);

                foreach (string filePath in files)
                {
                    _files.Add(filePath);
                    SearchAndReplace(filePath);
                }
            }
            if (isValidFile)
                e.Effects = DragDropEffects.Move;
            else
                e.Effects = DragDropEffects.None;

            this.telegramsList.BorderThickness = new Thickness(0);
        }

        private void telegramsList_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop, true))
            {
                this.telegramsList.BorderThickness = new Thickness(3);

            }
        }
    }
}

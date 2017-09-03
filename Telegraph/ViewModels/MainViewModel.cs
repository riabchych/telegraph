using GalaSoft.MvvmLight;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows.Data;

namespace Telegraph.ViewModel
{
    public class MainViewModel : PropertyChangedNotification
    {
        protected struct TlgRegex
        {
            private const string firstPartRegex = @"^���[\s\.]+([�-ߨ����\1]+){1}[\s]+([0-9\-\:]+){1}[\s]*$[\s\-]+^���[\s\.]+([�-ߨ����\1]+){1}";
            private const string secondPartRegex = @"^���[\s\.]+([[�-ߨ����\1]+)[\s]*$[\s\r\n]+^���[\s\.]+([[�-ߨ����\1]+){1}[\s]+([0-9\-\:]+)[\s]?";
            private const string basePartRegex = @"[�-ߨ����\s\d\W]+^([�-ߨ����\s\d]*).��[\s\.]+([0-9]*)?[\s]?$[\s]+^([�-ߨ����\d\/\s]+)[\s\n\r]+^([�-ߨ����\w\W\s]+)^[\s]*��[\.]?[\s]*([\d\/\s\-]+[\s]*[\d]+)[\s]+([�-ߨ����\s\.\-0-9]+)[\r\n]^([\d]+[\s\.\\\/]+[\d]+[\s\.\\\/]+[\d]+)[\s]+([�-ߨ����\s\-\1]+)$[\s]+([�-ߨ����\-\1]+)?[\s]+([�-ߨ����\s\1\.]+)$[\r\n]?[\s\-]+";
            private const string urgencyRegex = @"[\s\d]{10,}....(���̲���.)";

            public static string FirstPartRegex => firstPartRegex;
            public static string SecondPartRegex => secondPartRegex;
            public static string BasePartRegex => basePartRegex;
            public static string UrgencyRegex => urgencyRegex;
        }
        private Page selectFiles;
        private Page selectImportType;
        private ApplicationContext db;
        private Task dbTask;

        public string GetTempFile(string fileExtension)
        {
            string temp = System.IO.Path.GetTempPath();
            string res = string.Empty;
            while (true)
            {
                res = string.Format("{0}.{1}", Guid.NewGuid().ToString(), fileExtension);
                res = System.IO.Path.Combine(temp, res);
                if (!System.IO.File.Exists(res))
                {
                    try
                    {
                        System.IO.FileStream s = System.IO.File.Create(res);
                        s.Close();
                        break;
                    }
                    catch (Exception)
                    {

                    }
                }
            }
            return res;
        }

        public IEnumerable<Telegram> Telegrams
        {
            get { return GetValue(() => Telegrams); }
            set { SetValue(() => Telegrams, value); }
        }

        public ObservableCollection<Telegram> TelegramsCollection
        {
            get { return GetValue(() => TelegramsCollection); }
            set { SetValue(() => TelegramsCollection, value); }
        }

        public CollectionViewSource TelegramsViewSource
        {
            get { return GetValue(() => TelegramsViewSource); }
            set { SetValue(() => TelegramsViewSource, value); }
        }

        public ICollectionView SourceCollection
        {
            get { return TelegramsViewSource.View; }
        }

        public Page CurrentPage
        {
            get { return GetValue(() => CurrentPage); }
            set { SetValue(() => CurrentPage, value); }
        }

        public bool IsBusy
        {
            get { return GetValue(() => IsBusy); }
            set { SetValue(() => IsBusy, value); }
        }

        public Task DbTask { get => dbTask; set => dbTask = value; }
        public ApplicationContext Db { get => db; set => db = value; }
        public Page SelectFiles { get => selectFiles; set => selectFiles = value; }
        public Page SelectImportType { get => selectImportType; set => selectImportType = value; }
    }
}
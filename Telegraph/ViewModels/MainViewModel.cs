using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows.Data;

namespace Telegraph.ViewModels
{
    public abstract class MainViewModel : PropertyChangedNotification
    {
        protected struct TlgRegex
        {
            private const string firstPartRegex = @"^ÂÈÊ[\s\.]+([À-ß¨¯²ª¥\1]+){1}[\s]+([0-9\-\:]+){1}[\s]*$[\s\-]+^ÏÐÄ[\s\.]+([À-ß¨¯²ª¥\1]+){1}";
            private const string secondPartRegex = @"^ÏÐÄ[\s\.]+([[À-ß¨¯²ª¥\1]+)[\s]*$[\s\r\n]+^ÂÈÊ[\s\.]+([[À-ß¨¯²ª¥\1]+){1}[\s]+([0-9\-\:]+)[\s]?";
            private const string basePartRegex = @"[À-ß¨¯²ª¥\s\d\W]+^([À-ß¨¯²ª¥\s\d]*).ÍÐ[\s\.]+([0-9]*)?[\s]?$[\s]+^([À-ß¨¯²ª¥\d\/\s]+)[\s\n\r]+^([À-ß¨¯²ª¥\w\W\s]+)^[\s]*ÍÐ[\.]?[\s]*([\d\/\s\-]+[\s]*[\d]+)[\s]+([À-ß¨¯²ª¥\s\.\-0-9]+)[\r\n]^([\d]+[\s\.\\\/]+[\d]+[\s\.\\\/]+[\d]+)[\s]+([À-ß¨¯²ª¥\s\-\1]+)$[\s]+([À-ß¨¯²ª¥\-\1]+)?[\s]+([À-ß¨¯²ª¥\s\1\.]+)$[\r\n]?[\s\-]+";
            private const string urgencyRegex = @"[\s\d]{10,}....(ÒÅÐÌ²ÍÎÂ.)";

            public static string FirstPartRegex => firstPartRegex;
            public static string SecondPartRegex => secondPartRegex;
            public static string BasePartRegex => basePartRegex;
            public static string UrgencyRegex => urgencyRegex;
        }
        private Page selectFiles;
        private Page selectImportType;
        private ApplicationContext db;
        private Task dbTask;

        protected string GetTempFile(string fileExtension)
        {
            string res = string.Empty;

            while (true)
            {
                res = System.IO.Path.ChangeExtension(System.IO.Path.GetTempFileName(), fileExtension);
                if (!System.IO.File.Exists(res))
                {
                    try
                    {
                        System.IO.File.Create(res).Close();
                        break;
                    }
                    catch (Exception)
                    {

                    }
                }
            }
            return res;
        }

        protected bool CheckData(Telegram data) => ((data.IncNum < 1) ||
                string.IsNullOrWhiteSpace(data.From) ||
                string.IsNullOrWhiteSpace(data.To) ||
                string.IsNullOrWhiteSpace(data.Text) ||
                string.IsNullOrWhiteSpace(data.SubNum) ||
                string.IsNullOrWhiteSpace(data.Date) ||
                string.IsNullOrWhiteSpace(data.SenderPos) ||
                string.IsNullOrWhiteSpace(data.SenderRank) ||
                string.IsNullOrWhiteSpace(data.SenderName) ||
                string.IsNullOrWhiteSpace(data.Executor) ||
                string.IsNullOrWhiteSpace(data.Phone) ||
                string.IsNullOrWhiteSpace(data.HandedBy))
            ? false : true;

        public IEnumerable<Telegram> Telegrams
        {
            get { return GetValue(() => Telegrams); }
            set { SetValue(() => Telegrams, value); }
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

        protected Task DbTask { get => dbTask; set => dbTask = value; }
        protected ApplicationContext Db { get => db; set => db = value; }
        protected Page SelectFiles { get => selectFiles; set => selectFiles = value; }
        protected Page SelectImportType { get => selectImportType; set => selectImportType = value; }
    }
}